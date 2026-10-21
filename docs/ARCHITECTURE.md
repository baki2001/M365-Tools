# m365tools — architecture contract

Linux/WSL-only Python 3.12 toolkit for Microsoft 365 / SharePoint administration.
First tool: `spmig` (SharePoint library migration). Installed with `uv tool install .`, run as `m365`.

This file is the contract every module is built against. Keep it in sync when you change an interface.

## Principles

1. **Core has no UI.** `m365tools.core` and `m365tools.tools.*` never print or prompt. They raise, return, and
   emit `ProgressEvent`s through a callback. The TUI (`m365tools.tui`) and the CLI (`m365tools.cli`) render.
2. **Plain HTTPS only.** SharePoint REST (`/_api/...`) and Microsoft Graph via `httpx`. No PnP, no CSOM.
3. **Two independent connections** per job (source, destination), each with its own tenant + token. Same tenant is
   just the special case where both resolve to the same `TenantConfig`.
4. **Never widen access.** Any permission change only removes/downgrades; a "lock" that would grant a right is a bug.
5. **Everything resumable.** Long runs write `journal.jsonl` and `status.json`; a restart reads the journal first.
6. **Human-readable first.** Every run ends with `report.html` + `report.txt`; machine output (CSV/JSON) alongside.
7. Not over-engineered: no plugin system, no database, no async framework beyond `asyncio` + `httpx.AsyncClient`.

## Package layout

```
src/m365tools/
  core/
    config.py      TenantConfig, EndpointConfig, RouteConfig, ProfileConfig, AppConfig (pydantic), load_config/save_config,
                   add_tenant, get_tenant/get_profile, tenant_for_url, profile_to_job_options  (TOML, XDG config dir)
    auth.py        TokenProvider(tenant, prompt=None): MSAL PublicClientApplication; silent first, then interactive
                   browser (WSL: wslview/$BROWSER) or device code; per-tenant cache ~/.cache/m365tools (chmod 600);
                   sharepoint_scope(tenant) = https://<host>/.default; GRAPH_SCOPE; tenant_setup_instructions(tenant_id)
    http.py        SpClient (see "HTTP client")
    sharepoint.py  SharePoint REST domain calls (see "SharePoint API surface")
    graph.py       GraphClient(SpClient): search_sites, find_user_by_mail, list_users_by_local_part
    events.py      ProgressEvent dataclasses, EventSink protocol, null_sink, CollectingSink
    journal.py     Journal (append-only fsync'ed jsonl), StatusWriter (atomic status.json, >= 3 s apart)
    report/        model.py (Report, Section), builders.py (audit/dryrun/migrate/verify/lock_report), html.py, text.py,
                   csv_.py, json_.py, write_all(run_dir, report) -> report.html + report.txt
    errors.py      M365Error > ConfigError, AuthError, NotFoundError(available=[...]), AccessDeniedError,
                   ThrottledError(retry_after), TransferError, ValidationError
    logging.py     setup_logging(run_dir) -> run.log (+ SinkHandler bridging log records to the EventSink)
    paths.py       config_dir/cache_dir/data_dir (M365TOOLS_*_DIR or XDG), runs_dir, new_run_dir(name, mode)
  tools/spmig/
    models.py      Item, LibraryRef, Inventory, Plan/PlanEntry/Action, Finding/Severity, Endpoint, Route, JobOptions,
                   UserMapEntry, RunSummary, CompareRow/CompareResult
    gateway.py     SpGateway Protocol, GatewayFactory Protocol, RestGateway   (the only door to SharePoint)
    inventory.py   build_inventory(gw, library, folder="", sink, *, page_size, label) -> Inventory; inventory_base()
    plan.py        build_plan(src, dst, opts, dst_base, route_bases, route_invs) -> Plan; plan_to_json/plan_from_json
    checks.py      run_checks(ctx: CheckContext, sink) -> list[Finding]   (rule functions registered in CHECKS)
    engine.py      Migrator(src_gw, dst_gw, plan, opts, run_dir, sink, journal, status, usermap, *, dry_run, route_targets)
                   .run() -> RunSummary
    verify.py      compare(src_inv, dst_inv) -> CompareResult
    lock.py        preview_lock(gw, lib, include_owners) / apply_lock(gw, lib, changes, sink)
    usermap.py     build_usermap(src_inv, dst_gw), save_usermap/load_usermap, UserMap (resolve(email))
    job.py         Job, RestGatewayFactory, BlockingFindingsError  — the only thing UI calls
  tui/             Textual app (screens: home, source, destination, options, audit, plan, migrate, verify, lock)
  cli.py           typer app `m365`:
                     m365 tenant  add | list | login | setup-help
                     m365 profile list | show
                     m365 spmig   audit | dryrun | migrate | verify | lock | tui | sites
```

Dependency direction: `core` has no knowledge of the CLI/TUI. `core.sharepoint` imports the `Item`/`LibraryRef` models
from `tools.spmig.models`; `core.config` imports `Endpoint`/`JobOptions`. spmig modules other than `gateway.RestGateway`
never import `core.sharepoint`; they see only the `SpGateway` protocol (tests use `tests/fake_gateway.py`).

## Core models (tools/spmig/models.py)

All are `@dataclass(slots=True)`; authoritative field names:

```python
class Severity(StrEnum): BLOCKING="blocking"; WARN="warn"; INFO="info"
class Action(StrEnum): CREATE_FOLDER; COPY; SKIP; CONFLICT; EXCLUDED

Item(id, is_folder, rel_path, name, ext, size, created, modified, author_email, author_name, editor_email,
     editor_name, version, checked_out_to, prog_id, server_relative_url, unique_permissions=False, etag=None)
LibraryRef(site_url, web_server_relative_url, list_id, title, root_server_relative_url, item_count,
           enable_versioning, major_version_limit, enable_minor_versions, force_checkout, enable_moderation)
Inventory(library, base_url, items)          # files(), folders(), total_bytes(), index() -> {lower rel_path: Item}
PlanEntry(action, item, target_server_relative_url, reason="", route=None)
Plan(entries, src, dst, src_base, dst_base)  # counts(), bytes_to_copy() (COPY only), to_copy(), folders_to_create()
Finding(severity, category, path, detail, fix="")
Endpoint(tenant, site_url, library=None, folder="")      # library None => default document library
Route(prefix, destination: Endpoint)
JobOptions(source, destination, overwrite="skip_if_same_or_newer"|"always"|"never", restore_metadata=True,
           excludes=[], routes=[], lock_source_after=False, workers=2, usermap_path=None, profile=None)
UserMapEntry(source_email, target_email, display_name, status="mapped"|"unmapped"|"auto")
RunSummary(mode, started, finished, run_dir, copied, skipped, failed, metadata_warnings, bytes_copied,
           stopped_early, findings)
CompareRow(status="ok"|"missing"|"size_differs"|"source_newer"|"only_in_destination", is_folder, rel_path,
           src_size, dst_size, src_modified, dst_modified, detail)
```

## Events (core/events.py)

`PhaseChanged(phase, detail)`, `ItemStarted(rel_path, size, index, total)`, `ItemStep(rel_path, step)` with step in
download/upload/metadata/folder, `ItemDone(rel_path, ok, seconds, message, skipped)`, `Throttled(seconds, attempt, what)`,
`LogLine(level, text)`, `Stats(done, total, bytes_done, bytes_total, mbps, eta_s, failed, warnings, copied, skipped)`,
`AuthPrompt(tenant, message, url, code)`. `EventSink` is `Callable[[ProgressEvent], None]`; `null_sink` and
`CollectingSink` are provided.

## HTTP client (core/http.py)

`SpClient(token_provider, scope, sink=null_sink, max_attempts=6, base_delay=2.0)`; use as `async with`. Defaults
`Accept: application/json;odata=nometadata`; JSON bodies are sent as `application/json;odata=verbose`.

- `request(method, url, *, json, data, content, headers, params, stream)` returns only 2xx/3xx responses.
  - 429/502/503/504 and transient transport errors (connect/read/write timeouts, `RemoteProtocolError`, ...) are
    retried with exponential backoff (cap 300 s) or `Retry-After` (cap 900 s); each wait emits `Throttled`.
  - Retries exhausted: HTTP status -> `ThrottledError(retry_after=...)`; transport error -> `TransferError`.
  - 401 refreshes the token once, then `AuthError`; 403 `AccessDeniedError`; 404 `NotFoundError`; others `M365Error`.
- `get_json`, `post_json` (empty body -> `{}`), `iter_odata(url)` (follows `odata.nextLink` / `@odata.nextLink` / `__next`).
- `download_to(url, dest)` streams to `dest.part`, renames on success; a body-read failure restarts the download.
- `upload_bytes(url, data: bytes|Path, headers=None, method="POST") -> parsed JSON` (`{}` when empty).
- `request_digest(site_url)` caches `X-RequestDigest` per site (not needed for bearer-token calls; kept for completeness).

## Gateway (tools/spmig/gateway.py)

`SpGateway` is a `runtime_checkable` Protocol bound to one site (`site_url`) and one authenticated client:
`get_default_library`, `get_library(name)`, `iter_items(lib)`, `build_item(raw, lib, base_url)`,
`iter_unique_permission_ids(lib)`, `ensure_folder(path)`, `download_file(url, dest)`,
`upload_file(folder, name, src, overwrite=True) -> UploadResult(unique_id, item_id)`,
`set_item_metadata(lib, item_id, created, modified, author_email, editor_email)`, `get_site_usage() -> (used, quota)`
(quota 0 when unknown), `probe_write_access(folder)` (raises `AccessDeniedError`), `get_role_assignments(lib)`,
`get_associated_owner_group_id()`, `has_unique_role_assignments(lib)`, `break_inheritance(lib, copy=True)`,
`set_reader_only(lib, principal_id)`, `get_fields(lib)`, `find_site_user(email) -> SiteUser | None`.

`GatewayFactory` is `async (Endpoint) -> SpGateway`. `RestGateway(client, site_url)` adapts `core.sharepoint`
functions (imported lazily) and normalises their results into the gateway's own types. The gateway module defines its own
`UserNotFoundError(NotFoundError)`, `RoleAssignment(principal_id, principal_title, principal_type: str, role_names,
can_write)`, `FieldInfo`, `SiteUser`, `UploadResult`; `RestGateway.set_item_metadata` translates
`core.sharepoint.UserNotFoundError` into the gateway one, which is what the engine catches.

`RestGatewayFactory(tenants, sink)` (in `job.py`) builds one `TokenProvider` + `SpClient` per tenant, shared across that
tenant's sites; `aclose()` closes them.

## SharePoint API surface (core/sharepoint.py)

All async; first argument is the client; site-scoped calls take `site_url` as the next argument. Requests use
`odata=nometadata`; parsers also tolerate `{"d": ...}` verbose shapes. Server-relative paths always go through
`GetFileByServerRelativePath(decodedurl=...)` / `GetFolderByServerRelativePath(decodedurl=...)` /
`AddUsingPath(DecodedUrl=...)`; string literals double apostrophes then percent-encode (`quote(safe="/")`).

- `search_sites(client, query, *, root_url, graph_client=None, row_limit=50) -> list[SiteHit]` via
  `/_api/search/query?querytext='contentclass:STS_Site <q>*'`; falls back to Graph `/sites?search=`.
- `get_web(client, site_url) -> WebInfo(server_relative_url, title, language)`.
- `get_default_library(client, site_url)` via `web/DefaultDocumentLibrary?$expand=RootFolder`.
- `get_library(client, site_url, name)`: by Title, then RootFolder name, case-insensitive; `None`/empty = default.
  Raises `NotFoundError(available=[titles])` (BaseTemplate 101, not hidden).
- `iter_items(client, lib, fields=None, *, page_size=2000)`: **keyset paging** on `/lists(guid'..')/items` with
  `$filter=ID gt <last>&$orderby=ID&$top=N` (no CAML/GetItems). User columns are projected with `$expand`
  (Author, Editor, CheckoutUser). Fields: Id, FileLeafRef, FileRef, FileDirRef, FSObjType, File_x0020_Size, Created,
  Modified, Author/EMail, Author/Title, Editor/EMail, Editor/Title, OData__UIVersionString, CheckoutUser/Title,
  CheckoutUser/EMail, ProgId.
- `build_item(row, base_url) -> Item | None` (None for the base folder or anything outside it).
- `iter_unique_permission_ids(client, lib)`: same keyset pager over `Id,HasUniqueRoleAssignments` (page 5000).
- `ensure_folder(client, site_url, folder)`: `GetFolderByServerRelativePath ... $select=Exists`, recursing to the parent,
  then `POST /_api/web/folders {"ServerRelativeUrl": ...}` (`odata=nometadata`); an "already exists" error
  (text or `-2130245363`) counts as success. Never creates a library root.
- `download_file(client, site_url, url, dest)`: `GetFileByServerRelativePath(decodedurl=..)/$value` streamed.
- `upload_file(client, site_url, folder, name, src_path, overwrite=True)`: <= 10 MB a single
  `Folder/Files/AddUsingPath(DecodedUrl=..,Overwrite=..)` with raw bytes; larger: empty `AddUsingPath` then
  `StartUpload/ContinueUpload/FinishUpload` in 10 MB chunks (`CancelUpload` on failure). Returns
  `{"UniqueId","ServerRelativeUrl","Length","ItemId"}`; `ItemId` comes from `ListItemAllFields?$select=Id`.
- `set_item_metadata(client, lib, item_id, created, modified, author_email, editor_email, *, date_format)` via
  `ValidateUpdateListItem` with `bNewDocumentUpdate=true` (no new version). Dates are ISO-8601 `...Z`
  (`format_sp_datetime`); users as `[{"Key":"i:0#.f|membership|<email>"}]`. Raises `core.sharepoint.UserNotFoundError`
  for a rejected Author/Editor, `M365Error` for other field errors.
- `get_site_usage(client, site_url) -> (used, quota|None)` from `/_api/site/Usage`: quota = `Storage / StoragePercentageUsed`
  (fraction 0..1).
- `probe_write_access(client, site_url, folder) -> bool`: creates then deletes `__m365_probe_<uuid>` (POST with
  `X-HTTP-Method: DELETE`); False on 403.
- `get_role_assignments(client, lib)` (`$expand=Member,RoleDefinitionBindings`; `can_write` = Low bits intersect
  Add 0x2 | Edit 0x4 | Delete 0x8 | ManageLists 0x800), `get_associated_owner_group_id`, `has_unique_role_assignments`,
  `break_inheritance(client, lib, copy_role_assignments=True, clear_subscopes=False)`.
- `set_reader_only(client, lib, principal_id)`: requires unique permissions and a currently write-capable principal
  (else `ValidationError`, never widens). **Adds Reader (RoleTypeKind 2) first, then removes every other binding**, so
  there is no moment without access.
- `get_fields(client, lib)`: custom columns (not hidden, not FromBaseType, CanBeDeleted).
- `find_site_user(client, site_url, email)`: read-only `/web/siteusers?$filter=Email eq '..'`; `ensure_user`
  (POST `/web/ensureuser {"logonName": email}`, a write) and `resolve_user(..., ensure=False)`.

The unverified assumptions in this module are listed in [REAL-TENANT-CHECKLIST.md](REAL-TENANT-CHECKLIST.md).

## Auth (core/auth.py)

- `TokenProvider(tenant, prompt=None)`; `await provider.token(scope, force_refresh=False)`. SharePoint scope is
  `https://<sharepoint_host>/.default`, Graph `https://graph.microsoft.com/.default`. Silent acquisition first;
  interactive browser if possible (`can_open_browser()`; WSL uses `wslview`/`$BROWSER`), else device code, surfaced as an
  `AuthPrompt` through the callback.
- `tenant_setup_instructions(tenant_id)` returns the Entra app registration text (public client, redirect
  `http://localhost`, delegated SharePoint `AllSites.FullControl`, Graph `Sites.Read.All`, `User.Read`).

## Planning rules (tools/spmig/plan.py)

Entries are ordered folders (shallowest first) then files. Per item, in this order:

1. **Routes win over excludes.** If `rel_path` equals or starts with a route prefix (longest prefix wins) it is
   targeted under that route's destination base (`entry.route = prefix`). Exclude patterns that address only the route
   prefix itself (e.g. `"0. Archief/**"` for route `0. Archief`) are ignored for routed items; other patterns (also
   `"0. Archief/Sub/**"`) still apply. A route whose destination could not be resolved never falls back to the main
   destination: its items are `EXCLUDED` ("route destination unavailable") and the audit reports a BLOCKING finding.
2. **Excludes**: any glob in `opts.excludes` -> `EXCLUDED`.
3. Folder: exists at destination -> `SKIP`, else `CREATE_FOLDER` (a route root folder is `SKIP`).
4. File: missing -> `COPY`; exists -> `never` = `SKIP`, `always` = `COPY`, `skip_if_same_or_newer` = `SKIP` when
   `dst.modified >= src.modified - 2 s`, else `CONFLICT` (a folder where a file should go is also `CONFLICT`; so is an empty
   destination file for a non-empty source, the leftover of an interrupted chunked upload). Naive datetimes count as UTC.
   The engine copies `CONFLICT` entries; they are reported as warnings.

Globs are **case-insensitive** on the `/`-separated `rel_path`; `*` and `**` both cross `/`; a leading `**/` also
matches at the root; `X/**` also matches the folder `X` itself. All path comparisons elsewhere (index, journal) are
case-insensitive too.

## Checks (tools/spmig/checks.py) — each `(ctx: CheckContext) -> list[Finding]`

| rule (category) | severity |
|---|---|
| destination library resolvable (`destination`) | BLOCKING |
| write access probe on destination folder (`write_access`) | BLOCKING |
| quota: dst free >= src total * 1.1 (`quota`; INFO when usage/quota is unknown) | BLOCKING |
| target path > 400 chars (`path_length`), illegal chars / leading-trailing space or dot / `~$` names (`illegal_name`) | BLOCKING per item |
| unique permissions (`unique_permissions`) | WARN |
| checked-out files (`checked_out`) | WARN per item |
| version history (`versions`; only latest copied) | INFO |
| custom columns missing at destination, `SYSTEM_FIELDS` ignored (`columns`) | WARN |
| authors/editors not resolvable at destination via usermap or siteusers (`users`) | WARN per user |
| conflicts (`conflict`) | WARN |
| OneNote notebooks (`onenote`) | WARN |
| files >= 100 MB (`large_file`) | INFO |
| library settings differ (`library_settings`) | INFO |
| estimated duration (`duration`) | INFO |

A check that raises is reported as a WARN finding `check_failed` instead of aborting the audit.

## Engine (tools/spmig/engine.py)

`Migrator(src_gw, dst_gw, plan, opts, run_dir, sink=null_sink, journal=None, status=None, usermap=None, *, dry_run=False,
route_targets=None, retries=2, retry_pause=5.0, ...)`. The **plan is a constructor argument**; `await migrator.run()`
takes no arguments and returns a `RunSummary`. `route_targets` maps route prefix -> `(gateway, LibraryRef)`.

- Phases: `folders` (CREATE_FOLDER entries, shallowest first) -> `files` (`opts.workers` concurrent via a semaphore)
  -> `finalize` (tmp cleanup, `results.csv`, final status).
- Per file: download to `run_dir/tmp/<uuid>` -> `upload_file(overwrite=True)` -> `set_item_metadata` (author/editor mapped
  through the `UserMap`) -> journal line -> delete tmp. The journal is written after the metadata step.
- Metadata: `UserNotFoundError` -> retry with dates only, count a metadata warning (result `WARN`), the file is still OK.
- Resume: `journal.done_paths()` (case-insensitive) are skipped. Files with only failed journal entries
  (`Journal.failed_paths()`) are copied again even when the re-built plan says `SKIP` (an upload whose metadata step
  failed looks "newer" at the destination). Dry run never touches the journal or destination.
- Folder-creation failures count in `failed`. SIGINT is handed back to the previous handler after the run.
- Retry: `SpClient` handles HTTP retries; for non-`M365Error` exceptions the engine retries the whole file `retries` times.
- Watchdog: a file running longer than `max(15 min, 5 x size / max(0.2 MB/s, measured rate))` emits one `LogLine("warn")`.
- Stop: SIGINT (main thread) or `request_stop()` sets an event; in-flight files finish, nothing new starts,
  `RunSummary.stopped_early = True`.
- `status.json` is rewritten at most every 3 s (phase, mode, counts, bytes, mbps, eta_s, current file/step, run_dir).

## Job API (tools/spmig/job.py)

```python
job = Job(gateways, opts, sink=null_sink, run_dir_factory=None, *, journal_factory=..., status_factory=...,
          migrator_kwargs=None)
# gateways: GatewayFactory, or Mapping[str, TenantConfig] (then RestGatewayFactory is used)
# run_dir_factory: (mode) -> Path ; default m365tools.core.paths.new_run_dir(opts.profile or "adhoc", mode)
await job.audit()                                   -> RunSummary           (findings in summary.findings)
await job.dryrun()                                  -> (Plan, RunSummary)
await job.migrate(resume_run_dir=None, force=False) -> RunSummary           (raises BlockingFindingsError)
await job.verify()                                  -> (CompareResult, RunSummary)
await job.lock(apply=False, include_owners=False)   -> list[LockChange]
job.request_stop(); await job.aclose()
```

- `BlockingFindingsError(ValidationError)` carries `.findings`, `.blocking`, `.run_dir`. `migrate` runs the audit first
  and refuses on blocking findings unless `force=True`; it also raises it when the destination cannot be resolved.
- `migrate(resume_run_dir=...)` reuses that folder: the existing `journal.jsonl` makes the engine skip finished files
  and `results.csv` is appended to.
- With `opts.lock_source_after`, the source is locked at the end only when the run finished without failures and was
  not stopped early.
- `lock` previews or applies on the **source** library; the site owners group is kept unless `include_owners=True`.
- After a run `job.state/plan/findings/usermap/run_dir` hold the last results for the UI.
- Only the source library is required to exist; a missing destination (or route destination) is reported as a
  BLOCKING finding and the plan is skipped.

## Run folder

`~/.local/share/m365tools/runs/<profile-or-adhoc>/<YYYYmmdd-HHMMSS>-<mode>/` (`M365TOOLS_DATA_DIR` overrides), always with
an empty-on-success `tmp/`. Contents per mode (`;`-delimited CSVs, UTF-8 with BOM):

| mode | files |
|---|---|
| audit | `findings.csv`, `plan.json` (when destination resolved), `usermap.csv` (proposal; only when metadata restore is on and no usermap was supplied), `report.html`, `report.txt` |
| dryrun | `plan.json`, `results.csv` (DRYRUN rows), `report.html`, `report.txt` |
| migrate | `findings.csv`, `plan.json`, `usermap.csv` (proposal), `journal.jsonl`, `status.json`, `results.csv`, `lock.csv` (only with `lock_source_after`), `report.html`, `report.txt` |
| verify | `plan.json` is NOT written; `verify.csv`, `report.html`, `report.txt` |
| lock | `lock.csv`, `report.html`, `report.txt` |

`run.log` is written by `setup_logging(run_dir)` when the CLI/TUI enables it. `results.csv` columns:
`time;result;path;size_mb;seconds;message` with result OK/WARN/FAILED/SKIPPED/DRYRUN.

## Reports (core/report)

`Report(title, mode, source, destination, started, finished, summary, sections)`; `Section(title, kind, columns, rows, note)`.
HTML is one self-contained file (inline CSS, no JS dependencies); text report is 100 columns max.

## Config (~/.config/m365tools/config.toml)

```toml
[tenants.contoso]
tenant_id = "contoso.onmicrosoft.com"
client_id = "11111111-2222-3333-4444-555555555555"
sharepoint_host = "contoso.sharepoint.com"

[profiles.servicemanagement]
source      = { tenant="contoso", site_url="https://contoso.sharepoint.com/sites/OldDocuments", library="Servicemanagement" }
destination = { tenant="contoso", site_url="https://contoso.sharepoint.com/sites/Service" }
excludes    = ["0. Archief/**"]
routes      = [{ prefix="0. Archief", destination={ tenant="contoso", site_url="https://contoso.sharepoint.com/sites/Service", library="Archief" } }]
```

Optional profile keys: `overwrite` (`skip_if_same_or_newer` | `always` | `never`), `restore_metadata`,
`lock_source_after`, `workers` (1-16), `usermap_path`. Unknown keys are rejected.

## Lessons baked in (from the Servicemanagement migration, Oct 2026)

- Default library name is language-dependent ("Gedeelde documenten"); always resolve via DefaultDocumentLibrary.
- Deleted users (`gone.user@contoso.com`) fail metadata set → fall back to dates only, count as warning, never fail the file.
- Office files change size on upload (SharePoint injects metadata) → verify ignores size for Office extensions.
- Throttling is real at ~0.5–1 MB/s single-stream; show "throttled, waiting Ns" rather than looking stuck.
- Lock must only downgrade principals that hold Edit/Add/Delete/ManageLists; "Limited Access" entries untouched.
- Dry run must not report percent based on bytes (showed 50%); use file count in dry run.
- Ctrl+C mid-file: journal only after metadata step; resume re-copies that one file.

## Testing

- Unit tests with `respx` mocking SharePoint REST; fixtures in `tests/fixtures/` include a trimmed inventory JSON
  modelled on the Servicemanagement library (1210 files / 559 folders counts preserved in summary fixtures).
- `tests/test_plan.py`, `test_checks.py`, `test_engine.py` (resume, Ctrl+C, metadata fallback), `test_lock.py`
  (never-widen), `test_verify.py` (office tolerance), `test_report.py` (HTML renders, text ≤100 cols).
- `uv run ruff check` and `uv run pytest` must pass.
