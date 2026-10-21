# m365tools

A Linux / WSL toolkit for Microsoft 365 administration, driven from the terminal (`m365` CLI plus a Textual TUI).
It talks plain HTTPS to SharePoint REST and Microsoft Graph: no PnP, no CSOM, no PowerShell.

The first tool is **`spmig`**: migrate a SharePoint document library (or a folder of it) to another library or
tenant, keeping folder structure, original dates and authors, with audit, dry run, verification and an optional
read-only lock of the source.

## Install

Requires Python 3.12+ and [uv](https://docs.astral.sh/uv/).

```bash
uv tool install .                 # installs the `m365` command
# or, without installing:
uvx --from . m365 --help
```

WSL: sign-in opens a browser. Install `wslu` (`sudo apt install wslu`) so `wslview` can open your Windows browser, or set
`BROWSER`. With no browser available a device code is shown instead.

## Tenant setup

Each tenant needs an Entra app registration (public client, no secret):

- Redirect URI `http://localhost` (platform "Mobile and desktop applications"); allow public client flows.
- Delegated permissions: SharePoint `AllSites.FullControl`, Graph `Sites.Read.All`, `User.Read`; grant admin consent.

`m365 tenant setup-help` prints step-by-step instructions (including reusing an existing PnP app registration).

```bash
m365 tenant add contoso --tenant-id contoso.onmicrosoft.com --client-id <app-id> --host contoso.sharepoint.com
m365 tenant login contoso         # browser sign-in; tokens cached in ~/.cache/m365tools (mode 600)
m365 tenant list
```

## Quick start

Define a profile in `~/.config/m365tools/config.toml` (tenants are added by `m365 tenant add`):

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

Optional keys: `overwrite` (`skip_if_same_or_newer` default, `always`, `never`), `restore_metadata`, `workers` (1-16, default 2),
`lock_source_after`, `usermap_path`. A destination without `library` uses the site's default document library.

```bash
m365 profile show servicemanagement
m365 spmig audit    --profile servicemanagement   # read-only checks + findings.csv
m365 spmig dryrun   --profile servicemanagement   # shows what would be copied
m365 spmig migrate  --profile servicemanagement   # audit first; refuses on blocking findings
m365 spmig verify   --profile servicemanagement   # compares source and destination
m365 spmig lock     --profile servicemanagement   # preview; add --apply to make the source read-only
m365 spmig tui                                    # interactive version of all of the above
m365 spmig sites contoso                          # search sites by name
```

Always run audit and dryrun first, and open `report.html` in the run folder.

## What is and is not migrated

Migrated: folders and files (latest version), Created/Modified dates, Created By / Modified By (when the user exists at
the destination or is mapped, see below), large files (chunked upload), apostrophes and other awkward names.

Not migrated: version history (only the current version), permissions on folders and files (the audit counts items with
unique permissions), custom columns and their values, check-out state, OneNote notebooks (copied as files, warned about),
sharing links, workflows and flows. Audit reports each of these.

## Resume and stopping

Every finished file is written to `journal.jsonl` after its metadata step. To continue an interrupted migration pass the
same run folder (`m365 spmig migrate --profile X --resume <run folder>`); done files are skipped, the one in flight is copied
again. Without `--resume` a new run folder is created and files that are already current at the destination are skipped by
the `overwrite` policy. Ctrl+C is graceful: running files finish, nothing new starts, the summary says "stopped early".

## Cross-tenant: user map

Authors and editors can only be restored if they exist at the destination. The audit writes a proposal `usermap.csv`
(`source_email;target_email;display_name;status`, `;`-delimited, Excel friendly). Fill in `target_email`, set
`status` to `mapped`, and give it to the next steps with `usermap_path` in the profile or `--usermap path.csv`.
Unmapped or unknown users do not fail files: dates are still restored and a metadata warning is counted.

## Run folder

`~/.local/share/m365tools/runs/<profile>/<YYYYmmdd-HHMMSS>-<mode>/` (override with `M365TOOLS_DATA_DIR`):

| file | written by |
|---|---|
| `findings.csv`, `usermap.csv` | audit, migrate |
| `plan.json` | audit, dryrun, migrate |
| `journal.jsonl`, `status.json` | migrate (live progress, resume data) |
| `results.csv` | dryrun, migrate (OK / WARN / FAILED / SKIPPED per file) |
| `verify.csv` | verify |
| `lock.csv` | lock, migrate with `lock_source_after` |
| `report.html`, `report.txt` | every mode |
| `run.log`, `tmp/` | log; temporary downloads (removed on exit) |

## Troubleshooting

- **"throttled, waiting Ns"**: SharePoint throttles uploads (often 0.5-1 MB/s per stream). The client honours
  `Retry-After` and retries; just wait. More workers rarely help.
- **Deleted users** (former employees): metadata falls back to dates only and counts a warning; map them to a current user.
- **Size differs for Office files**: SharePoint rewrites Office documents on upload, so verify ignores size for them.
- **Ctrl+C**: safe; resume as described above.
- **Blocking findings** (path > 400 chars, illegal names, no write access, not enough quota): fix them, or accept the
  consequences with `migrate --force`.
- **Wrong library name**: the error lists the available libraries; the default library name is language dependent.
- **Browser does not open on WSL**: install `wslu`, or sign in with the device code shown in the terminal.
- **First run against a new tenant**: see [docs/REAL-TENANT-CHECKLIST.md](docs/REAL-TENANT-CHECKLIST.md) and run
  `uv run python scripts/smoke_real_tenant.py --tenant NAME --site URL` (read-only).

## Development

```bash
uv run pytest
uv run ruff check
```

Design and module contracts: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md). Changes: [CHANGELOG.md](CHANGELOG.md).
