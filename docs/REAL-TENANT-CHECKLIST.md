# Real-tenant checklist

The SharePoint REST layer (`src/m365tools/core/sharepoint.py`) was written against documented behaviour and
mocked responses. These assumptions must be verified on the first run against a real tenant, ideally on a scratch
site. `uv run python scripts/smoke_real_tenant.py --tenant NAME --site URL` exercises the read-only ones
(marked **[smoke]**); the write ones need a throw-away library and an `m365 spmig migrate` of a few files.

`m365` below means the CLI of the Microsoft 365 CLI project (`m365 request`) or plain `curl` with a bearer token;
use whichever you have. Paths are relative to `https://<tenant>.sharepoint.com/sites/<site>`.

1. **ISO date format in ValidateUpdateListItem** (`format_sp_datetime`).
   Check: migrate one file whose source Created/Modified are known; compare dates at the destination.
   How: `m365 request --url '<site>/_api/web/lists(guid'<id>')/items(<n>)?$select=Created,Modified'`.
   If wrong (day/month swapped, shifted by the time zone, or the call returns `HasException`): the value is parsed per the
   web's regional settings. Pass `date_format=` to `set_item_metadata` in `core/sharepoint.py` (e.g.
   `lambda d: d.astimezone(web_tz).strftime("%d-%m-%Y %H:%M")`) and wire the web's locale in through
   `RestGateway.set_item_metadata` in `tools/spmig/gateway.py`.

2. **`POST /_api/web/folders` with `{"ServerRelativeUrl": ...}` under `odata=nometadata`** (`_create_folder`).
   Check: create a folder with a nested new path under a library. Expect 200/201.
   How: `curl -X POST -H 'Accept: application/json;odata=nometadata' -H 'Content-Type: application/json;odata=nometadata' -d '{"ServerRelativeUrl":"/sites/x/Shared Documents/t1"}' <site>/_api/web/folders`.
   If 400 "A type name is required" / missing `__metadata`: send `{"__metadata":{"type":"SP.Folder"},"ServerRelativeUrl":...}`
   with `odata=verbose` headers in `_create_folder` (`core/sharepoint.py`), or switch to `GetFolderByServerRelativePath(...)/AddSubFolderUsingPath`.
   Paths containing `#` or `%` go through `POST /_api/web/folders/AddUsingPath(DecodedUrl='...',Overwrite=false)`
   instead (`_post_new_folder`); create a folder named `#1 100%` once to confirm that call (and its "already exists" error).

3. **"Already exists" error detection** (`_create_folder`).
   Check: POST the same folder twice and read the error. SharePoint normally answers 200 and returns the existing folder;
   some paths give an error containing "already exists" or code `-2130245363`.
   How: repeat the curl from item 2 and inspect the body.
   If the text/code differ (localised message, other code): extend the `"already exists" in ... or "-2130245363" in ...`
   test in `_create_folder` (`core/sharepoint.py`). Until then a real race fails the folder, not the run.

4. **`GetFolderByServerRelativePath(...)?$select=Exists` on a missing folder** (`_folder_exists`).
   Check: does a missing folder give 404 (handled via `NotFoundError`) or 200 with `"Exists": false`?
   How: `m365 request --url "<site>/_api/web/GetFolderByServerRelativePath(decodedurl='/sites/x/Shared%20Documents/nope')?\$select=Exists"`.
   Both are handled; if you see 500 or 400 instead, map that status in `_folder_exists` (`core/sharepoint.py`) or in `SpClient.request` (`core/http.py`).

5. **`$orderby=ID` keyset paging under the 5000 list-view threshold** (`_iter_keyset`) **[smoke]**.
   Check: on a library with more than 5000 items, `$filter=ID gt N&$orderby=ID&$top=2000` must succeed on every page and the
   item count must equal `ItemCount`. Also check that nothing is skipped or duplicated.
   How: `.../items?$select=Id&$filter=ID gt 4990&$orderby=ID&$top=2000`.
   If you get "threshold" errors: lower `page_size`, or fall back to the folder-by-folder `GetItems` CAML approach in `iter_items`.
   `iter_unique_permission_ids` uses `page_size=5000`; lower it in `core/sharepoint.py` if that page is rejected.

6. **Selectable fields: `File_x0020_Size`, `ProgId`, `OData__UIVersionString`, `Author/EMail` expand** (`DEFAULT_ITEM_FIELDS`, `build_item`) **[smoke]**.
   Check: the first 5 items come back with size, version string, author email and a ProgId (may be null for plain files).
   How: `.../items?$select=Id,File_x0020_Size,ProgId,OData__UIVersionString,Author/EMail,Editor/EMail&$expand=Author,Editor&$top=5`.
   If `_UIVersionString` must be requested as `_UIVersionString`, or the email key is `Email` not `EMail`: adjust
   `DEFAULT_ITEM_FIELDS` and `_user()` / `build_item()` in `core/sharepoint.py`. A size of 0 for every file means
   `File_x0020_Size` was not returned; use `File/Length` (expand `File`) as `build_item` already half-supports.

7. **OneNote detection: `ProgId` vs `HTML_x0020_File_x0020_Type`** (`checks.check_onenote`, `Item.prog_id`).
   Check: a OneNote notebook folder should have `ProgId = "OneNote.Notebook"`; some tenants expose it only as
   `HTML_x0020_File_x0020_Type = "OneNote.Notebook"`.
   How: `.../items?$select=Id,FileLeafRef,ProgId,HTML_x0020_File_x0020_Type&$filter=FSObjType eq 1`.
   If `ProgId` is empty for notebooks: add `HTML_x0020_File_x0020_Type` to `DEFAULT_ITEM_FIELDS` and fall back to it when
   setting `prog_id` in `build_item` (`core/sharepoint.py`).

8. **`site/Usage` `StoragePercentageUsed`: fraction or percent** (`get_site_usage`) **[smoke]**.
   Check: compare `Storage / StoragePercentageUsed` with the quota shown in the admin center.
   How: `m365 request --url '<site>/_api/site/Usage'`; a 10 GB site on a 100 GB quota should show `0.1` (fraction).
   If it is a percentage (10.0): divide by 100 in `get_site_usage` (`core/sharepoint.py`). If `Usage` is nested differently, fix the
   `data.get("Usage", data)` line. When the quota cannot be derived the quota check only reports INFO (`tools/spmig/checks.py`).

9. **`BasePermissions` `Low`/`High` types** (`_low_bits`, `get_role_assignments`) **[smoke]**.
   Check: in `RoleDefinitionBindings`, `BasePermissions.Low` may be a number or a string (`"2147483647"`); `_int` copes with both.
   How: `.../roleassignments?$expand=Member,RoleDefinitionBindings`; confirm Full Control / Edit / Contribute give `can_write` true and Read / Limited Access false.
   If `BasePermissions` is missing from the expansion: add `RoleDefinitionBindings/BasePermissions` to `$select`, or query `roledefinitions` once and join by Id (`core/sharepoint.py`).

10. **`roledefinitions` `RoleTypeKind` 2 = Reader** (`_reader_role_id`).
    Check: exactly one definition with `RoleTypeKind 2` per site, named "Read"/"Lezen" depending on language.
    How: `m365 request --url "<site>/_api/web/roledefinitions?\$select=Id,Name,RoleTypeKind"`.
    If missing (custom or deleted Read role): `set_reader_only` raises NotFoundError and the lock stops safely; fix by restoring the role or matching by name in `_reader_role_id`.

11. **Search response shape** (`search_sites`).
    Check: `/_api/search/query` with `nometadata` returns `PrimaryQueryResult.RelevantResults.Table.Rows[].Cells[]` as `{Key,Value}` lists; in some tenants rows are wrapped in `results`.
    How: `curl -H 'Accept: application/json;odata=nometadata' "<root>/_api/search/query?querytext='contentclass:STS_Site test*'&rowlimit=5&selectproperties='Title,Path'"`.
    If empty because the search index is stale or the response differs: adjust the traversal in `search_sites` (`core/sharepoint.py`); the Graph fallback needs a `graph_client` (`m365 spmig sites` should pass one).

12. **`ensureuser` `logonName`** (`ensure_user`, used only with `resolve_user(..., ensure=True)`).
    Check: `POST /_api/web/ensureuser {"logonName": "user@domain"}` resolves a plain email; some tenants need the claims form `i:0#.f|membership|user@domain`.
    How: `curl -X POST -H 'Content-Type: application/json;odata=nometadata' -d '{"logonName":"i:0#.f|membership|user@domain"}' <site>/_api/web/ensureuser` (write: adds the user to the site).
    If only the claims form works: change the body in `ensure_user` to `_claims`-style login (`core/sharepoint.py`). Also confirm that `ValidateUpdateListItem` accepts `[{"Key":"i:0#.f|membership|<email>"}]` for Author/Editor (`_claims`) and rejects deleted users with `HasException`/`ErrorMessage` (otherwise `UserNotFoundError` is never raised and the whole metadata call fails).

13. **Probe delete via `X-HTTP-Method: DELETE`** (`probe_write_access`).
    Check: after an audit no `__m365_probe_*` folder remains in the destination.
    How: list folders of the audited folder: `.../GetFolderByServerRelativePath(decodedurl='<folder>')/Folders?$filter=startswith(Name,'__m365_probe')`.
    If the folder stays (400/405 on the tunnelled DELETE): use `DELETE` as a real method via `client.request("DELETE", ...)` with `IF-MATCH: *` in `probe_write_access` (`core/sharepoint.py`), and note the extra folder in the audit hint.

14. **URL encoding of apostrophes, spaces, `#`, `%`** (`_lit`, `_url`) **[smoke]**.
    Check: download, upload and folder-create for names such as `DAP's plan #1 100%.docx` and a folder `0. Archief`.
    How: create these on a scratch library, then `m365 spmig dryrun` + `migrate` and verify.
    If requests 400: check that apostrophes are doubled **before** percent-encoding (`_lit`), that `#` and `%` are encoded (`quote(safe="/")`), and that query strings use `%20` not `+` (`_url`), all in `core/sharepoint.py`.

15. **Upload of files > 10 MB and exactly 10 MB** (`upload_file`).
    Check: `StartUpload/ContinueUpload/FinishUpload` with `uploadId=guid'...'` and `fileOffset` on a 25 MB file; sizes match at the destination.
    How: migrate a 25 MB file, `verify` reports ok.
    If 400 on `ContinueUpload`: the offset or guid literal format differs (`core/sharepoint.py`, `upload_file`). `AddUsingPath` with an empty body then chunking must produce a single version.

16. **`ValidateUpdateListItem` with `bNewDocumentUpdate=true` keeps one version** (`set_item_metadata`).
    Check: after migrating a versioned library the destination file shows version 1.0/1 (not 2.0), and Modified/Editor are the source values.
    How: `.../items(<id>)?$select=OData__UIVersionString,Modified,Editor/EMail&$expand=Editor`.
    If a new version appears or Modified is overwritten by "now": drop `bNewDocumentUpdate` or write Modified last (`core/sharepoint.py`).
