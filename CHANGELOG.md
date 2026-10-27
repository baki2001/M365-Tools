# Changelog

## Unreleased

- spmig: dates are written the way the destination site reads them. SharePoint parses the Created and Modified
  text with the site's regional settings, so on a Dutch site every ISO date was refused ("U moet een geldige
  datum opgeven binnen het bereik van 1-1-1900 tot 31-12-8900") and every file counted as failed. Dates now go
  out in the site's own format and time zone (`31-1-2024 11:00:05` for nl-NL).
- spmig: a file whose metadata step fails after a good upload counts as copied, with a metadata warning. Resume
  sets the metadata again without uploading the file a second time, also for files that failed this way
  before this change. If the first 25 files all fail their metadata step, the run stops.
- spmig audit: a new `metadata` check sets dates on a test file at the destination and reads them back, so a
  date problem shows up before the copy starts. A `dates` check lists files with dates SharePoint cannot store
  (before 1900 or after 8900); the migration uses the other date for those.
- spmig: the journal records seconds per step (download, upload, metadata) and the run log ends with the share
  of time each step took. Modified dates within a minute count as equal, since some sites keep dates to the
  minute.

- Windows Terminal and WSL: themes now get 24-bit colour. Windows Terminal does not set `COLORTERM`, so
  the TUI fell back to 256 colours and every background turned black. It now sets `COLORTERM=truecolor` when
  `WT_SESSION` is present and `COLORTERM` is empty.
- Colour themes: `ctrl+t` (also in the `:` palette) opens a picker with 16 themes, previewed live as you
  move: Catppuccin Mocha, Macchiato, Frappé and Latte, Dracula, Gruvbox dark and light, Nord, Tokyo Night,
  One Dark, Solarized dark and light, Rosé Pine, Kanagawa, Everforest, and Terminal colours (the terminal's
  own ANSI palette). The choice is saved as `[ui] theme` in config.toml.
- Calmer, more consistent layout. Counts read the same on every screen (`● 3 blocking   ● 1 warn`, zeros
  dimmed) instead of filled pills; active filters sit at the end of that row. Table headers no longer have a
  filled band, and Plan rows lost the `▮` marker. The thick coloured bar down the left of the main pane is
  gone: the cursor and box borders show focus. The idle Run page shows the state and the log instead of eight
  empty stat boxes; during a migration the stats are a compact grid. Plan's blocking warning is one line.
  The Job pipeline uses the sidebar's state glyphs. The header lists all three tools with the open one
  marked. Log lines wrap under the message, not under the time. Counts use proper plurals ("1 run",
  "2 runs") instead of "run(s)".
- Quieter edges. The header keeps the brand, the tools and the signed-in account in grey; tenant and profile
  moved to the bottom of the sidebar. The key bar shows only the keys for the screen you are on, and `?` lists
  the rest (audit, dry run, migrate, verify, `:`, `g` and `q` still work everywhere). Keys are coloured text
  instead of boxed chips. The selected entry in the sidebar, the home tool list, the groups list, onboard runs
  and pickers is marked with a thin accent line on its left instead of a filled bar, and the table cursor is a
  soft grey band. The groups tabs are no longer filled. Plan's totals line is a single grey line.
- spmig Runs (`7`): Enter loads a run back into the workspace. The job gets the source, destination and options
  the run used, Checks gets the audit findings and Plan gets the dry run, so you can continue to migrate after
  restarting the TUI. Loading an audit also brings the newest dry run with the same job, and the other way
  round. `t` now opens report.txt (Enter used to). The recent runs panel on the Job page says how.
- Groups screen reworked for orientation: tabs with their keys (`1` Groups, the landing tab, `2` People,
  `3` Bundles); a "Get started" checklist is the default right pane (`0` or `esc` returns to it) with the
  next step and its key; `D` sets your own domains so only colleagues count; the People list shows a title or
  domain instead of "unassigned"; help has a glyph legend. Tool switching is documented as `g` then `1/2/3`
  (`ctrl+1/2/3` where the terminal sends it).
- The header shows the signed-in account as soon as any operation has a token in the cache, not only after
  `ctrl+l`.
- Onboarding is its own tool: `3 onboard` on the home screen (`m365 onboard tui`). A guided flow: new person
  (like a colleague or a bundle, plus add-ons) with a preview of group changes, manual items and the registry
  entry; offboard; check one person; inventory right there when there is no snapshot; recent runs on the left.
  The groups screen no longer has `o` / `x`.
- Quitting always asks. `esc` always goes back one step and, from the top of a tool, to home. `ctrl+1/2/3`
  jump between tools (and `1/2/3` on the onboard screen); the header shows the other tools; every help screen
  explains back and switch.
- The TUI is now the **m365 toolbox**: `m365` opens a home screen listing the tools with their last run
  (`1` spmig, `2` groups, `t` switch tenant, `ctrl+l` sign in); `g` returns home from any tool when nothing
  runs. Every screen shares one header: the mark and name, the current tool, tenant and signed-in account.
  `m365 spmig tui` and `m365 groups tui` open a tool directly. In spmig, `g` now means home; lists go to the
  top with `home`.
- Groups screen: an explorer with People, Groups and Bundles tabs, `/` filter, person and group details
  (bundles, add-ons, each group with ok / missing / unexplained, licences, apps, roles; purpose, owner,
  members and why), registry and snapshot state with drift counts. Actions: inventory with step list and log,
  drift with jump to the person or group, bundle suggestions with accept, sync, import, onboard (like a
  colleague or a bundle, plus add-ons) and offboard, all previewed and applied only after a typed word;
  edit a group's purpose and owner; reload the registry. Help and palette entries for all of it.
- New tool `groups` (access-group registry): `m365 groups inventory|suggest|drift|explain|sync|import`,
  `m365 onboard UPN` (`--bundle`, `--addon`, `--like`), `m365 onboard check UPN`, `m365 offboard UPN`. Writing
  commands preview by default and need `--apply`. Per-tenant TOML registry in `~/.config/m365tools/groups/`.
- Graph client: groups by prefix, members/owners, users with licences, subscribed SKUs, enterprise apps and their
  assignments, directory roles, group creation and membership changes.
- `m365 tenant setup-help` lists the extra Graph permissions the groups tool needs.

## 0.1.0 - 2026-10-21

First release.

- `m365` CLI (`tenant add|list|login|setup-help`, `profile list|show`, `spmig audit|dryrun|migrate|verify|lock|tui|sites`)
  and a Textual TUI.
- Core: TOML config, MSAL public-client auth (browser, WSL `wslview`, device-code fallback, per-tenant token cache),
  async HTTP client with throttling/backoff and `Retry-After` handling, SharePoint REST layer (keyset paging,
  chunked uploads, metadata restore without new versions), Graph helpers, journal/status files, HTML/text/CSV reports.
- `spmig`: library inventory, plan with excludes (case-insensitive globs) and routes (routes win over excludes),
  14 audit checks, resumable migration engine, verification, never-widen source lock, user-map workflow, `Job` API.
- Docs: architecture contract, real-tenant verification checklist, read-only smoke script.
