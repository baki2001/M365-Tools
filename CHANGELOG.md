# Changelog

## Unreleased

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
