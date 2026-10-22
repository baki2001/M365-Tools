# Changelog

## Unreleased

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
