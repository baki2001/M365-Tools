# groups: access registry, drift and onboarding

Second tool in the toolbox. Keeps track of the `SG-*` access groups in a tenant, what they are for,
who is in them, and what a new person should get. Built for a small company (tens of people) where
roles are fuzzy and most people are "a bundle plus a few extras".

## Decisions (researched, do not revisit)

- **Flat groups only.** Entra does not honour nested groups for enterprise-app assignment or for
  group-based licensing, and SharePoint only partly. So a bundle ("role") is a registry concept,
  never an Entra group that contains other groups. Onboarding adds a person **directly** to every
  capability group in the bundle.
- **One group = one capability.** Naming `SG-<Kind>-<Name>`: `SG-Doc-Finance` (SharePoint site),
  `SG-App-Jira` (enterprise app), `SG-Lic-E3` (licence group), `SG-Role-...` is **not** used
  (bundles are not groups). Prefix is configurable per tenant (default `SG-`).
- **Inventory first, roles later.** The tool imports what exists. Bundles are proposed from actual
  memberships; the admin names the ones that make sense. Exceptions are allowed and always visible.
- **Registry is a TOML file** per tenant at `~/.config/m365tools/groups/<tenant>.toml`. Human-edited
  and tool-edited; the tool never deletes entries, only adds or marks.
- **Writes are explicit.** Every writing command previews first and needs `--apply` (CLI) or a typed
  confirmation (TUI), like `spmig lock`. Nothing is ever removed from the tenant except by
  `offboard --apply`, and groups are never deleted.
- Graph only (`https://graph.microsoft.com/.default`, delegated). Extra app permissions the admin
  grants once: `Group.Read.All`, `User.Read.All`, `Directory.Read.All`, `Application.Read.All`
  for reading; `Group.ReadWrite.All`, `GroupMember.ReadWrite.All` for sync/onboard/offboard.
  Licences and app assignments are **read-only** in v1 (shown, not changed). `m365 tenant setup-help`
  lists them.

## Registry file

```toml
prefix = "SG-"

[groups."SG-Doc-Finance"]
id = "..."                       # Entra object id once it exists (filled by sync/inventory)
kind = "doc"                     # doc | app | lic | other
purpose = "Read/write in the Finance SharePoint site"
owner = "someone@contoso.com"     # who decides membership (informational)

[bundles.Servicedesk]
purpose = "Everyone on the service desk"
groups = ["SG-Doc-Service", "SG-Doc-Hub", "SG-App-Jira", "SG-Lic-E3"]

[people."evelien@contoso.com"]
bundles = ["Servicedesk"]
addons = ["SG-Doc-Finance"]       # named exceptions on top of the bundles
note = "Also does invoicing"      # optional, free text
```

Expected groups for a person = union(bundle groups) ∪ addons. A person not in `[people]` has no
expectation and is simply reported as "unassigned" with their actual groups.

## Inventory (read-only, `groups inventory`)

Pulls into one `Snapshot` (saved as `snapshot.json` in the run folder):

- Groups whose displayName starts with the prefix: id, name, description, securityEnabled,
  mailEnabled, groupTypes, members (direct, users only), owners.
- Enabled users (members, not guests): id, upn, displayName, jobTitle, department, accountEnabled,
  assignedLicenses → sku part numbers via `/subscribedSkus`.
- Enterprise apps with assignment required (`appRoleAssignmentRequired eq true`): name +
  `appRoleAssignedTo` (users and groups) → shown as capability `App:<name>`.
- Directory role members (`/directoryRoles` → `/members`) → capability `Role:<name>`.

Matrix = people × capabilities (groups, `App:`, `Lic:`, `Role:`). Written as `matrix.csv`.
Phases emitted like spmig audit (`groups → users → licences → apps → roles → report`) so the TUI step
list works unchanged.

## Suggest (`groups suggest`)

No ML. Group people by their exact set of registry-prefix groups. Each set shared by ≥ 2 people is a
candidate bundle, named after the most common jobTitle/department of its people (editable). For every
other person: nearest candidate by Jaccard, shown as `candidate + addons − missing`. The admin accepts
candidates (keypress / `--accept NAME=...`) and they are written to `[bundles]` and `[people]`.

## Drift (`groups drift`)

Compares registry expectation with the snapshot. Findings, severity like spmig:

- `blocking`: none in v1.
- `warn`: person has a prefix group that is not explained (not in bundles/addons) — "unexplained";
  person is missing an expected group — "missing"; prefix group exists in tenant but not in
  registry — "unregistered"; registry group does not exist in tenant — "not created".
- `info`: person not in `[people]` — "unassigned"; group has no owner; group is empty.

Each finding has a `fix` (the command that resolves it). Report uses `core/report` like spmig.

## Explain (`groups explain <person|group>`)

Person: bundles, addons, actual groups with purpose, unexplained/missing, licences, apps, roles.
Group: purpose, owner, members, which bundles include it, who has it as add-on, unexplained holders.

## Sync (`groups sync`)

Registry → tenant, preview then `--apply`: create missing groups (security, mail-disabled,
description = purpose), fill `id`. Never removes members, never deletes groups.
Tenant → registry: `groups import` adds unregistered prefix groups to `[groups]` with
`purpose = ""` so the admin fills them in (and the drift report nags until they do).

## Onboard / offboard

- `m365 onboard <upn> --bundle Servicedesk [--addon SG-Doc-Finance ...]` or `--like <upn>`
  (copies that person's registry entry, or their actual groups if unregistered). Preview: which
  groups will be added, which licences/apps the bundle implies but the tool cannot set (listed as
  "manual"). `--apply` adds memberships and writes the `[people]` entry.
- `m365 offboard <upn>`: preview removal from every registry group (and the `[people]` entry gets
  `offboarded = "2026-10-22"`), `--apply` does it. Does not disable the account.
- `m365 onboard check <upn>`: drift for one person.

## Code layout

```
tools/groups/
  models.py      Snapshot, GroupInfo, Person, Capability, Registry, Bundle, PersonEntry, Finding(reuse spmig? no: own small Finding with same fields)
  registry.py    load/save TOML (tomllib + tomli_w), expectations(person) -> set[str]
  inventory.py   build Snapshot from GraphGateway with phases
  suggest.py     candidates(snapshot, registry) -> list[Candidate]
  drift.py       findings(snapshot, registry) -> list[Finding]
  explain.py     explain_person / explain_group -> structured dict for CLI/TUI/report
  sync.py        plan_sync / apply_sync, plan_import
  onboard.py     plan_onboard / apply, plan_offboard / apply
  gateway.py     GraphGateway Protocol + RestGraphGateway (core/graph.py additions) + FakeGraphGateway (tests)
  job.py         GroupsJob(gateway, registry_path, sink, run_dir_factory): inventory(), suggest(), drift(), explain(), sync(apply), import_(apply), onboard(..., apply), offboard(..., apply)
```

CLI: `m365 groups inventory|suggest|drift|explain|sync|import` and `m365 onboard ...`,
`m365 offboard ...`, all with `--tenant` (default: the only/first tenant) and `--json`.
TUI: toolbox home (pick spmig / groups), groups workspace with sections
**1 Inventory, 2 Groups, 3 People, 4 Bundles, 5 Drift, 6 Sync, 7 Onboard**, same keys, theme and widgets as spmig.
