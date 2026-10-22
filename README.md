# m365 toolbox

Terminal tools for Microsoft 365 admins, built for Linux and WSL. One command, `m365`, opens the toolbox:
pick a tool with `1`, `2` or `3`; `esc` steps back and ends at home, `g` goes home at once, `ctrl+1`/`2`/`3`
jump between tools where the terminal supports it (otherwise `g` then the number), and `q` quits (it asks
first).

| Tool | What it does |
|---|---|
| **spmig** | Copies a SharePoint document library to another site, in your tenant or another one. Keeps folders, file dates and who created or changed each file. Checks first, shows what it would do, writes a report. |
| **groups** | Keeps track of your `SG-*` access groups: what each is for, who is in it and why, what drifted, and which bundles (role packages) exist. |
| **onboard** | Brings people in and out: a new colleague gets the groups of a bundle or of someone in the same role; a leaver is taken out of every registry group. Preview first, always. |

Everything runs in your terminal: a guided screen (`m365`) or plain commands (`m365 spmig ...`,
`m365 groups ...`).

---

## 1. Install (5 minutes)

You need Python 3.12 or newer and [uv](https://docs.astral.sh/uv/) (`curl -LsSf https://astral.sh/uv/install.sh | sh`).

```bash
git clone https://github.com/baki2001/M365-Tools
cd M365-Tools
uv tool install .        # gives you the `m365` command
m365 --help
```

**WSL users:** signing in opens a browser. Run `sudo apt install wslu` once so the tool can open your
Windows browser. If no browser can be opened, you get a code to type in at microsoft.com/devicelogin instead.

## 2. Connect a tenant (once per tenant)

The tool signs in as you, through an app registration in Entra ID. It has no secret and only works with
an interactive login, so nothing sensitive is stored.

If you already have a PnP PowerShell app registration, reuse its Client ID. Otherwise run
`m365 tenant setup-help`; it prints the exact clicks (public client, redirect `http://localhost`,
delegated permissions SharePoint `AllSites.FullControl` and Graph `Sites.Read.All`, `User.Read`, admin consent).

```bash
m365 tenant add contoso --tenant-id contoso.onmicrosoft.com --client-id <app id> --host contoso.sharepoint.com
m365 tenant login contoso      # browser opens, sign in, done
```

Your login is cached, so you won't be asked again for a while.

## 3. Migrate a library (spmig)

### The easy way: guided screen

```bash
m365            # then 1 for spmig (or: m365 spmig tui)
```

One keyboard-driven screen: sections on the left (**1 Job, 2 Checks, 3 Plan, 4 Run, 5 Verify, 6 Lock,
7 Runs**, each with a ○ ◐ ● ✗ state and a one-line summary), the selected section on the right, and the
keys that work right now in the bottom bar. In **Job**, press `e` on Source and Destination to pick tenant,
site (type to search), library and folder; `o` edits options, `s`/`p` save/load a profile. Then `a` audit,
`d` dry run, `m` migrate (asks first; type `force` if the audit found blocking problems), `v` verify.
Move with `1`-`7`, `j`/`k`, `tab` (sidebar ↔ main), `esc` back; `:` command palette, `?` all keys,
`ctrl+l` sign in, `q` quit. Nothing is written to SharePoint until you confirm the migrate step.

### The scripted way: commands

Same steps, one command each. Use this when you run the same migration more than once, or from a script.

```bash
m365 spmig audit   --src-site https://contoso.sharepoint.com/sites/Old --src-library Finance \
                   --dst-site https://contoso.sharepoint.com/sites/Finance
m365 spmig dryrun  ...same options...
m365 spmig migrate ...same options...
m365 spmig verify  ...same options...
```

Typing the same options every time gets old, so save them as a **profile** in `~/.config/m365tools/config.toml`
(the guided screen can save one for you):

```toml
[profiles.finance]
source      = { tenant = "contoso", site_url = "https://contoso.sharepoint.com/sites/Old", library = "Finance" }
destination = { tenant = "contoso", site_url = "https://contoso.sharepoint.com/sites/Finance" }
```

Then: `m365 spmig audit --profile finance`, and so on. Leave out `library` on the destination to use the
site's own Documents library.

### What each step does

| Step | Touches SharePoint? | What you get |
|---|---|---|
| **audit** | read only | List of problems, sorted by severity. Red ones block the migration until fixed. |
| **dryrun** | read only | The exact list of folders and files that would be created, copied or skipped. |
| **migrate** | writes to destination | Live progress with speed and ETA. Press Ctrl+C any time; it stops cleanly. |
| **verify** | read only | Side-by-side check that everything arrived. |
| **lock** | writes to source | Makes the old library read-only so people stop working in it. Shows a preview first; add `--apply` to do it. |

Every step leaves a folder with `report.html` (open it in a browser), `report.txt` and CSV files.
The path is printed at the end.

## 4. Good to know about migrating

**It can be stopped and continued.** Press Ctrl+C and the running file finishes, then it stops.
Start it again later with `m365 spmig migrate --profile X --resume <run folder>` and it carries on
where it left off. Running the same migration again without `--resume` only copies files that are new
or changed since, which is handy right before you switch people over to the new site.

**What comes along:** all folders and files (current version), Created/Modified dates, and Created By /
Modified By when that person exists at the destination.

**What does not:** older versions of files, permissions set on individual folders or files, custom columns,
sharing links. The audit tells you how many items are affected so you can decide if that matters.

**Moving to another tenant:** people's accounts have different addresses there. The audit writes a
`usermap.csv` with every author it found. Fill in the new address per person, then pass the file with
`--usermap usermap.csv`. Anyone not mapped still gets their file, just without the original author name.

**Slow?** SharePoint throttles uploads at roughly 0.5 to 1 MB/s per connection. The tool waits and retries
by itself and shows "throttled, waiting" while it does. More workers rarely help.

## 5. Access groups (groups)

Groups follow one pattern: `SG-<Kind>-<Name>`, one capability each (`SG-Doc-Finance` for a site,
`SG-App-Jira` for an app, `SG-Lic-E3` for a licence). A **bundle** is a role package, like "Servicedesk": a
list of those groups. Bundles live only in a registry file (`~/.config/m365tools/groups/<tenant>.toml`),
never as groups in Entra; onboarding adds a person directly to each group of their bundle.
Run `m365 tenant setup-help` once for the extra Graph permissions.

### The easy way: guided screen

```bash
m365            # then 2 for groups (or: m365 groups tui)
```

It opens on a **Get started** checklist (inventory, your own domain with `D`, import, purposes, bundles,
drift) that always shows the next key to press; `0` or `esc` brings it back.
On the left a list with three tabs, **Groups**, **People** and **Bundles** (`1` `2` `3` or `t`),
`/` to filter, `j`/`k` to move. The right side shows what is highlighted: for a person their bundles,
add-ons and every group with ● ok, ○ missing or ▲ unexplained, plus licences, apps and roles; for a group
its purpose, owner, members and why each one is there. Actions: `i` inventory, `d` drift (`enter` on a
finding jumps to that person or group), `s` suggest bundles (`a` accepts one), `y` sync, `m` import, `e` edit
a group's purpose and owner, `r` reload the registry. Everything that writes shows a preview first; `A`
applies it after you type `apply`.

### The scripted way: commands

```bash
m365 groups inventory                 # read the tenant, save a snapshot
m365 groups suggest                   # propose bundles; --accept 1=Servicedesk --apply writes one
m365 groups drift                     # registry vs tenant
m365 groups explain anna@contoso.com  # why does Anna have what she has (or: a group name)
```

| Step | Touches the tenant? | What you get |
|---|---|---|
| **inventory** | read only | Snapshot of groups, members, licences, apps and roles; a matrix CSV and a report. |
| **suggest** | read only | Candidate bundles from people who share the same groups. Accepting writes only the registry. |
| **drift** | read only | Who is missing a group they should have, who has one nobody can explain, unregistered groups. |
| **sync** / **import** | sync creates groups; import writes the registry | Registry and tenant brought in line, preview first, `--apply` to do it. |

**Good to know:** the tool never deletes a group, and never removes anyone except with `offboard`. Licences,
app assignments and roles are shown but not changed. Each run leaves a folder with `report.html` and CSVs
under `runs/groups-<tenant>/`.

## 6. Onboard and offboard (onboard)

```bash
m365            # then 3 for onboard (or: m365 onboard tui)
```

A short guided flow. `o` new person: type their sign-in name, pick "like <colleague>" or a bundle, add extra
groups if needed, then read the preview: every group change (planned, already, not created), the manual work
(licences, apps, roles) in its own box, and the registry entry that will be saved. `A` and typing `apply` does
it. `x` offboards a leaver the same way (you type `offboard`); `c` checks one person against the registry.
No snapshot yet? Press `i` right there. Recent onboard and offboard runs are listed on the left.

```bash
m365 onboard new.colleague@contoso.com --like anna@contoso.com            # preview
m365 onboard new.colleague@contoso.com --bundle Servicedesk --addon SG-Doc-Finance --apply
m365 onboard check anna@contoso.com
m365 offboard leaver@contoso.com --apply
```

**Good to know:** licences, apps and roles are listed but never changed; groups outside the registry are left
alone on offboarding.

## 7. When something goes wrong

| You see | What it means | What to do |
|---|---|---|
| Red (blocking) findings in audit | Path too long, illegal file name, no write access, or not enough storage | Fix it at the source, or run `migrate --force` if you accept losing those items |
| "user not found" warnings | A former employee is the author of a file | Nothing, or map them to a current user in the usermap |
| "Size differs" in verify for Office files | SharePoint rewrites Office files on upload | Nothing, verify already ignores this for Office files |
| Library not found | Wrong name, or the default library has a localized name | The error lists the libraries that do exist |
| Browser does not open (WSL) | `wslu` not installed | `sudo apt install wslu`, or use the device code shown |

First time against a new tenant? Run the read-only check script and compare with
[docs/REAL-TENANT-CHECKLIST.md](docs/REAL-TENANT-CHECKLIST.md):

```bash
uv run python scripts/smoke_real_tenant.py --tenant contoso --site https://contoso.sharepoint.com/sites/Finance
```

## For the curious

- All options: `m365 spmig migrate --help`. Notable ones: `--exclude "Archive/**"`, `--overwrite always|never`,
  `--workers 4`, `--no-metadata`, `--json` for scripting.
- Run folders live in `~/.local/share/m365tools/runs/`. Config in `~/.config/m365tools/config.toml`.
  Login cache in `~/.cache/m365tools/`.
- Exit codes: 0 ok, 1 failures or blocking findings, 2 migrate refused, 130 stopped by Ctrl+C.
- How it is built, and how to add the next tool: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).
- Development: `uv run pytest` and `uv run ruff check`.
