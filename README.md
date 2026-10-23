# m365 toolbox

Terminal tools for Microsoft 365 admins, built for Linux and WSL. Run `m365` to open the toolbox,
then pick a tool with `1`, `2` or `3`. `esc` goes back a step until you're home, `g` jumps straight home,
and `ctrl+1`/`2`/`3` switch between tools if your terminal supports it (if not, use `g` and then the
number). `q` quits after asking.

| Tool | What it does |
|---|---|
| spmig | Copies a SharePoint document library to another site, in your own tenant or a different one. Folders, file dates and the people who created or last changed each file come along. It checks first, shows you what it's going to do, and writes a report. |
| groups | Keeps track of your `SG-*` access groups: what each one is for, who's in it and why, what has drifted, and which bundles (role packages) exist. |
| onboard | Adds and removes people. A new colleague gets the groups from a bundle or from someone in the same role, and a leaver is removed from every registry group. You always see a preview first. |

You can use the guided screen (`m365`) or plain commands (`m365 spmig ...`, `m365 groups ...`). Both run
in the terminal.

## 1. Install (5 minutes)

You need Python 3.12 or newer and [uv](https://docs.astral.sh/uv/) (`curl -LsSf https://astral.sh/uv/install.sh | sh`).

```bash
git clone https://github.com/baki2001/M365-Tools
cd M365-Tools
uv tool install .        # gives you the `m365` command
m365 --help
```

On WSL, signing in opens a browser, so run `sudo apt install wslu` once to let the tool open your Windows
browser. If it can't open a browser at all, you get a code to enter at microsoft.com/devicelogin instead.

## 2. Connect a tenant (once per tenant)

The tool signs in as you through an app registration in Entra ID. The app has no secret and only works
with an interactive login, so nothing sensitive gets stored.

If you already have a PnP PowerShell app registration, you can reuse its Client ID. Otherwise run
`m365 tenant setup-help` and it prints the exact clicks: public client, redirect `http://localhost`,
delegated permissions SharePoint `AllSites.FullControl` and Graph `Sites.Read.All` and `User.Read`, then
admin consent.

```bash
m365 tenant add contoso --tenant-id contoso.onmicrosoft.com --client-id <app id> --host contoso.sharepoint.com
m365 tenant login contoso      # browser opens, sign in, done
```

The login is cached, so it won't ask again for a while.

## 3. Migrate a library (spmig)

### The easy way: guided screen

```bash
m365            # then 1 for spmig (or: m365 spmig tui)
```

Everything happens on one screen and you drive it with the keyboard. The sections are on the left
(1 Job, 2 Checks, 3 Plan, 4 Run, 5 Verify, 6 Lock, 7 Runs), each with a ○ ◐ ● ✗ state and a one-line
summary. The selected section shows on the right, and the bottom bar lists the keys you can use right now.

In Job, press `e` on Source and Destination to pick the tenant, site (type to search), library and folder.
`o` edits options and `s`/`p` save and load a profile. After that it's `a` to audit, `d` for a dry run,
`m` to migrate (it asks first, and you have to type `force` if the audit found blocking problems) and `v`
to verify.

To move around, use `1`-`7`, `j`/`k`, `tab` to switch between the sidebar and the main panel, and `esc` to
go back. `:` opens the command palette, `?` lists all keys, `ctrl+l` signs in and `q` quits. Nothing gets
written to SharePoint until you confirm the migrate step.

### The scripted way: commands

These are the same steps, one command each. Use them when you run the same migration more than once, or
from a script.

```bash
m365 spmig audit   --src-site https://contoso.sharepoint.com/sites/Old --src-library Finance \
                   --dst-site https://contoso.sharepoint.com/sites/Finance
m365 spmig dryrun  ...same options...
m365 spmig migrate ...same options...
m365 spmig verify  ...same options...
```

Typing the same options every time gets old, so save them as a profile in `~/.config/m365tools/config.toml`
(the guided screen can save one for you):

```toml
[profiles.finance]
source      = { tenant = "contoso", site_url = "https://contoso.sharepoint.com/sites/Old", library = "Finance" }
destination = { tenant = "contoso", site_url = "https://contoso.sharepoint.com/sites/Finance" }
```

Then run `m365 spmig audit --profile finance`, and so on. If you leave out `library` on the destination,
it uses the site's own Documents library.

### What each step does

| Step | Touches SharePoint? | What you get |
|---|---|---|
| audit | read only | A list of problems sorted by severity. Red ones block the migration until you fix them. |
| dryrun | read only | The exact folders and files that would be created, copied or skipped. |
| migrate | writes to destination | Live progress with speed and ETA. You can press Ctrl+C at any point and it stops cleanly. |
| verify | read only | A side-by-side check that everything arrived. |
| lock | writes to source | Makes the old library read-only so people stop working in it. It shows a preview first; add `--apply` to actually do it. |

Every step leaves a folder with `report.html` (open it in a browser), `report.txt` and some CSV files.
It prints the path at the end.

## 4. Good to know about migrating

You can stop a migration and pick it up later. Press Ctrl+C and it finishes the file it's on, then stops.
Start it again with `m365 spmig migrate --profile X --resume <run folder>` and it carries on where it left
off. If you run the same migration again without `--resume`, it only copies files that are new or changed
since last time, which is handy right before you move people over to the new site.

It copies all folders and files (current version), the Created and Modified dates, and Created By and
Modified By when that person exists at the destination.

It doesn't copy older versions of files, permissions set on individual folders or files, custom columns
or sharing links. The audit tells you how many items that affects, so you can decide whether it matters.

When you move to another tenant, people's accounts have different addresses there. The audit writes a
`usermap.csv` with every author it found. Fill in each person's new address and pass the file with
`--usermap usermap.csv`. Anyone you don't map still gets their files, just without the original author
name.

If it feels slow, that's SharePoint throttling uploads to roughly 0.5 to 1 MB/s per connection. The tool
waits and retries on its own and shows "throttled, waiting" while it does. Adding workers rarely helps.

## 5. Access groups (groups)

Groups follow one pattern, `SG-<Kind>-<Name>`, with one capability each: `SG-Doc-Finance` for a site,
`SG-App-Jira` for an app, `SG-Lic-E3` for a licence. A bundle is a role package like "Servicedesk", which
is just a list of those groups. Bundles only exist in a registry file
(`~/.config/m365tools/groups/<tenant>.toml`), never as groups in Entra, and onboarding adds a person
directly to each group in their bundle. Run `m365 tenant setup-help` once for the extra Graph permissions.

### The easy way: guided screen

```bash
m365            # then 2 for groups (or: m365 groups tui)
```

It opens on a Get started checklist (inventory, your own domain with `D`, import, purposes, bundles,
drift) that always shows the next key to press. `0` or `esc` brings it back.

On the left is a list with three tabs, Groups, People and Bundles (`1` `2` `3` or `t`). `/` filters and
`j`/`k` move. The right side shows whatever is highlighted. For a person that's their bundles, add-ons and
every group marked ● ok, ○ missing or ▲ unexplained, plus licences, apps and roles. For a group it's the
purpose, owner, members and why each member is there.

The actions are `i` inventory, `d` drift (`enter` on a finding jumps to that person or group), `s` suggest
bundles (`a` accepts one), `y` sync, `m` import, `e` to edit a group's purpose and owner, and `r` to reload
the registry. Anything that writes shows a preview first, and `A` applies it once you type `apply`.

### The scripted way: commands

```bash
m365 groups inventory                 # read the tenant, save a snapshot
m365 groups suggest                   # propose bundles; --accept 1=Servicedesk --apply writes one
m365 groups drift                     # registry vs tenant
m365 groups explain anna@contoso.com  # why does Anna have what she has (or: a group name)
```

| Step | Touches the tenant? | What you get |
|---|---|---|
| inventory | read only | A snapshot of groups, members, licences, apps and roles, with a matrix CSV and a report. |
| suggest | read only | Candidate bundles based on people who share the same groups. Accepting one only writes the registry. |
| drift | read only | Who is missing a group they should have, who has one nobody can explain, and which groups aren't registered. |
| sync / import | sync creates groups; import writes the registry | Brings the registry and the tenant in line. Preview first, `--apply` to do it. |

The tool never deletes a group, and it only removes people when you run `offboard`. Licences, app
assignments and roles are shown but never changed. Each run leaves a folder with `report.html` and CSVs
under `runs/groups-<tenant>/`.

## 6. Onboard and offboard (onboard)

```bash
m365            # then 3 for onboard (or: m365 onboard tui)
```

This one is a short guided flow. Press `o` for a new person and type their sign-in name, pick `like
<colleague>` or a bundle, and add extra groups if needed. Then read the preview. It shows every group
change (planned, already, not created), the manual work (licences, apps, roles) in its own box, and the
registry entry that will be saved. Press `A` and type `apply` to do it.

`x` offboards a leaver the same way (you type `offboard`), and `c` checks one person against the
registry. If you don't have a snapshot yet, press `i` right there. Recent onboard and offboard runs are
listed on the left.

```bash
m365 onboard new.colleague@contoso.com --like anna@contoso.com            # preview
m365 onboard new.colleague@contoso.com --bundle Servicedesk --addon SG-Doc-Finance --apply
m365 onboard check anna@contoso.com
m365 offboard leaver@contoso.com --apply
```

Licences, apps and roles are listed but never changed, and offboarding leaves groups outside the
registry alone.

## 7. When something goes wrong

| You see | What it means | What to do |
|---|---|---|
| Red (blocking) findings in audit | Path too long, illegal file name, no write access, or not enough storage | Fix it at the source, or run `migrate --force` if you're fine losing those items |
| "user not found" warnings | A former employee is the author of a file | Nothing, or map them to a current user in the usermap |
| "Size differs" in verify for Office files | SharePoint rewrites Office files on upload | Nothing, verify already ignores this for Office files |
| Library not found | Wrong name, or the default library has a localized name | The error lists the libraries that do exist |
| Browser does not open (WSL) | `wslu` isn't installed | `sudo apt install wslu`, or use the device code shown |

Trying it on a new tenant for the first time? Run the read-only check script and compare the results with
[docs/REAL-TENANT-CHECKLIST.md](docs/REAL-TENANT-CHECKLIST.md):

```bash
uv run python scripts/smoke_real_tenant.py --tenant contoso --site https://contoso.sharepoint.com/sites/Finance
```

## More details

- All options: `m365 spmig migrate --help`. Useful ones are `--exclude "Archive/**"`,
  `--overwrite always|never`, `--workers 4`, `--no-metadata` and `--json` for scripting.
- Run folders are in `~/.local/share/m365tools/runs/`, config is in `~/.config/m365tools/config.toml`,
  and the login cache is in `~/.cache/m365tools/`.
- Exit codes: 0 ok, 1 failures or blocking findings, 2 migrate refused, 130 stopped by Ctrl+C.
- How it's built and how to add the next tool: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).
- Development: `uv run pytest` and `uv run ruff check`.
