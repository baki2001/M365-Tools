# m365tools

Terminal tools for Microsoft 365 admins, built for Linux and WSL.

The first tool is **spmig**: it copies a SharePoint document library from one site to another,
inside your tenant or to a different tenant. It keeps the folder structure, the file dates and
who created or changed each file. It checks everything first, shows you what it would do, and
writes a report you can hand to someone else.

Everything runs in your terminal: a guided screen (`m365`) or plain commands (`m365 spmig ...`).

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

## 3. Migrate a library

### The easy way: guided screen

```bash
m365
```

Pick the tenant, type part of the site name, pick the site and the library. Do the same for the
destination. The screen then walks you through **audit → dry run → migrate → verify**. Nothing is
written to SharePoint until you confirm the migrate step.

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

## 4. Good to know

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

## 5. When something goes wrong

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
