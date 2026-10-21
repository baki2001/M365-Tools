"""Self-contained HTML report renderer."""

from __future__ import annotations

from html import escape as _e

from .format import fmt_duration
from .model import Report, Section

CSS = """
:root{--bg:#fff;--fg:#1f2328;--mut:#656d76;--card:#f6f8fa;--bd:#d0d7de;--bl:#ffebe9;--blf:#a40e26;
--wr:#fff3cd;--wrf:#7a4b00;--in:#eaeef2;--inf:#424a53}
@media (prefers-color-scheme:dark){:root{--bg:#0d1117;--fg:#e6edf3;--mut:#8b949e;--card:#161b22;--bd:#30363d;
--bl:#4c1d24;--blf:#ffb3b8;--wr:#4a3a0a;--wrf:#f2cc60;--in:#262c36;--inf:#b1bac4}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,
"Helvetica Neue",Arial,sans-serif}
main{max-width:1100px;margin:0 auto;padding:24px 16px}
h1{margin:0 0 4px;font-size:24px}
.meta{color:var(--mut);margin:0 0 16px}.meta b{color:var(--fg);font-weight:600}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;margin:16px 0}
.card{background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:10px 12px}
.card .k{color:var(--mut);font-size:12px}.card .v{font-size:20px;font-weight:600;overflow-wrap:anywhere}
details{border:1px solid var(--bd);border-radius:8px;margin:12px 0;background:var(--bg)}
summary{cursor:pointer;padding:10px 14px;font-weight:600;font-size:16px}
.body{padding:0 14px 14px}.note{color:var(--mut);margin:0 0 8px}
.wrap{overflow:auto;max-height:70vh}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{text-align:left;padding:5px 8px;border-bottom:1px solid var(--bd);vertical-align:top;overflow-wrap:anywhere}
th{position:sticky;top:0;background:var(--card);z-index:1}
.kv th{position:static;width:30%;background:none;font-weight:600}
.sev-blocking{background:var(--bl);color:var(--blf);font-weight:600}
.sev-warn{background:var(--wr);color:var(--wrf);font-weight:600}
.sev-info{background:var(--in);color:var(--inf)}
.callout{border:1px solid var(--wrf);background:var(--wr);color:var(--wrf);border-radius:8px;padding:10px 14px;margin:16px 0}
.callout h2{margin:0 0 6px;font-size:16px}.callout ul{margin:0;padding-left:20px}
pre{white-space:pre-wrap;margin:0}
footer{color:var(--mut);font-size:12px;margin-top:24px;border-top:1px solid var(--bd);padding-top:8px}
@media print{:root{--bg:#fff;--fg:#000}body{font-size:11px}.wrap{max-height:none;overflow:visible}
details{break-inside:avoid-page}th{position:static}main{max-width:none;padding:0}}
"""


def _table(s: Section) -> str:
    head = "".join(f"<th>{_e(str(c))}</th>" for c in s.columns)
    body = []
    for r in s.rows:
        tds = []
        for i, c in enumerate(r):
            cls = ""
            if s.severity_col == i:
                v = str(c).strip().lower()
                if v in ("blocking", "warn", "info"):
                    cls = f' class="sev-{v}"'
            tds.append(f"<td{cls}>{_e(str(c))}</td>")
        body.append("<tr>" + "".join(tds) + "</tr>")
    if not body:
        return '<p class="note">(none)</p>'
    return f'<div class="wrap"><table><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'


def _kv(s: Section) -> str:
    rows = "".join(
        f"<tr><th>{_e(str(r[0]))}</th><td>{_e(str(r[1]) if len(r) > 1 else '')}</td></tr>" for r in s.rows
    )
    return f'<table class="kv"><tbody>{rows}</tbody></table>'


def _section(s: Section) -> str:
    inner = f'<p class="note">{_e(s.note)}</p>' if s.note else ""
    inner += {"table": _table, "kv": _kv}.get(s.kind, lambda x: f"<pre>{_e(x.text)}</pre>")(s)
    op = "" if s.collapsed else " open"
    return f'<details{op}><summary>{_e(s.title)}</summary><div class="body">{inner}</div></details>'


def render_html(report: Report) -> str:
    dur = fmt_duration((report.finished - report.started).total_seconds())
    cards = "".join(
        f'<div class="card"><div class="k">{_e(k)}</div><div class="v">{_e(v)}</div></div>'
        for k, v in report.summary
    )
    callout = ""
    if report.not_migrated:
        li = "".join(f"<li>{_e(x)}</li>" for x in report.not_migrated)
        callout = f'<div class="callout"><h2>What was NOT migrated</h2><ul>{li}</ul></div>'
    foot = f"<div>{_e(report.footer)}</div>" if report.footer else ""
    return (
        '<!doctype html>\n<html lang="en"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        f"<title>{_e(report.title)}</title><style>{CSS}</style></head><body><main>"
        f"<h1>{_e(report.title)}</h1>"
        f'<p class="meta"><b>Mode:</b> {_e(report.mode)} &middot; <b>Source:</b> {_e(report.source)} '
        f"&middot; <b>Destination:</b> {_e(report.destination)}<br>"
        f"<b>Started:</b> {report.started:%Y-%m-%d %H:%M:%S} &middot; "
        f"<b>Finished:</b> {report.finished:%Y-%m-%d %H:%M:%S} &middot; <b>Duration:</b> {dur}</p>"
        f'<div class="cards">{cards}</div>{callout}'
        + "".join(_section(s) for s in report.sections)
        + f"<footer>{foot}<div>Generated by m365tools spmig</div></footer></main></body></html>\n"
    )
