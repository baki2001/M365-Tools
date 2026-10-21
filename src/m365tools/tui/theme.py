"""One place for the look: the Catppuccin Mocha palette, state glyphs, per-value colours and small renderers.

The palette is registered as the Textual theme ``m365`` and every token is also a CSS variable (``$mauve``,
``$surface1`` ...), so ``app.tcss`` and the Rich renderables built here come from the same colours.

Roles: mauve = focus and selection, blue = keys and links, green / yellow / red = ok / warn / error,
peach = "needs your input", subtext = secondary text, overlay = muted / disabled.
"""

from __future__ import annotations

from collections.abc import Iterable

from rich.text import Text
from textual.theme import Theme

from m365tools.core.events import LogLine
from m365tools.core.report.format import fmt_duration, fmt_size

# ------------------------------------------------------------------------------------------------ palette

BASE = "#1e1e2e"
MANTLE = "#181825"
CRUST = "#11111b"
SURFACE0 = "#313244"
SURFACE1 = "#45475a"
TEXT = "#cdd6f4"
SUBTEXT0 = "#a6adc8"
OVERLAY0 = "#6c7086"
BLUE = "#89b4fa"
LAVENDER = "#b4befe"
MAUVE = "#cba6f7"
GREEN = "#a6e3a1"
YELLOW = "#f9e2af"
PEACH = "#fab387"
RED = "#f38ba8"
TEAL = "#94e2d5"

# Role names used across the TUI.
FG = TEXT
SUB = SUBTEXT0  # secondary text: labels, summaries, breadcrumbs
DIM = OVERLAY0  # muted: zero counts, placeholders, separators
ACCENT = MAUVE  # focus and selection
KEY = BLUE  # keys and links

PALETTE = {
    "base": BASE, "mantle": MANTLE, "crust": CRUST, "surface0": SURFACE0, "surface1": SURFACE1,
    "subtext0": SUBTEXT0, "overlay0": OVERLAY0, "blue": BLUE, "lavender": LAVENDER, "mauve": MAUVE,
    "green": GREEN, "yellow": YELLOW, "peach": PEACH, "red": RED, "teal": TEAL,
}

THEME_VARIABLES = {
    **PALETTE,
    "block-cursor-background": MAUVE,
    "block-cursor-foreground": CRUST,
    "block-cursor-text-style": "bold",
    "block-cursor-blurred-background": SURFACE1,
    "block-cursor-blurred-foreground": TEXT,
    "block-cursor-blurred-text-style": "none",
    "block-hover-background": f"{SURFACE0} 60%",
    "border": MAUVE,
    "border-blurred": SURFACE1,
    "footer-key-foreground": BLUE,
    "input-cursor-background": MAUVE,
    "input-selection-background": f"{MAUVE} 35%",
    "scrollbar": SURFACE1,
    "scrollbar-hover": OVERLAY0,
    "scrollbar-active": MAUVE,
    "scrollbar-background": BASE,
    "scrollbar-background-hover": BASE,
    "scrollbar-background-active": BASE,
    "scrollbar-corner-color": BASE,
}

THEME = Theme(
    name="m365",
    primary=MAUVE,
    secondary=BLUE,
    accent=MAUVE,
    foreground=TEXT,
    background=BASE,
    surface=MANTLE,
    panel=SURFACE0,
    boost=SURFACE0,
    success=GREEN,
    warning=YELLOW,
    error=RED,
    dark=True,
    variables=THEME_VARIABLES,
)

# ------------------------------------------------------------------------------------------------ glyphs

# Section states shown in the sidebar and the pipeline strip.
TODO, RUNNING, DONE, FAILED = "todo", "running", "done", "failed"
GLYPH = {TODO: "○", RUNNING: "◐", DONE: "●", FAILED: "✗"}
STATE_COLOR = {TODO: OVERLAY0, RUNNING: YELLOW, DONE: GREEN, FAILED: RED}
GLYPH_STYLE = {TODO: OVERLAY0, RUNNING: f"bold {YELLOW}", DONE: GREEN, FAILED: f"bold {RED}"}
SPINNER = "◐◓◑◒"
DOT = "●"
BLOCK = "▮"
SEP = "  ·  "

# ------------------------------------------------------------------------------------------------ value colours

SEVERITY_STYLE = {"blocking": RED, "warn": YELLOW, "info": TEAL}
ACTION_STYLE = {"create_folder": TEAL, "copy": GREEN, "skip": OVERLAY0, "conflict": YELLOW, "excluded": MAUVE}
ACTION_LABEL = {"create_folder": "create", "copy": "copy", "skip": "skip", "conflict": "conflict",
                "excluded": "excluded"}
VERIFY_STYLE = {"ok": GREEN, "missing": RED, "size_differs": YELLOW, "source_newer": PEACH,
                "only_in_destination": MAUVE}
VERIFY_LABEL = {"ok": "ok", "missing": "missing", "size_differs": "size differs", "source_newer": "source newer",
                "only_in_destination": "only in destination"}
LEVEL_STYLE = {"debug": OVERLAY0, "info": TEXT, "ok": GREEN, "warn": YELLOW, "error": f"bold {RED}"}
MODE_STYLE = {"audit": BLUE, "dryrun": TEAL, "migrate": GREEN, "verify": MAUVE, "lock": PEACH}


# ------------------------------------------------------------------------------------------------ formatters


def size(n: float | None) -> str:
    return "" if n is None else fmt_size(n)


def duration(seconds: float | None) -> str:
    return "--:--:--" if seconds is None else fmt_duration(seconds)


def short_site(url: str) -> str:
    """``https://contoso.sharepoint.com/sites/Finance`` -> ``Finance``."""
    return url.rstrip("/").rsplit("/", 1)[-1] if url else ""


# ------------------------------------------------------------------------------------------------ renderers


def pill(text: str, color: str, *, fg: str = CRUST, bold: bool = True) -> Text:
    """`` text `` as a filled pill: ``fg`` on ``color``."""
    return Text(f" {text} ", style=f"{'bold ' if bold else ''}{fg} on {color}")


def chip(key: str) -> Text:
    """A key as a reverse-video chip: `` a `` in bold text on surface1."""
    return Text(f" {key} ", style=f"bold {TEXT} on {SURFACE1}")


def key_hint(key: str, description: str) -> Text:
    """`` a  audit``: key chip, then the label in subtext (key bar, help, empty states)."""
    return Text.assemble(chip(key), " ", (description, SUB))


def key_hints(pairs: Iterable[tuple[str, str]], sep: str = "  ") -> Text:
    t = Text()
    for i, (k, d) in enumerate(pairs):
        if i:
            t.append(sep)
        t.append_text(key_hint(k, d))
    return t


def glyph_word(word: str, color: str, glyph: str = DOT) -> Text:
    """``● blocking`` with glyph and word in the value colour."""
    return Text.assemble((glyph, color), " ", (word, color))


def counts_line(counts: dict[str, int], styles: dict[str, str], labels: dict[str, str] | None = None) -> Text:
    """The summary ribbon: one filled pill per value (`` ● 3 blocking ``); zero counts are muted pills."""
    t = Text()
    for k, n in counts.items():
        label = f"{DOT} {n:,} {(labels or {}).get(k, k)}"
        t.append_text(pill(label, styles.get(k, SUB)) if n else pill(label, SURFACE0, fg=OVERLAY0, bold=False))
        t.append("  ")
    return t


def state_word(status: str, words: dict[str, str] | None = None) -> Text:
    w = (words or {}).get(status, {TODO: "todo", RUNNING: "running", DONE: "done", FAILED: "failed"}[status])
    return Text(w, style=STATE_COLOR[status])


def empty_text(title: str, body: str, keys: Iterable[tuple[str, str]] = ()) -> Text:
    """Centred empty state: a bold line, a hint in subtext, then key chips."""
    t = Text(justify="center")
    t.append(title, style=f"bold {TEXT}").append("\n\n").append(body, style=SUB)
    keys = list(keys)
    if keys:
        t.append("\n\n").append_text(key_hints(keys, sep="    "))
    return t


def log_text(ev: LogLine) -> Text:
    """``12:01:02 warn  text`` coloured by level. User data is never parsed as markup."""
    style = LEVEL_STYLE.get(ev.level, TEXT)
    return Text.assemble((f"{ev.at:%H:%M:%S} ", OVERLAY0), (f"{ev.level:<5} ", style), (ev.text, style))
