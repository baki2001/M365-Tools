"""One place for the look: colour palette, state glyphs, per-value styles and small text formatters.

The same hex colours feed the Textual theme (used by ``app.tcss`` as ``$primary``, ``$error`` ...) and the Rich
styles used inside tables and status lines, so everything on screen comes from one palette.
"""

from __future__ import annotations

from rich.text import Text
from textual.theme import Theme

from m365tools.core.events import LogLine
from m365tools.core.report.format import fmt_duration, fmt_size

# ------------------------------------------------------------------------------------------------ palette

BG = "#16161e"
SURFACE = "#1a1b26"
PANEL = "#24283b"
FG = "#c0caf5"
DIM = "#565f89"
ACCENT = "#7aa2f7"  # focus, selection, key hints
CYAN = "#7dcfff"
MAGENTA = "#bb9af7"
GREEN = "#9ece6a"
YELLOW = "#e0af68"
RED = "#f7768e"

THEME = Theme(
    name="m365",
    primary=ACCENT,
    secondary=MAGENTA,
    accent=ACCENT,
    foreground=FG,
    background=BG,
    surface=SURFACE,
    panel=PANEL,
    success=GREEN,
    warning=YELLOW,
    error=RED,
    dark=True,
    variables={
        "block-cursor-background": PANEL,
        "block-cursor-foreground": FG,
        "block-cursor-text-style": "bold",
        "block-cursor-blurred-background": SURFACE,
        "footer-key-foreground": ACCENT,
        "input-selection-background": f"{ACCENT} 35%",
    },
)

# ------------------------------------------------------------------------------------------------ glyphs

# Section states shown in the sidebar.
TODO, RUNNING, DONE, FAILED = "todo", "running", "done", "failed"
GLYPH = {TODO: "○", RUNNING: "◐", DONE: "●", FAILED: "✗"}
GLYPH_STYLE = {TODO: DIM, RUNNING: f"bold {YELLOW}", DONE: GREEN, FAILED: f"bold {RED}"}
SPINNER = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"
BAR = "▌"

# ------------------------------------------------------------------------------------------------ value styles

SEVERITY_STYLE = {"blocking": f"bold {RED}", "warn": YELLOW, "info": DIM}
ACTION_STYLE = {"create_folder": CYAN, "copy": GREEN, "skip": DIM, "conflict": YELLOW, "excluded": MAGENTA}
ACTION_LABEL = {"create_folder": "create", "copy": "copy", "skip": "skip", "conflict": "conflict",
                "excluded": "excluded"}
VERIFY_STYLE = {"ok": GREEN, "missing": f"bold {RED}", "size_differs": YELLOW, "source_newer": YELLOW,
                "only_in_destination": MAGENTA}
LEVEL_STYLE = {"debug": DIM, "info": FG, "ok": GREEN, "warn": YELLOW, "error": f"bold {RED}"}


# ------------------------------------------------------------------------------------------------ formatters


def size(n: float | None) -> str:
    return "" if n is None else fmt_size(n)


def duration(seconds: float | None) -> str:
    return "--:--:--" if seconds is None else fmt_duration(seconds)


def log_text(ev: LogLine) -> Text:
    """``12:01:02 warn  text`` coloured by level. User data is never parsed as markup."""
    style = LEVEL_STYLE.get(ev.level, FG)
    return Text.assemble((f"{ev.at:%H:%M:%S} ", DIM), (f"{ev.level:<5} ", style), (ev.text, style))


def key_hint(key: str, description: str) -> Text:
    """``key description`` as used in the status bar and help."""
    return Text.assemble((key, f"bold {ACCENT}"), " ", (description, DIM))


def counts_line(counts: dict[str, int], styles: dict[str, str], labels: dict[str, str] | None = None) -> Text:
    """``■ 3 blocking  ■ 5 warn`` with each square in its value colour (zero counts dimmed)."""
    t = Text()
    for k, n in counts.items():
        style = styles.get(k, FG) if n else DIM
        t.append("■ ", style=style).append(f"{n:,} {(labels or {}).get(k, k)}", style=style if n else DIM)
        t.append("   ")
    return t


def short_site(url: str) -> str:
    """``https://contoso.sharepoint.com/sites/Finance`` -> ``Finance``."""
    return url.rstrip("/").rsplit("/", 1)[-1] if url else ""
