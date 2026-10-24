"""One place for the look: the colour themes, state glyphs, per-value colours and small renderers.

A theme is a :class:`Palette` of sixteen roles. Each palette becomes a Textual theme whose CSS variables
(``$accent``, ``$surface1``, ``$fg-sub`` ...) style ``app.tcss``. Renderables built in Python use the role
constants below (``ACCENT``, ``RED``, ``SUB`` ...). Those are placeholder colours that :class:`PaletteFilter`
swaps for the active palette when a line is painted, so switching the theme recolours everything already on
screen without rebuilding it.

Roles: accent = focus and selection, blue = keys and links, green / yellow / red = ok / warn / error,
orange = "needs your input", sub = secondary text, dim = muted / disabled, on_fill = text on a filled colour.
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, fields
from functools import lru_cache

from rich.color import Color as RichColor
from rich.segment import Segment
from rich.style import Style
from rich.table import Table
from rich.text import Text
from textual.color import Color
from textual.filter import LineFilter
from textual.theme import Theme

from m365tools.core.events import LogLine
from m365tools.core.report.format import fmt_duration, fmt_size

# ------------------------------------------------------------------------------------------------ palettes


@dataclass(frozen=True, slots=True)
class Palette:
    """One colour theme. Colours are ``#rrggbb``, or ANSI names (``red``, ``bright_black``, ``default``) when
    ``ansi`` is set, so the terminal's own colours show through."""

    label: str
    dark: bool
    bg: str  # main background
    bg_alt: str  # sidebars, header and key bar
    on_fill: str  # text on a filled colour (pills, the selection bar)
    surface0: str  # quiet fills: table header, panels
    surface1: str  # key chips, unfocused selection, box borders
    fg: str
    sub: str  # labels, summaries
    dim: str  # zero counts, placeholders, separators
    accent: str
    red: str
    green: str
    yellow: str
    orange: str
    blue: str
    cyan: str
    purple: str
    ansi: bool = False


ROLES = tuple(f.name for f in fields(Palette) if f.name not in ("label", "dark", "ansi"))

DEFAULT_THEME = "catppuccin-mocha"

PALETTES: dict[str, Palette] = {
    "catppuccin-mocha": Palette(
        "Catppuccin Mocha", True, bg="#1e1e2e", bg_alt="#181825", on_fill="#11111b", surface0="#313244",
        surface1="#45475a", fg="#cdd6f4", sub="#a6adc8", dim="#6c7086", accent="#cba6f7", red="#f38ba8",
        green="#a6e3a1", yellow="#f9e2af", orange="#fab387", blue="#89b4fa", cyan="#94e2d5", purple="#cba6f7"),
    "catppuccin-macchiato": Palette(
        "Catppuccin Macchiato", True, bg="#24273a", bg_alt="#1e2030", on_fill="#181926", surface0="#363a4f",
        surface1="#494d64", fg="#cad3f5", sub="#a5adcb", dim="#6e738d", accent="#c6a0f6", red="#ed8796",
        green="#a6da95", yellow="#eed49f", orange="#f5a97f", blue="#8aadf4", cyan="#8bd5ca", purple="#c6a0f6"),
    "catppuccin-frappe": Palette(
        "Catppuccin Frappé", True, bg="#303446", bg_alt="#292c3c", on_fill="#232634", surface0="#414559",
        surface1="#51576d", fg="#c6d0f5", sub="#a5adce", dim="#737994", accent="#ca9ee6", red="#e78284",
        green="#a6d189", yellow="#e5c890", orange="#ef9f76", blue="#8caaee", cyan="#81c8be", purple="#ca9ee6"),
    "catppuccin-latte": Palette(
        "Catppuccin Latte", False, bg="#eff1f5", bg_alt="#e6e9ef", on_fill="#eff1f5", surface0="#dce0e8",
        surface1="#ccd0da", fg="#4c4f69", sub="#6c6f85", dim="#9ca0b0", accent="#8839ef", red="#d20f39",
        green="#40a02b", yellow="#df8e1d", orange="#fe640b", blue="#1e66f5", cyan="#179299", purple="#8839ef"),
    "dracula": Palette(
        "Dracula", True, bg="#282a36", bg_alt="#21222c", on_fill="#191a21", surface0="#343746",
        surface1="#44475a", fg="#f8f8f2", sub="#c0c2d0", dim="#6272a4", accent="#bd93f9", red="#ff5555",
        green="#50fa7b", yellow="#f1fa8c", orange="#ffb86c", blue="#8be9fd", cyan="#8be9fd", purple="#ff79c6"),
    "gruvbox-dark": Palette(
        "Gruvbox Dark", True, bg="#282828", bg_alt="#1d2021", on_fill="#1d2021", surface0="#3c3836",
        surface1="#504945", fg="#ebdbb2", sub="#bdae93", dim="#7c6f64", accent="#d3869b", red="#fb4934",
        green="#b8bb26", yellow="#fabd2f", orange="#fe8019", blue="#83a598", cyan="#8ec07c", purple="#d3869b"),
    "gruvbox-light": Palette(
        "Gruvbox Light", False, bg="#fbf1c7", bg_alt="#f2e5bc", on_fill="#fbf1c7", surface0="#ebdbb2",
        surface1="#d5c4a1", fg="#3c3836", sub="#665c54", dim="#a89984", accent="#8f3f71", red="#9d0006",
        green="#79740e", yellow="#b57614", orange="#af3a03", blue="#076678", cyan="#427b58", purple="#8f3f71"),
    "nord": Palette(
        "Nord", True, bg="#2e3440", bg_alt="#292e39", on_fill="#2e3440", surface0="#3b4252",
        surface1="#434c5e", fg="#eceff4", sub="#d8dee9", dim="#6b7489", accent="#88c0d0", red="#bf616a",
        green="#a3be8c", yellow="#ebcb8b", orange="#d08770", blue="#81a1c1", cyan="#8fbcbb", purple="#b48ead"),
    "tokyo-night": Palette(
        "Tokyo Night", True, bg="#1a1b26", bg_alt="#16161e", on_fill="#16161e", surface0="#292e42",
        surface1="#3b4261", fg="#c0caf5", sub="#a9b1d6", dim="#565f89", accent="#bb9af7", red="#f7768e",
        green="#9ece6a", yellow="#e0af68", orange="#ff9e64", blue="#7aa2f7", cyan="#7dcfff", purple="#bb9af7"),
    "one-dark": Palette(
        "One Dark", True, bg="#282c34", bg_alt="#21252b", on_fill="#1e2127", surface0="#2c313a",
        surface1="#3e4451", fg="#dcdfe4", sub="#abb2bf", dim="#5c6370", accent="#c678dd", red="#e06c75",
        green="#98c379", yellow="#e5c07b", orange="#d19a66", blue="#61afef", cyan="#56b6c2", purple="#c678dd"),
    "solarized-dark": Palette(
        "Solarized Dark", True, bg="#002b36", bg_alt="#00252e", on_fill="#002b36", surface0="#073642",
        surface1="#0d4654", fg="#93a1a1", sub="#839496", dim="#586e75", accent="#6c71c4", red="#dc322f",
        green="#859900", yellow="#b58900", orange="#cb4b16", blue="#268bd2", cyan="#2aa198", purple="#d33682"),
    "solarized-light": Palette(
        "Solarized Light", False, bg="#fdf6e3", bg_alt="#eee8d5", on_fill="#fdf6e3", surface0="#eee8d5",
        surface1="#ddd6c1", fg="#586e75", sub="#657b83", dim="#93a1a1", accent="#6c71c4", red="#dc322f",
        green="#859900", yellow="#b58900", orange="#cb4b16", blue="#268bd2", cyan="#2aa198", purple="#d33682"),
    "rose-pine": Palette(
        "Rosé Pine", True, bg="#191724", bg_alt="#1f1d2e", on_fill="#191724", surface0="#26233a",
        surface1="#403d52", fg="#e0def4", sub="#908caa", dim="#6e6a86", accent="#c4a7e7", red="#eb6f92",
        green="#9ccfd8", yellow="#f6c177", orange="#ebbcba", blue="#9ccfd8", cyan="#3e8fb0", purple="#c4a7e7"),
    "kanagawa": Palette(
        "Kanagawa", True, bg="#1f1f28", bg_alt="#16161d", on_fill="#16161d", surface0="#2a2a37",
        surface1="#363646", fg="#dcd7ba", sub="#c8c093", dim="#727169", accent="#957fb8", red="#e46876",
        green="#98bb6c", yellow="#e6c384", orange="#ffa066", blue="#7e9cd8", cyan="#7aa89f", purple="#957fb8"),
    "everforest": Palette(
        "Everforest", True, bg="#2d353b", bg_alt="#232a2e", on_fill="#232a2e", surface0="#343f44",
        surface1="#475258", fg="#d3c6aa", sub="#9da9a0", dim="#7a8478", accent="#a7c080", red="#e67e80",
        green="#a7c080", yellow="#dbbc7f", orange="#e69875", blue="#7fbbb3", cyan="#83c092", purple="#d699b6"),
    "terminal": Palette(
        "Terminal colours", True, bg="default", bg_alt="default", on_fill="black", surface0="default",
        surface1="bright_black", fg="default", sub="default", dim="bright_black", accent="magenta", red="red",
        green="green", yellow="yellow", orange="bright_red", blue="blue", cyan="cyan", purple="magenta",
        ansi=True),
}


def palette(name: str | None) -> Palette:
    return PALETTES.get(name or "", PALETTES[DEFAULT_THEME])


def _css(p: Palette, value: str) -> str:
    return f"ansi_{value}" if p.ansi else value


def css_variables(p: Palette) -> dict[str, str]:
    """The CSS variables ``app.tcss`` uses, plus the Textual component colours that follow the palette."""
    c = {role: _css(p, getattr(p, role)) for role in ROLES}
    ansi = {"ansi-background": "ansi_black", "ansi-foreground": "ansi_white"} if p.ansi else {}
    return {
        **ansi,
        "bg": c["bg"], "bg-alt": c["bg_alt"], "on-fill": c["on_fill"], "surface0": c["surface0"],
        "surface1": c["surface1"], "fg-sub": c["sub"], "fg-dim": c["dim"], "red": c["red"], "green": c["green"],
        "yellow": c["yellow"], "orange": c["orange"], "blue": c["blue"], "cyan": c["cyan"], "purple": c["purple"],
        "block-cursor-background": c["accent"],
        "block-cursor-foreground": c["on_fill"],
        "block-cursor-text-style": "bold",
        "block-cursor-blurred-background": c["surface1"],
        "block-cursor-blurred-foreground": c["fg"],
        "block-cursor-blurred-text-style": "none",
        "block-hover-background": c["surface0"],
        "border": c["accent"],
        "border-blurred": c["surface1"],
        "footer-key-foreground": c["blue"],
        "input-cursor-background": c["accent"],
        "input-selection-background": c["surface1"],
        "scrollbar": c["surface1"],
        "scrollbar-hover": c["dim"],
        "scrollbar-active": c["accent"],
        "scrollbar-background": c["bg"],
        "scrollbar-background-hover": c["bg"],
        "scrollbar-background-active": c["bg"],
        "scrollbar-corner-color": c["bg"],
    }


def textual_theme(name: str, p: Palette) -> Theme:
    c = {role: _css(p, getattr(p, role)) for role in ROLES}
    return Theme(
        name=name, primary=c["accent"], secondary=c["blue"], accent=c["accent"], foreground=c["fg"],
        background=c["bg"], surface=c["bg_alt"], panel=c["surface0"], boost=c["surface0"], success=c["green"],
        warning=c["yellow"], error=c["red"], dark=p.dark, ansi=p.ansi, variables=css_variables(p),
    )


def swatch(p: Palette) -> Text:
    """A palette's colours as small blocks, drawn in its real colours (for the theme picker)."""
    t = Text()
    for c in (p.accent, p.red, p.yellow, p.green, p.cyan, p.blue, p.purple):
        t.append("■ ", style=c)
    return t


THEMES = [textual_theme(name, p) for name, p in PALETTES.items()]
THEME_VARIABLES = css_variables(PALETTES[DEFAULT_THEME])

# ------------------------------------------------------------------------------------------------ role colours

# Placeholder colours, one per role, that no real palette uses. PaletteFilter paints them in the active theme.
_TOKENS = {role: f"#0100{i + 1:02x}" for i, role in enumerate(ROLES)}

BG, BG_ALT, ON_FILL = _TOKENS["bg"], _TOKENS["bg_alt"], _TOKENS["on_fill"]
SURFACE0, SURFACE1 = _TOKENS["surface0"], _TOKENS["surface1"]
FG, SUB, DIM = _TOKENS["fg"], _TOKENS["sub"], _TOKENS["dim"]
ACCENT = _TOKENS["accent"]
RED, GREEN, YELLOW, ORANGE = _TOKENS["red"], _TOKENS["green"], _TOKENS["yellow"], _TOKENS["orange"]
BLUE, CYAN, PURPLE = _TOKENS["blue"], _TOKENS["cyan"], _TOKENS["purple"]
KEY = BLUE  # keys and links


class PaletteFilter(LineFilter):
    """Swaps the role placeholders in rendered lines for the colours of one palette (a new filter per theme)."""

    def __init__(self, p: Palette) -> None:
        super().__init__()
        self.palette = p
        self._map = {Color.parse(token).rgb: RichColor.parse(getattr(p, role)) for role, token in _TOKENS.items()}

    @lru_cache(1024)  # noqa: B019 - one filter per theme switch; the cache dies with it
    def _style(self, style: Style) -> Style:
        color, bgcolor = style.color, style.bgcolor
        new_color = self._map.get(color.triplet) if color is not None and color.triplet else None
        new_bg = self._map.get(bgcolor.triplet) if bgcolor is not None and bgcolor.triplet else None
        if new_color is None and new_bg is None:
            return style
        return style + Style.from_color(new_color or color, new_bg or bgcolor)

    def apply(self, segments: list[Segment], background: Color) -> list[Segment]:
        restyle = self._style
        return [Segment(text, restyle(style), control) if style is not None else Segment(text, style, control)
                for text, style, control in segments]


# ------------------------------------------------------------------------------------------------ glyphs

# Section states shown in the sidebar and the pipeline strip.
TODO, RUNNING, DONE, FAILED = "todo", "running", "done", "failed"
GLYPH = {TODO: "○", RUNNING: "◐", DONE: "●", FAILED: "✗"}
STATE_COLOR = {TODO: DIM, RUNNING: YELLOW, DONE: GREEN, FAILED: RED}
GLYPH_STYLE = {TODO: DIM, RUNNING: f"bold {YELLOW}", DONE: GREEN, FAILED: f"bold {RED}"}
SPINNER = "◐◓◑◒"
DOT = "●"
BLOCK = "▮"
SEP = "  ·  "

# ------------------------------------------------------------------------------------------------ value colours

SEVERITY_STYLE = {"blocking": RED, "warn": YELLOW, "info": CYAN}
ACTION_STYLE = {"create_folder": CYAN, "copy": GREEN, "skip": DIM, "conflict": YELLOW, "excluded": PURPLE}
ACTION_LABEL = {"create_folder": "create", "copy": "copy", "skip": "skip", "conflict": "conflict",
                "excluded": "excluded"}
VERIFY_STYLE = {"ok": GREEN, "missing": RED, "size_differs": YELLOW, "source_newer": ORANGE,
                "only_in_destination": PURPLE}
VERIFY_LABEL = {"ok": "ok", "missing": "missing", "size_differs": "size differs", "source_newer": "source newer",
                "only_in_destination": "only in destination"}
LEVEL_STYLE = {"debug": DIM, "info": FG, "ok": GREEN, "warn": YELLOW, "error": f"bold {RED}"}
MODE_STYLE = {"audit": BLUE, "dryrun": CYAN, "migrate": GREEN, "verify": PURPLE, "lock": ORANGE}


# ------------------------------------------------------------------------------------------------ formatters


def size(n: float | None) -> str:
    return "" if n is None else fmt_size(n)


def duration(seconds: float | None) -> str:
    return "--:--:--" if seconds is None else fmt_duration(seconds)


def short_site(url: str) -> str:
    """``https://contoso.sharepoint.com/sites/Finance`` -> ``Finance``."""
    return url.rstrip("/").rsplit("/", 1)[-1] if url else ""


# ------------------------------------------------------------------------------------------------ renderers


def pill(text: str, color: str, *, fg: str = ON_FILL, bold: bool = True) -> Text:
    """`` text `` as a filled pill: ``fg`` on ``color``. The fill is a span, so text appended later is plain."""
    return Text.assemble((f" {text} ", f"{'bold ' if bold else ''}{fg} on {color}"))


def chip(key: str) -> Text:
    """A key as a chip: `` a `` in bold text on surface1."""
    return Text.assemble((f" {key} ", f"bold {FG} on {SURFACE1}"))


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
    """The summary row every screen uses: ``● 3 blocking   ● 1 warn   ● 0 info``. The dot takes the value colour,
    the count is bold, the label is secondary; zero counts are dimmed as a whole."""
    t = Text(no_wrap=True, overflow="ellipsis")
    for i, (k, n) in enumerate(counts.items()):
        if i:
            t.append("   ")
        label = (labels or {}).get(k, k)
        if n:
            t.append(f"{DOT} ", style=styles.get(k, SUB)).append(f"{n:,}", style=f"bold {FG}")
            t.append(f" {label}", style=SUB)
        else:
            t.append(f"{DOT} 0 {label}", style=DIM)
    return t


def facets(line: Text, pairs: Iterable[tuple[str, str]]) -> Text:
    """Append the active filters to a summary row: ``   │   f all   / invoice``. Pairs with no value are left
    out."""
    shown = [(k, v) for k, v in pairs if v]
    if shown:
        line.append("   │  ", style=DIM)
        for k, v in shown:
            line.append(f" {k} ", style=SUB).append(v, style=f"bold {ACCENT}")
    return line


def notice(color: str, *parts: tuple[str, str] | str) -> Text:
    """A one-line notice under a summary row: a bar in ``color``, then the text. Never wraps."""
    t = Text("▌ ", style=color, no_wrap=True, overflow="ellipsis")
    for part in parts:
        if isinstance(part, str):
            t.append(part, style=SUB)
        else:
            t.append(*part)
    return t


def plural(n: int, word: str, many: str | None = None) -> str:
    """``1 run``, ``2 runs``, ``1,204 entries`` (pass ``many`` for irregular plurals)."""
    return f"{n:,} {word if n == 1 else (many or word + 's')}"


def state_word(status: str, words: dict[str, str] | None = None) -> Text:
    w = (words or {}).get(status, {TODO: "todo", RUNNING: "running", DONE: "done", FAILED: "failed"}[status])
    return Text(w, style=STATE_COLOR[status])


def empty_text(title: str, body: str, keys: Iterable[tuple[str, str]] = ()) -> Text:
    """Centred empty state: a bold line, a hint in subtext, then key chips."""
    t = Text(justify="center")
    t.append(title, style=f"bold {FG}").append("\n\n").append(body, style=SUB)
    keys = list(keys)
    if keys:
        t.append("\n\n").append_text(key_hints(keys, sep="    "))
    return t


def log_text(ev: LogLine) -> Table:
    """``12:01:02 warn  text`` coloured by level; long messages wrap under the message column, not under the
    time. User data is never parsed as markup."""
    style = LEVEL_STYLE.get(ev.level, FG)
    grid = Table.grid(padding=(0, 1, 0, 0))
    grid.add_column(no_wrap=True, width=8)
    grid.add_column(no_wrap=True, width=5)
    grid.add_column(ratio=1, overflow="fold")
    grid.add_row(Text(f"{ev.at:%H:%M:%S}", style=DIM), Text(ev.level, style=style), Text(ev.text, style=style))
    return grid
