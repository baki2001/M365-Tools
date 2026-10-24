"""Colour themes: every palette is complete and valid, role colours are repainted per theme, and the ctrl+t
picker previews, reverts on esc and saves the choice to the config file."""

from __future__ import annotations

import tomllib
from pathlib import Path

from rich.color import Color as RichColor
from rich.segment import Segment
from rich.style import Style
from textual.color import Color

from m365tools.core.config import load_config, save_config
from m365tools.tui import theme
from m365tools.tui.app import detect_truecolor
from m365tools.tui.modals import PickerModal

from .test_tui import make_app, make_cfg


def test_every_palette_parses_and_never_uses_a_role_placeholder():
    placeholders = {Color.parse(getattr(theme, name)).rgb for name in ("BG", "FG", "SUB", "DIM", "ACCENT", "RED")}
    for name, p in theme.PALETTES.items():
        for role in theme.ROLES:
            value = getattr(p, role)
            Color.parse(f"ansi_{value}" if p.ansi else value)  # raises on a typo
            if not p.ansi:
                assert Color.parse(value).rgb not in placeholders, (name, role)
        variables = theme.css_variables(p)
        assert {"bg", "bg-alt", "on-fill", "fg-sub", "fg-dim", "orange", "cyan", "purple"} <= set(variables)
    assert theme.DEFAULT_THEME in theme.PALETTES


def test_palette_filter_repaints_role_colours_only():
    nord = theme.PALETTES["nord"]
    f = theme.PaletteFilter(nord)
    own = Style(color="#123456")
    segments = [Segment("a", Style.parse(f"bold {theme.RED} on {theme.SURFACE1}")), Segment("b", own),
                Segment("c", None)]
    out = f.apply(segments, Color.parse("#000000"))
    assert out[0].style.color == RichColor.parse(nord.red)
    assert out[0].style.bgcolor == RichColor.parse(nord.surface1)
    assert out[0].style.bold
    assert out[1].style == own and out[2].style is None


def test_default_theme_is_not_written_and_a_choice_round_trips(tmp_path: Path):
    path = tmp_path / "config.toml"
    cfg = make_cfg()
    save_config(cfg, path)
    assert "ui" not in tomllib.loads(path.read_text(encoding="utf-8"))
    cfg.ui.theme = "gruvbox-dark"
    save_config(cfg, path)
    assert tomllib.loads(path.read_text(encoding="utf-8"))["ui"] == {"theme": "gruvbox-dark"}
    assert load_config(path).ui.theme == "gruvbox-dark"


async def test_app_starts_with_the_configured_theme_and_falls_back_on_unknown_names():
    app = make_app()
    app.cfg.ui.theme = "dracula"
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        assert app.theme == "dracula" and app._palette_filter.palette is theme.PALETTES["dracula"]
    app = make_app()
    app.cfg.ui.theme = "no-such-theme"
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        assert app.theme == theme.DEFAULT_THEME


async def test_picker_previews_reverts_on_esc_and_saves_on_enter(tmp_path: Path):
    app = make_app()
    app.cfg.path = tmp_path / "config.toml"
    names = list(theme.PALETTES)
    async with app.run_test(size=(140, 45)) as pilot:
        await pilot.pause()
        await pilot.press("ctrl+t")
        await pilot.pause()
        assert isinstance(app.screen, PickerModal)
        await pilot.press("down")
        await pilot.pause()
        assert app.theme == names[1]  # live preview
        await pilot.press("escape")
        await pilot.pause()
        assert app.theme == theme.DEFAULT_THEME
        assert not app.cfg.path.exists()

        await pilot.press("ctrl+t")
        await pilot.pause()
        await pilot.press(*"nord", "enter")
        await pilot.pause()
        assert app.theme == "nord"
        assert load_config(app.cfg.path).ui.theme == "nord"


def test_windows_terminal_gets_truecolor_unless_colorterm_is_set():
    env = {"WT_SESSION": "abc", "TERM": "xterm-256color"}
    detect_truecolor(env)
    assert env["COLORTERM"] == "truecolor"
    env = {"WT_SESSION": "abc", "COLORTERM": "256"}
    detect_truecolor(env)
    assert env["COLORTERM"] == "256"
    env = {"TERM": "xterm-256color"}
    detect_truecolor(env)
    assert "COLORTERM" not in env
