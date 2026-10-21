"""Modal dialogs: sign-in prompt, confirmation, read-only text."""

from __future__ import annotations

from typing import ClassVar

from rich.text import Text
from textual.app import ComposeResult
from textual.binding import Binding, BindingType
from textual.containers import Horizontal, Vertical, VerticalScroll
from textual.screen import ModalScreen
from textual.widgets import Button, Label, Static

from m365tools.core.events import AuthPrompt


def _auth_text(ev: AuthPrompt) -> Text:
    t = Text(ev.message)
    if ev.url:
        t.append("\n\nURL:  ", style="bold").append(ev.url, style="bold underline cyan")
    if ev.code:
        t.append("\nCode: ", style="bold").append(f" {ev.code} ", style="bold reverse green")
    return t


class AuthPromptModal(ModalScreen[None]):
    """Shows browser / device-code sign-in instructions. Closes itself once the job makes progress."""

    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape", "dismiss", "Close")]

    def __init__(self, ev: AuthPrompt) -> None:
        super().__init__()
        self.ev = ev

    def compose(self) -> ComposeResult:
        with Vertical(classes="dialog"):
            yield Label(Text(f"Sign in to tenant '{self.ev.tenant}'"), classes="dialog-title")
            yield Static(_auth_text(self.ev), id="auth-body")
            if self.ev.code:
                yield Label("The code has been copied to the clipboard (if your terminal allows it).",
                            classes="hint")
            with Horizontal(classes="buttons"):
                yield Button("Close", id="close", variant="primary")

    def on_mount(self) -> None:
        if self.ev.code:
            self.app.copy_to_clipboard(self.ev.code)

    def show(self, ev: AuthPrompt) -> None:
        self.ev = ev
        self.query_one("#auth-body", Static).update(_auth_text(ev))
        if ev.code:
            self.app.copy_to_clipboard(ev.code)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        event.stop()
        self.dismiss()


class ConfirmModal(ModalScreen[bool]):
    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape", "cancel", "Cancel"), Binding("y", "confirm", "Yes")]

    def __init__(self, title: str, body: str, confirm_label: str = "Yes", danger: bool = True) -> None:
        super().__init__()
        self.title_text = title
        self.body = body
        self.confirm_label = confirm_label
        self.danger = danger

    def compose(self) -> ComposeResult:
        with Vertical(classes="dialog"):
            yield Label(Text(self.title_text), classes="dialog-title")
            yield Static(Text(self.body))
            with Horizontal(classes="buttons"):
                yield Button(self.confirm_label, id="yes", variant="error" if self.danger else "primary")
                yield Button("Cancel", id="no")

    def on_button_pressed(self, event: Button.Pressed) -> None:
        event.stop()
        self.dismiss(event.button.id == "yes")

    def action_confirm(self) -> None:
        self.dismiss(True)

    def action_cancel(self) -> None:
        self.dismiss(False)


class TextModal(ModalScreen[None]):
    BINDINGS: ClassVar[list[BindingType]] = [Binding("escape", "dismiss", "Close")]

    def __init__(self, title: str, text: str) -> None:
        super().__init__()
        self.title_text = title
        self.text = text

    def compose(self) -> ComposeResult:
        with Vertical(classes="dialog wide"):
            yield Label(self.title_text, classes="dialog-title")
            with VerticalScroll():
                yield Static(Text(self.text))
            with Horizontal(classes="buttons"):
                yield Button("Close", id="close", variant="primary")

    def on_button_pressed(self, event: Button.Pressed) -> None:
        event.stop()
        self.dismiss()
