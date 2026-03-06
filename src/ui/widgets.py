"""Reusable UI helpers for tabs built with CustomTkinter."""

from tkinter import ttk

import customtkinter as ctk

from ui.theme import (
    BTN_PRIMARY,
    BTN_SECONDARY,
    BTN_TOOL,
    CARD_BORDER_COLOR,
    CARD_FG_COLOR,
    FONT_BTN,
    FONT_BTN_PRIMARY,
    FONT_BTN_TOOL,
    FONT_SECTION,
    GAP_MD,
    H_PRIMARY,
    H_SECONDARY,
    H_TOOL,
    LABEL_COLOR,
    PAD_INNER,
)


def apply_ttk_styles() -> None:
    """Configure ttk styles used by mixed CTk/ttk screens."""
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure(
        "Enabled.TCombobox",
        fieldbackground="white",
        background="white",
        foreground="black",
    )
    style.map(
        "Enabled.TCombobox",
        fieldbackground=[("readonly", "white"), ("!disabled", "white")],
        foreground=[("readonly", "black"), ("!disabled", "black")],
    )

    style.configure(
        "Disabled.TCombobox",
        fieldbackground="#e6e6e6",
        background="#e6e6e6",
        foreground="#7a7a7a",
    )
    style.map(
        "Disabled.TCombobox",
        fieldbackground=[("disabled", "#e6e6e6")],
        foreground=[("disabled", "#7a7a7a")],
    )

    style.configure("Treeview", rowheight=24)
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure("TCombobox", padding=4)


def make_section_card(parent, title: str) -> ctk.CTkFrame:
    """Create a consistent card section with title and body frame."""
    card = ctk.CTkFrame(
        parent,
        corner_radius=12,
        border_width=1,
        border_color=CARD_BORDER_COLOR,
        fg_color=CARD_FG_COLOR,
    )
    card.pack(fill="x", pady=(0, GAP_MD))

    header = ctk.CTkFrame(card, fg_color="transparent")
    header.pack(fill="x", padx=PAD_INNER, pady=(PAD_INNER, 6))
    ctk.CTkLabel(header, text=title, font=FONT_SECTION, **LABEL_COLOR).pack(anchor="w")

    body = ctk.CTkFrame(card, fg_color="transparent")
    body.pack(fill="x", padx=PAD_INNER, pady=(0, PAD_INNER))
    return body


def make_primary_button(parent, *, text: str, command, **kwargs) -> ctk.CTkButton:
    opts = dict(BTN_PRIMARY)
    opts.update(
        {
            "text": text,
            "command": command,
            "font": FONT_BTN_PRIMARY,
            "height": H_PRIMARY,
        }
    )
    opts.update(kwargs)
    return ctk.CTkButton(parent, **opts)


def make_secondary_button(parent, *, text: str, command, **kwargs) -> ctk.CTkButton:
    opts = dict(BTN_SECONDARY)
    opts.update(
        {
            "text": text,
            "command": command,
            "font": FONT_BTN,
            "height": H_SECONDARY,
        }
    )
    opts.update(kwargs)
    return ctk.CTkButton(parent, **opts)


def make_tool_button(parent, *, text: str, command, **kwargs) -> ctk.CTkButton:
    opts = dict(BTN_TOOL)
    opts.update(
        {
            "text": text,
            "command": command,
            "font": FONT_BTN_TOOL,
            "height": H_TOOL,
        }
    )
    opts.update(kwargs)
    return ctk.CTkButton(parent, **opts)
