"""Shared UI design tokens for CustomTkinter tabs."""

# Spacing
PAD_OUTER = 10
PAD_INNER = 14
GAP_SM = 6
GAP_MD = 10
PADY_SM = 6
PADY_MD = 10

# Typography
FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_SECTION = ("Segoe UI", 14, "bold")
FONT_BODY = ("Segoe UI", 13)
FONT_LABEL = ("Segoe UI", 13, "bold")
FONT_BUTTON = ("Segoe UI", 13, "bold")
FONT_BTN_PRIMARY = ("Segoe UI", 13, "bold")
FONT_BTN = ("Segoe UI", 12)
FONT_BTN_TOOL = ("Segoe UI", 11)

# Sizing
H_PRIMARY = 46
H_SECONDARY = 36
H_TOOL = 30

# Colors
LABEL_COLOR = {"text_color": ("#111111", "#eaeaea")}
CARD_BORDER_COLOR = ("#d0d0d0", "#3a3a3a")
CARD_FG_COLOR = ("#ffffff", "#1f1f1f")

BTN_PRIMARY = {
    "fg_color": ("#2563eb", "#2563eb"),
    "hover_color": ("#1d4ed8", "#1d4ed8"),
    "text_color": "white",
}

BTN_SECONDARY = {
    "fg_color": ("#e5e7eb", "#2a2a2a"),
    "hover_color": ("#d1d5db", "#3a3a3a"),
    "text_color": ("#111111", "#eaeaea"),
}

BTN_TOOL = {
    "fg_color": "transparent",
    "hover_color": ("#f3f4f6", "#2b2b2b"),
    "border_width": 1,
    "border_color": ("#cfcfcf", "#3a3a3a"),
    "text_color": ("#111111", "#eaeaea"),
}

CHECKBOX_STYLE = {
    "fg_color": ("#2563eb", "#2563eb"),
    "hover_color": ("#1d4ed8", "#1d4ed8"),
    "border_color": ("#9ca3af", "#4b5563"),
    "checkmark_color": ("#ffffff", "#ffffff"),
    "text_color": ("#111111", "#eaeaea"),
}
