from __future__ import annotations

import os
import sys
import tkinter as tk
from dataclasses import dataclass
from tkinter import ttk

try:
    import ttkbootstrap as ttkb  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    ttkb = None


@dataclass(frozen=True)
class ThemeContext:
    root: tk.Misc
    style: object
    using_bootstrap: bool
    colors: dict[str, str]
    font_family: str


THEME_NAME = "flatly"


def _detect_font_family() -> str:
    if sys.platform.startswith("win"):
        return "Segoe UI"
    if sys.platform.startswith("darwin"):
        return "Helvetica Neue"
    return "Noto Sans CJK SC"


def _build_color_palette() -> dict[str, str]:
    return {
        "bg": "#f5f7fb",
        "card": "#ffffff",
        "card_soft": "#f8fafc",
        "border": "#d8e0ea",
        "primary": "#2563eb",
        "primary_hover": "#1d4ed8",
        "secondary": "#64748b",
        "secondary_hover": "#475569",
        "success": "#16a34a",
        "success_hover": "#15803d",
        "warning": "#d97706",
        "warning_hover": "#b45309",
        "danger": "#dc2626",
        "danger_hover": "#b91c1c",
        "text": "#0f172a",
        "muted": "#64748b",
        "accent": "#eff6ff",
    }


def _configure_styles(style: ttk.Style, colors: dict[str, str], font_family: str) -> None:
    base_font = (font_family, 10)
    title_font = (font_family, 16, "bold")
    subtitle_font = (font_family, 10)
    section_font = (font_family, 11, "bold")

    style.configure(".", background=colors["bg"], foreground=colors["text"], font=base_font)
    style.configure("App.TFrame", background=colors["bg"])
    style.configure("Card.TFrame", background=colors["card"], relief="flat")
    style.configure("SoftCard.TFrame", background=colors["card_soft"], relief="flat")
    style.configure("Header.TLabel", background=colors["bg"], foreground=colors["text"], font=title_font)
    style.configure("SubHeader.TLabel", background=colors["bg"], foreground=colors["muted"], font=subtitle_font)
    style.configure("Section.TLabel", background=colors["card"], foreground=colors["text"], font=section_font)
    style.configure("Body.TLabel", background=colors["card"], foreground=colors["text"], font=base_font)
    style.configure("Hint.TLabel", background=colors["card"], foreground=colors["muted"], font=base_font)
    style.configure("Value.TLabel", background=colors["card"], foreground=colors["text"], font=(font_family, 11))
    style.configure("AccentValue.TLabel", background=colors["accent"], foreground=colors["primary"], font=(font_family, 11, "bold"))
    style.configure("Status.TLabel", background=colors["card_soft"], foreground=colors["primary"], font=(font_family, 10, "bold"))
    style.configure("TSeparator", background=colors["border"])

    button_common = {
        "font": (font_family, 10, "bold"),
        "padding": (14, 8),
        "borderwidth": 0,
        "relief": "flat",
    }
    style.configure("Primary.TButton", background=colors["primary"], foreground="#ffffff", **button_common)
    style.map(
        "Primary.TButton",
        background=[("active", colors["primary_hover"]), ("disabled", "#93c5fd")],
        foreground=[("disabled", "#e2e8f0")],
    )
    style.configure("Secondary.TButton", background=colors["secondary"], foreground="#ffffff", **button_common)
    style.map(
        "Secondary.TButton",
        background=[("active", colors["secondary_hover"]), ("disabled", "#cbd5e1")],
        foreground=[("disabled", "#e2e8f0")],
    )
    style.configure("Success.TButton", background=colors["success"], foreground="#ffffff", **button_common)
    style.map(
        "Success.TButton",
        background=[("active", colors["success_hover"]), ("disabled", "#86efac")],
        foreground=[("disabled", "#ecfdf5")],
    )
    style.configure("Warning.TButton", background=colors["warning"], foreground="#ffffff", **button_common)
    style.map(
        "Warning.TButton",
        background=[("active", colors["warning_hover"]), ("disabled", "#fcd34d")],
        foreground=[("disabled", "#fffbeb")],
    )
    style.configure("Danger.TButton", background=colors["danger"], foreground="#ffffff", **button_common)
    style.map(
        "Danger.TButton",
        background=[("active", colors["danger_hover"]), ("disabled", "#fca5a5")],
        foreground=[("disabled", "#fef2f2")],
    )

    style.configure("TEntry", padding=6)
    style.configure("TCombobox", padding=4)
    style.configure("TLabelframe", background=colors["card"], foreground=colors["text"])
    style.configure("TLabelframe.Label", background=colors["card"], foreground=colors["text"], font=(font_family, 10, "bold"))
    style.configure("Treeview", rowheight=26, font=base_font)
    style.configure("Treeview.Heading", font=(font_family, 10, "bold"))


def create_app_root(title: str, geometry: str = "900x680", minsize: tuple[int, int] = (720, 520)) -> ThemeContext:
    colors = _build_color_palette()
    font_family = _detect_font_family()

    if ttkb is not None:
        root = ttkb.Window(themename=THEME_NAME)
        style = root.style
        using_bootstrap = True
    else:
        root = tk.Tk()
        style = ttk.Style(root)
        using_bootstrap = False
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

    _configure_styles(style, colors, font_family)

    root.title(title)
    root.geometry(geometry)
    root.minsize(*minsize)
    try:
        root.configure(background=colors["bg"])
    except Exception:
        pass

    _set_window_icon(root)
    return ThemeContext(root=root, style=style, using_bootstrap=using_bootstrap, colors=colors, font_family=font_family)


def _set_window_icon(root: tk.Misc) -> None:
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "..", "icon.ico")
        icon_path = os.path.abspath(icon_path)
        if os.path.exists(icon_path) and sys.platform.startswith("win"):
            root.iconbitmap(icon_path)
    except Exception:
        pass
