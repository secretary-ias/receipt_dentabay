# app/theme.py
from __future__ import annotations
import tkinter as tk
from tkinter import ttk

# Web design tokens -> approximated HEX for Tk (from src/index.css)
# primary: hsl(193 85% 42%) , secondary: hsl(210 30% 92%), border: hsl(214 25% 88%)
# background: hsl(210 40% 98%), foreground: hsl(215 25% 15%)
# accent: hsl(187 80% 55%)
# Dark mode not implemented in Tk; we stick to light palette. :contentReference[oaicite:3]{index=3}

TOKENS = {
    "bg":      "#F4F8FB",   # ~ hsl(210,40%,98%)
    "fg":      "#2D3340",   # ~ hsl(215,25%,15%)
    "card":    "#FFFFFF",   # card background
    "border":  "#D7DFE7",   # ~ hsl(214,25%,88%)
    "muted":   "#EEF3F7",   # ~ hsl(210,30%,94%)
    "primary": "#14A3C7",   # ~ hsl(193,85%,42%)
    "primary_fg": "#FFFFFF",
    "accent":  "#1DC4B0",   # ~ hsl(187,80%,55%)
    "danger":  "#E03131",
}

def style_app(root: tk.Tk) -> None:
    root.configure(bg=TOKENS["bg"])
    style = ttk.Style(root)
    # Use 'clam' as a base to allow colors
    style.theme_use("clam")

    # Base fonts
    style.configure(".", font=("Segoe UI", 10), foreground=TOKENS["fg"], background=TOKENS["bg"])

    # Notebook (tabs "Receipts" / "Settings")
    style.configure("TNotebook", background=TOKENS["bg"], borderwidth=0)
    style.configure("TNotebook.Tab", padding=(14, 8), background=TOKENS["muted"], foreground=TOKENS["fg"])
    style.map("TNotebook.Tab",
              background=[("selected", TOKENS["card"])],
              foreground=[("selected", TOKENS["fg"])],
              bordercolor=[("selected", TOKENS["border"])])

    # Frames as "cards"
    style.configure("Card.TLabelframe", background=TOKENS["card"], bordercolor=TOKENS["border"], relief="solid",
                    borderwidth=1)
    style.configure("Card.TLabelframe.Label", background=TOKENS["card"], foreground="#556070", font=("Segoe UI", 10, "bold"))

    # Buttons
    style.configure("Primary.TButton", background=TOKENS["primary"], foreground=TOKENS["primary_fg"],
                    borderwidth=0, padding=(12, 6))
    style.map("Primary.TButton",
              background=[("active", "#1296B8"), ("disabled", "#7DCFE3")])

    style.configure("Ghost.TButton", background=TOKENS["muted"], foreground=TOKENS["fg"],
                    bordercolor=TOKENS["border"], borderwidth=1, padding=(10, 6))
    style.map("Ghost.TButton",
              background=[("active", "#E6EDF3")])

    # Entries
    style.configure("TEntry", fieldbackground="#FFFFFF", padding=6, bordercolor=TOKENS["border"], borderwidth=1)
    style.map("TEntry", bordercolor=[("focus", TOKENS["primary"])])

    # Combobox / Text
    style.configure("TCombobox", fieldbackground="#FFFFFF", padding=6, bordercolor=TOKENS["border"])
    style.configure("TText", bordercolor=TOKENS["border"], borderwidth=1)

    # Treeviews (tables) to match web tables on Receipts & Settings pages. :contentReference[oaicite:4]{index=4}
    style.configure("Treeview",
                    background="#FFFFFF",
                    fieldbackground="#FFFFFF",
                    bordercolor=TOKENS["border"],
                    rowheight=28)
    style.configure("Treeview.Heading",
                    background=TOKENS["muted"], relief="flat",
                    foreground="#536175", font=("Segoe UI Semibold", 10))
    style.map("Treeview.Heading", background=[("active", TOKENS["muted"])])
    style.map("Treeview", background=[("selected", "#E7F6FB")], foreground=[("selected", "#0F7491")])

    # Status label
    style.configure("Status.TLabel", background=TOKENS["bg"], foreground="#667085")

def card(frame: ttk.Frame, text: str) -> ttk.Labelframe:
    return ttk.Labelframe(frame, text=text, style="Card.TLabelframe", padding=(10, 8))
