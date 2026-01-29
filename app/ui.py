# app/ui.py
from __future__ import annotations

import calendar as pycal
import datetime as dt
import os
import re
import smtplib
import subprocess
import sys
import tkinter as tk
from dataclasses import dataclass, replace
from email.message import EmailMessage
from pathlib import Path
from tkinter import ttk, messagebox, filedialog, simpledialog
import webbrowser
from urllib.parse import quote as urlquote
from typing import List, Optional, Sequence, Tuple

from .config import ConfigManager
from .database import (
    ClinicDatabase,
    Receipt,
    ReceiptSummary,
    ReceiptItem,
    ReceiptDraftItem,
    PartialPayment,
    Patient,
    AppointmentDetail,
    DentalChartItem,
    MedicReportSummary,
    MedicReportDetail,
)
from .receipt import ReceiptPDFGenerator, PaymentEntry, PaymentProgress
from .theme import style_app, card

CURRENCY = "RM"

def fmt_money(value: float | int | None) -> str:
    v = float(value or 0.0)
    return f"{CURRENCY} {v:,.2f}"

def _rtf_to_text(rtf: str) -> str:
    txt = rtf.replace("\r\n", "\n")
    txt = re.sub(
        r"\\'([0-9a-fA-F]{2})",
        lambda m: bytes([int(m.group(1), 16)]).decode("latin1", "ignore"),
        txt,
    )
    txt = re.sub(r"\\pard", "\n", txt)
    txt = re.sub(r"\\par[d]?", "\n", txt)
    txt = re.sub(r"\\line", "\n", txt)
    txt = re.sub(r"\\tab", "\t", txt)
    txt = re.sub(r"\\[a-zA-Z]+\d*-?", "", txt)
    txt = re.sub(r"\\[{}\\]", "", txt)
    txt = re.sub(r"[{}]", "", txt)
    txt = re.sub(r"\n[ \t]*\n+", "\n\n", txt)
    return txt.strip()


def _text_to_rtf(text: str) -> str:
    if not text:
        return ""
    txt = text.replace("\r\n", "\n")
    parts: list[str] = ["{\\rtf1\\ansi"]
    for line in txt.split("\n"):
        escaped = line.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")
        parts.append(escaped + "\\par")
    parts.append("}")
    return "\n".join(parts)

def _clean_note_text(value: str | None) -> str:
    if not value:
        return ""
    text = value.replace("\r\n", "\n").strip()
    if not text:
        return ""
    # Strip leading font declarations like "Arial;Arial Rounded MT;"
    text = re.sub(r"^[^;]+;[^;]+;?\s*", "", text)
    is_rtf = text.lstrip().startswith("{\\rtf") or "\\par" in text or "\\pard" in text or "\\cf" in text or "\\lang" in text
    if is_rtf:
        text = re.sub(r"\{\\\*?\\fonttbl.*?\}", "", text, flags=re.DOTALL)
        text = re.sub(r"\{\\colortbl.*?\}", "", text, flags=re.DOTALL)
        text = re.sub(r"\{\\\*?\\generator.*?\}", "", text, flags=re.DOTALL)
        text = re.sub(r"\{\\info.*?\}", "", text, flags=re.DOTALL)
        text = re.sub(r"\{\\stylesheet.*?\}", "", text, flags=re.DOTALL)
        text = _rtf_to_text(text)
    # Normalise whitespace and drop empty or decorative lines.
    lines = [line.strip() for line in text.splitlines()]
    cleaned_lines: list[str] = []
    for line in lines:
        if not line:
            continue
        if re.fullmatch(r"[A-Za-z0-9 .,'()/\-]+;[A-Za-z0-9 .,'()/\-]+;?", line):
            continue
        cleaned_lines.append(line)
    return "\n".join(cleaned_lines)

def parse_date(d: str) -> dt.date | None:
    d = (d or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(d, fmt).date()
        except ValueError:
            continue
    return None

def digits_only(s: str) -> str:
    return "".join(ch for ch in s if ch.isdigit())

def normalize_msisdn_malaysia(s: str) -> str:
    raw = digits_only(s or "")
    if not raw:
        return ""
    if raw.startswith("0"):
        return "6" + raw
    if raw.startswith("60"):
        return raw
    return raw

def _bundle_base_dir() -> Path:
    """
    When frozen with PyInstaller (onefile), resources are extracted to sys._MEIPASS.
    Fall back to this file's directory when running from source.
    """
    try:
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            return Path(sys._MEIPASS)
    except Exception:
        pass
    return Path(__file__).resolve().parent

@dataclass
class _PaymentProgressState:
    sequence: int
    amount: float
    total_due: float
    paid_on: dt.datetime
    balance: float


@dataclass
class VisitNoteItem:
    stock_id: str
    name: str
    category: str
    unit_price: float
    qty: int = 1


@dataclass
class _SettlementItem:
    stock_id: str
    description: str
    qty: int = 1
    unit_price: float = 0.0
    notation_id: int = 0
    tooth_label: str = ""
    remarks: str = ""
    source: str = "note"


@dataclass(eq=False)
class _ReceiptEditableItem:
    stock_id: str
    description: str
    qty: int = 1
    unit_price: float = 0.0
    remark: str = ""


@dataclass
class _CalendarHost:
    grid: ttk.Frame
    month_var: tk.StringVar

class ReceiptApp(tk.Tk):
    def __init__(self, config_path: Path) -> None:
        super().__init__()
        self.withdraw()

        # Theme
        style_app(self)

        # Services
        self.cfg = ConfigManager(config_path)
        self.db: ClinicDatabase | None = None
        self.pdf: ReceiptPDFGenerator | None = None

        # Session
        self.session_user: str | None = None
        self._ready = False
        self.schedule_date = dt.date.today()
        self._calendar_month = self.schedule_date.replace(day=1)
        self._calendar_color_tags: dict[str, str] = {}
        self.schedule_appointments: list[AppointmentDetail] = []
        self.schedule_index: dict[str, AppointmentDetail] = {}
        self._calendar_hosts: list[_CalendarHost] = []
        self.shared_date_var = tk.StringVar(value=self.schedule_date.isoformat())
        self._shared_date_updating = False
        self.settlement_date_var = self.shared_date_var
        self.settlement_index: dict[str, AppointmentDetail] = {}
        self.settlement_items: list[_SettlementItem] = []
        self.settlement_item_map: dict[str, _SettlementItem] = {}
        self.settlement_current_note: MedicReportDetail | None = None
        self.settlement_selected_appointment: Optional[AppointmentDetail] = None
        self.payment_methods: list[tuple[str, str]] = []
        self.payment_method_map: dict[str, str] = {}
        self.settlement_subtotal_var = tk.StringVar(value=fmt_money(0))
        self.settlement_discount_var = tk.DoubleVar(value=0.0)
        self.settlement_rounding_var = tk.DoubleVar(value=0.0)
        self.settlement_total_var = tk.StringVar(value=fmt_money(0))
        self.settlement_payment_var = tk.StringVar()
        self.settlement_message_var = tk.StringVar(value="")
        self.stock_categories: list[str] | None = None
        self.stock_items_cache: dict[str, list[tuple[str, str, float]]] = {}
        self.settlement_receipt_text: str = ""
        self.settlement_receipt_header: str = ""
        self.settlement_current_receipt: Receipt | None = None
        self._prefill_visit_note_items: list[VisitNoteItem] | None = None
        self._prefill_visit_note_patient_id: str | None = None
        self.shared_date_var.trace_add("write", lambda *_: self._on_shared_date_changed())

        if not self._prompt_login():
            self.destroy()
            return

        self.title("Klinik Pergigian Dentabay")
        self._apply_default_size()
        self.deiconify()

        # State
        self.receipt_index: dict[str, ReceiptSummary] = {}
        self.receipt_edit_summary: ReceiptSummary | None = None
        self.receipt_edit_items: list[_ReceiptEditableItem] = []
        self.receipt_item_map: dict[str, int] = {}
        self.receipt_payments: list[PartialPayment] = []
        self.receipt_payment_map: dict[str, int] = {}
        self.selected_payment: _PaymentProgressState | None = None

        # Build UI
        self._build_tabs()
        self._render_calendar()
        self._attach_events()
        self._init_backing_services()

        if self.session_user:
            self.status_var.set(f"Logged in as {self.session_user}")
        else:
            self.status_var.set("Guest session")
        self._ready = True
        self._refresh_shared_views()

    # ---------------- tabs
    def _build_tabs(self) -> None:
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        self.schedule_tab = ttk.Frame(self.nb, padding=12)
        self.settlement_tab = ttk.Frame(self.nb, padding=12)
        self.receipts_tab = ttk.Frame(self.nb, padding=12)
        self.settings_tab = ttk.Frame(self.nb, padding=12)
        self.nb.add(self.schedule_tab, text="Schedule")
        self.nb.add(self.settlement_tab, text="Settlement")
        self.nb.add(self.receipts_tab, text="Receipts")
        self.nb.add(self.settings_tab, text="Settings")

        self._build_schedule_tab()
        self._build_settlement_tab()
        self._build_receipts_tab()
        self._build_settings_tab()

    def _register_calendar_host(self, grid: ttk.Frame, month_var: tk.StringVar) -> None:
        self._calendar_hosts.append(_CalendarHost(grid=grid, month_var=month_var))

    def _create_calendar_section(
        self,
        parent: ttk.Widget,
        *,
        label: str = "Calendar",
        month_var: tk.StringVar | None = None,
        show_selected: bool = False,
    ) -> ttk.Labelframe:
        section = card(parent, label)
        header = ttk.Frame(section)
        header.pack(fill="x", pady=(0, 4))
        ttk.Button(header, text="<", width=3, style="Ghost.TButton", command=lambda: self._change_schedule_month(-1)).pack(
            side="left"
        )
        month_value = month_var or tk.StringVar(value=self._calendar_month.strftime("%B %Y"))
        ttk.Label(header, textvariable=month_value, font=("Segoe UI Semibold", 10)).pack(side="left", expand=True, padx=6)
        ttk.Button(header, text=">", width=3, style="Ghost.TButton", command=lambda: self._change_schedule_month(1)).pack(
            side="right"
        )
        grid = ttk.Frame(section)
        grid.pack(fill="x", expand=False, pady=(4, 0))
        for idx in range(7):
            grid.columnconfigure(idx, weight=1)
        self._register_calendar_host(grid, month_value)
        if show_selected:
            ttk.Label(section, textvariable=self.schedule_selected_var, foreground="#475467").pack(
                anchor="w", pady=(6, 0)
            )
        return section

    def _set_shared_date(self, new_date: dt.date, *, refresh: bool = True) -> None:
        if not new_date:
            return
        if self.schedule_date == new_date and not refresh:
            return
        self.schedule_date = new_date
        self._calendar_month = new_date.replace(day=1)
        iso_value = new_date.isoformat()
        if self.shared_date_var.get() != iso_value:
            self._shared_date_updating = True
            self.shared_date_var.set(iso_value)
            self._shared_date_updating = False
        if refresh:
            self._render_calendar()
            self._refresh_shared_views()

    def _on_shared_date_changed(self) -> None:
        if self._shared_date_updating:
            return
        new_date = parse_date(self.shared_date_var.get())
        if not new_date:
            return
        self._set_shared_date(new_date, refresh=True)

    def _refresh_shared_views(self) -> None:
        self._update_schedule_header()
        if self._ready:
            try:
                self._load_schedule_for(self.schedule_date)
            except Exception:
                pass
            try:
                self._load_settlement_list()
            except Exception:
                pass
            if self.db:
                try:
                    if hasattr(self, "_search"):
                        self._search()
                except Exception:
                    pass

    def _render_calendar_into(self, target: ttk.Frame) -> None:
        for child in target.winfo_children():
            child.destroy()
        for idx, name in enumerate(("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")):
            ttk.Label(target, text=name, anchor="center").grid(row=0, column=idx, padx=1, pady=(0, 2))
        cal = pycal.Calendar(firstweekday=0)
        month_dates = cal.monthdatescalendar(self._calendar_month.year, self._calendar_month.month)
        for row_index, week in enumerate(month_dates, start=1):
            for col_index, day in enumerate(week):
                btn_style = "Ghost.TButton"
                is_current_month = day.month == self._calendar_month.month
                is_selected = day == self.schedule_date
                if is_selected:
                    btn_style = "Primary.TButton"
                btn = ttk.Button(target, text=str(day.day), width=3, style=btn_style)
                btn.grid(row=row_index, column=col_index, padx=1, pady=1, sticky="nsew")
                if not is_current_month:
                    btn.state(["disabled"])
                else:
                    btn.configure(command=lambda d=day: self._on_calendar_day(d))

        def _on_tab_changed(_evt=None):
            try:
                tab_text = self.nb.tab(self.nb.select(), "text")
            except Exception:
                tab_text = ""
            if "Schedule" in tab_text:
                try:
                    self._render_calendar()
                    self._load_schedule_for(self.schedule_date)
                except Exception:
                    pass
            elif "Receipts" in tab_text:
                try:
                    if hasattr(self, "_search"):
                        self._search()
                except Exception:
                    pass
            elif "Settlement" in tab_text:
                try:
                    if hasattr(self, "_load_settlement_list"):
                        self._load_settlement_list()
                except Exception:
                    pass

        self.nb.bind("<<NotebookTabChanged>>", _on_tab_changed)



        self.status_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self.status_var, anchor="w", style="Status.TLabel").pack(fill="x", pady=(6, 8))

    def _apply_default_size(self) -> None:
        try:
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            target_w = int(screen_w * 0.8)
            target_h = int(screen_h * 0.8)
            self.geometry(f"{target_w}x{target_h}")
            self.minsize(int(screen_w * 0.6), int(screen_h * 0.6))
        except Exception:
            self.geometry("1280x800")
            self.minsize(1080, 720)

    # ---------------- Schedule tab
    def _build_schedule_tab(self) -> None:
        container = ttk.Frame(self.schedule_tab)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(0, weight=1)

        calendar_column = ttk.Frame(container)
        calendar_column.grid(row=0, column=0, sticky="nw", padx=(0, 16))
        self.schedule_month_var = tk.StringVar(value="")

        right = ttk.Frame(container)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        self.schedule_selected_var = tk.StringVar(value=self.schedule_date.strftime("%A, %d %B %Y"))
        calendar_section = self._create_calendar_section(
            calendar_column, month_var=self.schedule_month_var, show_selected=True
        )
        calendar_section.pack(fill="x")

        info = ttk.Frame(right)
        info.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        info.columnconfigure(0, weight=1)
        self.schedule_selected_var = tk.StringVar(value="")
        ttk.Label(info, textvariable=self.schedule_selected_var, font=("Segoe UI Semibold", 12)).grid(
            row=0, column=0, sticky="w"
        )
        btn_bar = ttk.Frame(info)
        btn_bar.grid(row=0, column=1, sticky="e", padx=(12, 0))
        self.open_patient_btn = ttk.Button(btn_bar, text="Open Patient", style="Primary.TButton", command=self._open_patient_profile)
        self.open_patient_btn.pack(side="left")
        self.new_visit_btn = ttk.Button(btn_bar, text="New Visit Note", style="Primary.TButton", command=self._new_visit_note)
        self.new_visit_btn.pack(side="left", padx=(8, 0))
        self.open_patient_btn.state(["disabled"])
        self.new_visit_btn.state(["disabled"])
        self.schedule_summary_var = tk.StringVar(value="")
        ttk.Label(info, textvariable=self.schedule_summary_var, foreground="#667085").grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )
        ttk.Button(info, text="History Timeline", style="Ghost.TButton", command=self._open_history_timeline).grid(
            row=1, column=1, sticky="e", pady=(4, 0)
        )

        columns = ("time", "patient", "reason", "provider", "status", "room", "queue")
        self.appt_tree = ttk.Treeview(
            right,
            columns=columns,
            show="headings",
            height=14,
            selectmode="browse",
        )
        for name, label, width, anchor in (
            ("time", "Time", 80, tk.W),
            ("patient", "Patient", 160, tk.W),
            ("reason", "Reason", 220, tk.W),
            ("provider", "Provider", 130, tk.W),
            ("status", "Status", 120, tk.W),
            ("room", "Location", 110, tk.W),
            ("queue", "Queue #", 80, tk.CENTER),
        ):
            self.appt_tree.heading(name, text=label)
            self.appt_tree.column(name, width=width, anchor=anchor, stretch=False)
        self.appt_tree.grid(row=1, column=0, sticky="nsew")

        tree_scroll = ttk.Scrollbar(right, orient="vertical", command=self.appt_tree.yview)
        tree_scroll.grid(row=1, column=1, sticky="ns")
        self.appt_tree.configure(yscrollcommand=tree_scroll.set)
        self.appt_tree.bind("<Double-1>", self._open_history_from_schedule)
        self.appt_tree.bind("<Return>", self._open_visit_from_schedule)
        self.appt_tree.bind("<<TreeviewSelect>>", lambda _e: self._schedule_selection_changed())

        timeline_box = card(right, "Visit Timeline")
        timeline_box.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        right.rowconfigure(2, weight=1)
        timeline_box.columnconfigure(0, weight=1)
        timeline_box.rowconfigure(0, weight=1)

        timeline_columns = ("date", "appointment", "author", "diagnosis", "treatment", "history")
        self.timeline_tree = ttk.Treeview(
            timeline_box,
            columns=timeline_columns,
            show="headings",
            height=6,
            selectmode="browse",
        )
        for name, label, width, anchor in (
            ("date", "Generated", 150, tk.W),
            ("appointment", "Appointment", 140, tk.W),
            ("author", "Created By", 130, tk.W),
            ("diagnosis", "Diagnosis", 160, tk.W),
            ("treatment", "Treatment", 160, tk.W),
            ("history", "History", 260, tk.W),
        ):
            self.timeline_tree.heading(name, text=label)
            self.timeline_tree.column(name, width=width, anchor=anchor, stretch=False)
        self.timeline_tree.grid(row=0, column=0, sticky="nsew")
        timeline_scroll = ttk.Scrollbar(timeline_box, orient="vertical", command=self.timeline_tree.yview)
        timeline_scroll.grid(row=0, column=1, sticky="ns")
        self.timeline_tree.configure(yscrollcommand=timeline_scroll.set)

        self.schedule_message_var = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.schedule_message_var, foreground="#667085").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

        self._render_calendar()
        self._update_schedule_header()

    def _build_settlement_tab(self) -> None:
        container = ttk.Frame(self.settlement_tab)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=0)
        container.columnconfigure(1, weight=1)
        container.columnconfigure(2, weight=2)
        container.rowconfigure(0, weight=1)

        calendar_column = ttk.Frame(container)
        calendar_column.grid(row=0, column=0, sticky="nw", padx=(0, 16))
        calendar_section = self._create_calendar_section(calendar_column, show_selected=False)
        calendar_section.pack(fill="x")

        left = ttk.Frame(container)
        left.grid(row=0, column=1, sticky="nsew")
        left.columnconfigure(0, weight=1)
        left.columnconfigure(1, weight=0)
        left.rowconfigure(1, weight=1)

        filter_bar = ttk.Frame(left)
        filter_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(filter_bar, text="Today", style="Ghost.TButton", command=self._settlement_today).pack(
            side="left"
        )
        ttk.Button(filter_bar, text="Refresh", style="Ghost.TButton", command=self._load_settlement_list).pack(
            side="left", padx=(8, 0)
        )

        columns = ("time", "patient", "reason", "provider", "queue")
        self.settlement_tree = ttk.Treeview(
            left,
            columns=columns,
            show="headings",
            height=22,
            selectmode="browse",
        )
        for name, label, width, anchor in (
            ("time", "Time", 90, tk.W),
            ("patient", "Patient", 180, tk.W),
            ("reason", "Reason", 160, tk.W),
            ("provider", "Provider", 130, tk.W),
            ("queue", "Queue #", 80, tk.CENTER),
        ):
            self.settlement_tree.heading(name, text=label)
            self.settlement_tree.column(name, width=width, anchor=anchor, stretch=False)
        self.settlement_tree.grid(row=1, column=0, sticky="nsew")
        tree_scroll = ttk.Scrollbar(left, orient="vertical", command=self.settlement_tree.yview)
        tree_scroll.grid(row=1, column=1, sticky="ns")
        self.settlement_tree.configure(yscrollcommand=tree_scroll.set)
        self.settlement_tree.bind("<<TreeviewSelect>>", self._on_settlement_select)
        self.settlement_tree.bind("<Double-1>", self._on_settlement_select)

        ttk.Label(left, textvariable=self.settlement_message_var, foreground="#667085").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

        # Right column: patient + settlement details
        right = ttk.Frame(container)
        right.grid(row=0, column=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        info_card = card(right, "Patient")
        info_card.grid(row=0, column=0, sticky="ew")
        info_card.columnconfigure(0, weight=1)
        self.settlement_patient_var = tk.StringVar(value="Select a patient from the list.")
        self.settlement_contact_var = tk.StringVar(value="")
        ttk.Label(info_card, textvariable=self.settlement_patient_var, font=("Segoe UI Semibold", 11)).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(info_card, textvariable=self.settlement_contact_var, foreground="#475467").grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )

        note_card = card(right, "Visit Note")
        note_card.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        note_card.columnconfigure(0, weight=1)
        self.settlement_note_header_var = tk.StringVar(value="No visit note loaded.")
        ttk.Label(note_card, textvariable=self.settlement_note_header_var, foreground="#475467").grid(
            row=0, column=0, sticky="w"
        )
        self.settlement_note_text = tk.Text(note_card, wrap="word", height=6, state="disabled")
        self.settlement_note_text.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        note_scroll = ttk.Scrollbar(note_card, orient="vertical", command=self.settlement_note_text.yview)
        note_scroll.grid(row=1, column=1, sticky="ns", pady=(6, 0))
        self.settlement_note_text.configure(yscrollcommand=note_scroll.set)
        ttk.Button(
            note_card,
            text="Open Visit Note",
            style="Ghost.TButton",
            command=self._open_selected_visit_note,
        ).grid(row=2, column=0, sticky="e", pady=(6, 0))

        treatments_card = card(right, "Treatment Items")
        treatments_card.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        treatments_card.columnconfigure(0, weight=1)
        treatments_card.rowconfigure(0, weight=1)

        self.settlement_items_tree = ttk.Treeview(
            treatments_card,
            columns=("desc", "tooth", "qty", "unit", "total", "source"),
            show="headings",
            height=8,
            selectmode="browse",
        )
        for name, label, width, anchor in (
            ("desc", "Description", 260, tk.W),
            ("tooth", "Tooth", 80, tk.CENTER),
            ("qty", "Qty", 60, tk.CENTER),
            ("unit", "Unit Price", 110, tk.E),
            ("total", "Amount", 110, tk.E),
            ("source", "Source", 90, tk.W),
        ):
            self.settlement_items_tree.heading(name, text=label)
            self.settlement_items_tree.column(name, width=width, anchor=anchor, stretch=False)
        self.settlement_items_tree.grid(row=0, column=0, sticky="nsew")
        items_scroll = ttk.Scrollbar(treatments_card, orient="vertical", command=self.settlement_items_tree.yview)
        items_scroll.grid(row=0, column=1, sticky="ns")
        self.settlement_items_tree.configure(yscrollcommand=items_scroll.set)
        self.settlement_items_tree.bind("<<TreeviewSelect>>", lambda _e: self._update_settlement_item_buttons())

        btn_bar = ttk.Frame(treatments_card)
        btn_bar.grid(row=1, column=0, columnspan=2, sticky="e", pady=(8, 0))
        self.settlement_add_btn = ttk.Button(btn_bar, text="Add", command=self._settlement_add_item)
        self.settlement_edit_btn = ttk.Button(btn_bar, text="Edit", command=self._settlement_edit_item, state=tk.DISABLED)
        self.settlement_delete_btn = ttk.Button(
            btn_bar, text="Remove", command=self._settlement_remove_item, state=tk.DISABLED
        )
        self.settlement_add_btn.pack(side="left")
        self.settlement_edit_btn.pack(side="left", padx=(8, 0))
        self.settlement_delete_btn.pack(side="left", padx=(8, 0))

        totals = ttk.Frame(right)
        totals.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        totals.columnconfigure(1, weight=1)
        ttk.Label(totals, text="Subtotal").grid(row=0, column=0, sticky="w")
        ttk.Label(totals, textvariable=self.settlement_subtotal_var, font=("Segoe UI", 10, "bold")).grid(
            row=0, column=1, sticky="e"
        )
        ttk.Label(totals, text="Discount").grid(row=1, column=0, sticky="w", pady=(6, 0))
        discount_entry = ttk.Entry(totals, textvariable=self.settlement_discount_var, width=12)
        discount_entry.grid(row=1, column=1, sticky="e", pady=(6, 0))
        ttk.Label(totals, text="Rounding").grid(row=2, column=0, sticky="w", pady=(6, 0))
        rounding_entry = ttk.Entry(totals, textvariable=self.settlement_rounding_var, width=12)
        rounding_entry.grid(row=2, column=1, sticky="e", pady=(6, 0))
        ttk.Label(totals, text="Total Due").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Label(totals, textvariable=self.settlement_total_var, font=("Segoe UI", 11, "bold")).grid(
            row=3, column=1, sticky="e", pady=(8, 0)
        )

        payment_bar = ttk.Frame(right)
        payment_bar.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        payment_bar.columnconfigure(1, weight=1)
        ttk.Label(payment_bar, text="Payment Method").grid(row=0, column=0, sticky="w")
        self.settlement_payment_combo = ttk.Combobox(
            payment_bar,
            textvariable=self.settlement_payment_var,
            state="readonly",
            values=[],
            width=28,
        )
        self.settlement_payment_combo.grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Button(payment_bar, text="Refresh Methods", style="Ghost.TButton", command=self._refresh_payment_methods).grid(
            row=0, column=2, padx=(8, 0)
        )
        self.settlement_pay_btn = ttk.Button(
            payment_bar,
            text="Mark as Paid",
            style="Primary.TButton",
            command=self._complete_settlement,
            state=tk.DISABLED,
        )
        self.settlement_pay_btn.grid(row=0, column=3, padx=(16, 0))

        # Trace amount fields to recompute totals
        self.settlement_discount_var.trace_add("write", lambda *_: self._recalculate_settlement_totals())
        self.settlement_rounding_var.trace_add("write", lambda *_: self._recalculate_settlement_totals())

        self._ensure_payment_methods_loaded()
        self._refresh_payment_methods()
        self._load_settlement_list()

    def _settlement_today(self) -> None:
        self._set_shared_date(dt.date.today())

    def _load_settlement_list(self) -> None:
        tree = getattr(self, "settlement_tree", None)
        if tree is None:
            return
        tree.delete(*tree.get_children())
        self.settlement_index.clear()
        self.settlement_selected_appointment = None
        self._reset_settlement_items()
        self.settlement_patient_var.set("Select a patient from the list.")
        self.settlement_contact_var.set("")
        self.settlement_note_header_var.set("No visit note loaded.")
        self._update_settlement_note_display(None)

        db = self._ensure_login_database()
        if not db:
            self.settlement_message_var.set("Database connection required to view settlements.")
            return

        target = parse_date(self.settlement_date_var.get())
        if not target:
            target = dt.date.today()
            self.settlement_date_var.set(target.isoformat())

        try:
            settlements = db.appointments_for_status(target, (3,))
        except Exception as exc:
            self.settlement_message_var.set(f"Failed to load settlements: {exc}")
            return

        if not settlements:
            self.settlement_message_var.set("No patients queued for settlement on this date.")
            return

        for appt in settlements:
            values = (
                appt.scheduled.strftime("%H:%M"),
                appt.patient_name or appt.patient_id or "-",
                appt.reason or "-",
                appt.resource or "-",
                appt.queue_number or "-",
            )
            iid = tree.insert("", "end", values=values)
            self.settlement_index[iid] = appt

        first = tree.get_children()
        if first:
            tree.selection_set(first[0])
            tree.focus(first[0])
            self._on_settlement_select()
        self.settlement_message_var.set(f"{len(settlements)} patient(s) awaiting settlement.")

    def _on_settlement_select(self, _event=None) -> None:
        sel = self.settlement_tree.selection()
        if not sel:
            self.settlement_selected_appointment = None
            self._reset_settlement_items()
            self._update_settlement_note_display(None)
            self._update_settlement_item_buttons()
            self.settlement_pay_btn.configure(state=tk.DISABLED)
            return

        iid = sel[0]
        appointment = self.settlement_index.get(iid)
        if not appointment:
            return

        self.settlement_selected_appointment = appointment
        self.settlement_discount_var.set(0.0)
        self.settlement_rounding_var.set(0.0)
        self.settlement_payment_var.set("")

        patient_display = appointment.patient_name or appointment.patient_id or "Unknown patient"
        patient_id = appointment.patient_id or "-"
        self.settlement_patient_var.set(f"{patient_display} ({patient_id})")

        db = self._ensure_login_database()
        contact_summary = ""
        note: Optional[MedicReportDetail] = None
        if db and appointment.patient_id:
            try:
                profile = db.get_patient_profile(appointment.patient_id)
            except Exception:
                profile = None
            if profile:
                contacts = [
                    profile.phone_mobile or "",
                    profile.phone_fixed or "",
                    profile.email or "",
                ]
                contact_summary = " | ".join(filter(None, contacts))
            try:
                note = db.medic_report_for_appointment(appointment.patient_id, appointment.scheduled)
            except Exception as exc:
                self.settlement_message_var.set(f"Failed to load visit note: {exc}")
                note = None
        self.settlement_contact_var.set(contact_summary)
        self.settlement_current_note = note
        if db:
            self._populate_settlement_items(appointment, note, db)
        else:
            self._populate_settlement_items(appointment, note, self._ensure_login_database())
        self._update_settlement_note_display(note)
        self._update_settlement_item_buttons()
        self._ensure_payment_methods_loaded()
        self._refresh_payment_methods()
        if self.settlement_items:
            self.settlement_pay_btn.configure(state=tk.NORMAL)
        else:
            self.settlement_pay_btn.configure(state=tk.DISABLED)

    def _populate_settlement_items(
        self,
        appointment: AppointmentDetail,
        note: Optional[MedicReportDetail],
        db: ClinicDatabase,
    ) -> None:
        self.settlement_items = []
        self.settlement_receipt_text = ""
        self.settlement_receipt_header = ""
        if note and note.chart_items:
            self.settlement_current_receipt = None
            for chart in note.chart_items:
                description = chart.notation_title or chart.stock_name or f"Notation #{chart.notation_id}"
                self.settlement_items.append(
                    _SettlementItem(
                        stock_id=chart.stock_id or chart.notation_title or description,
                        description=description,
                        qty=1,
                        unit_price=float(chart.unit_price or 0.0),
                        notation_id=chart.notation_id,
                        tooth_label=f"{chart.tooth_plan}-{chart.tooth_id}" if chart.tooth_id else chart.tooth_plan,
                        remarks=chart.remarks or "",
                        source="note",
                    )
                )
            self._prefill_visit_note_items = [
                VisitNoteItem(
                    stock_id=item.stock_id,
                    name=item.description,
                    category=item.source or '',
                    unit_price=item.unit_price,
                    qty=item.qty,
                )
                    for item in self.settlement_items
                ]
            self._prefill_visit_note_patient_id = appointment.patient_id or ""
            if self.payment_methods:
                self.settlement_payment_var.set(self.payment_methods[0][1])
        else:
            items, header, text, primary_receipt = self._receipts_to_settlement_items(
                appointment, db, note.report_id if note else 0
            )
            if items:
                self.settlement_items = items
                self.settlement_receipt_header = header
                self.settlement_receipt_text = text
                self.settlement_current_receipt = primary_receipt
                self._prefill_visit_note_items = [
                    VisitNoteItem(
                        stock_id=itm.stock_id,
                        name=itm.description,
                        category=self._stock_category_for_item(itm.stock_id, db),
                        unit_price=itm.unit_price,
                        qty=itm.qty,
                    )
                    for itm in items
                ]
                self._prefill_visit_note_patient_id = appointment.patient_id or ""
                if primary_receipt:
                    self.settlement_payment_var.set(self._payment_label_for_code(primary_receipt.payment_code))
                else:
                    if self.payment_methods:
                        self.settlement_payment_var.set(self.payment_methods[0][1])
            else:
                self.settlement_current_receipt = None
                self._prefill_visit_note_items = None
                self._prefill_visit_note_patient_id = None
        self._render_settlement_items()

    def _reset_settlement_items(self) -> None:
        self.settlement_items = []
        self.settlement_item_map.clear()
        self.settlement_receipt_text = ""
        self.settlement_receipt_header = ""
        self._prefill_visit_note_items = None
        self._prefill_visit_note_patient_id = None
        self.settlement_current_receipt = None
        if hasattr(self, "settlement_items_tree"):
            self.settlement_items_tree.delete(*self.settlement_items_tree.get_children())
        self._recalculate_settlement_totals()

    def _render_settlement_items(self) -> None:
        tree = getattr(self, "settlement_items_tree", None)
        if not tree:
            return
        tree.delete(*tree.get_children())
        self.settlement_item_map.clear()
        for item in self.settlement_items:
            total = item.qty * item.unit_price
            iid = tree.insert(
                "",
                "end",
                values=(
                    item.description,
                    item.tooth_label or "-",
                    item.qty,
                    fmt_money(item.unit_price),
                    fmt_money(total),
                    item.source or "-",
                ),
            )
            self.settlement_item_map[iid] = item
        self._recalculate_settlement_totals()

    def _receipts_to_settlement_items(
        self,
        appointment: AppointmentDetail,
        db: ClinicDatabase,
        report_id: int = 0,
    ) -> tuple[list[_SettlementItem], str, str, Optional[Receipt]]:
        if not db:
            return [], "", "", None
        receipts: list[Receipt] = []
        if report_id:
            try:
                receipts = db.receipts_for_medic_report(report_id)
            except Exception as exc:
                self.settlement_message_var.set(f"Failed to load receipts for settlement: {exc}")
                receipts = []
        if not receipts and appointment.patient_id:
            try:
                receipt_pairs = db.receipt_items_for_patient_date(
                    appointment.patient_id,
                    appointment.scheduled.date(),
                )
            except Exception as exc:
                self.settlement_message_var.set(f"Failed to load receipts for settlement: {exc}")
                return [], "", "", None
            if not receipt_pairs:
                return [], "", "", None
            latest_receipt, latest_items = receipt_pairs[-1]
            receipts = [latest_receipt]
            receipt_items_map = {latest_receipt.rcpt_id: latest_items}
        else:
            receipt_items_map: dict[str, list[ReceiptItem]] = {}
            for receipt in receipts:
                try:
                    receipt_items_map[receipt.rcpt_id] = db.get_receipt_items(receipt.rcpt_id)
                except Exception:
                    receipt_items_map[receipt.rcpt_id] = []
        if not receipts:
            return [], "", "", None
        primary_receipt = max(receipts, key=lambda r: r.issued)
        items: list[_SettlementItem] = []
        header_parts: list[str] = []
        lines: list[str] = []
        for receipt in receipts:
            issued_label = receipt.issued.strftime("%H:%M") if isinstance(receipt.issued, dt.datetime) else ""
            header_parts.append(f"{receipt.rcpt_id} {issued_label}".strip())
            lines.append(f"Receipt {receipt.rcpt_id} ({issued_label})")
            for entry in receipt_items_map.get(receipt.rcpt_id, []):
                notation_id = 0
                description = entry.name or entry.item_id or "Item"
                try:
                    notation = db.notation_for_stock(entry.item_id or "")
                except Exception:
                    notation = None
                if notation:
                    notation_id = notation.notation_id
                    description = notation.title or notation.stock_name or description
                items.append(
                    _SettlementItem(
                        stock_id=entry.item_id or "",
                        description=description,
                        qty=max(1, int(entry.qty or 0)),
                        unit_price=float(entry.unit_price or 0.0),
                        notation_id=notation_id,
                        tooth_label="",
                        remarks=entry.remark or "",
                        source=receipt.rcpt_id,
                    )
                )
                lines.append(
                    f"  - {entry.name or entry.item_id or 'Item'} x{entry.qty or 1} @ {fmt_money(entry.unit_price or 0.0)}"
                )
        header = f"Receipt(s) linked to visit: " + ", ".join(header_parts)
        text = "\n".join(lines)
        return items, header, text, primary_receipt
        try:
            receipt_pairs = db.receipt_items_for_patient_date(
                appointment.patient_id,
                appointment.scheduled.date(),
            )
        except Exception as exc:
            self.settlement_message_var.set(f"Failed to load receipts for settlement: {exc}")
            return [], "", ""
        if not receipt_pairs:
            return [], "", ""

        items: list[_SettlementItem] = []
        header_parts: list[str] = []
        lines: list[str] = []
        for receipt, receipt_items in receipt_pairs:
            issued_label = receipt.issued.strftime("%H:%M") if isinstance(receipt.issued, dt.datetime) else ""
            header_parts.append(f"{receipt.rcpt_id} {issued_label}".strip())
            lines.append(f"Receipt {receipt.rcpt_id} ({issued_label})")
            for entry in receipt_items:
                notation_id = 0
                description = entry.name or entry.item_id or "Item"
                try:
                    notation = db.notation_for_stock(entry.item_id or "")
                except Exception:
                    notation = None
                if notation:
                    notation_id = notation.notation_id
                    description = notation.title or notation.stock_name or description
                items.append(
                    _SettlementItem(
                        stock_id=entry.item_id or "",
                        description=description,
                        qty=max(1, int(entry.qty or 0)),
                        unit_price=float(entry.unit_price or 0.0),
                        notation_id=notation_id,
                        tooth_label="",
                        remarks=entry.remark or "",
                        source=receipt.rcpt_id,
                    )
                )
                lines.append(
                    f"  - {entry.name or entry.item_id or 'Item'} x{entry.qty or 1} @ {fmt_money(entry.unit_price or 0.0)}"
                )

        header = f"Receipts on {appointment.scheduled.strftime('%d %b %Y')}: " + ", ".join(header_parts)
        text = "\n".join(lines)
        return items, header, text

    def _chart_item_from_settlement(
        self,
        db: ClinicDatabase,
        item: _SettlementItem,
    ) -> Optional[DentalChartItem]:
        notation_id = item.notation_id
        notation_title = item.description
        stock_name = item.description
        stock_id = item.stock_id
        if notation_id <= 0 and item.stock_id:
            try:
                notation = db.notation_for_stock(item.stock_id)
            except Exception:
                notation = None
            if notation:
                notation_id = notation.notation_id
                notation_title = notation.title or notation.stock_name or notation_title
                stock_name = notation.stock_name or notation_title
                stock_id = notation.stock_id or item.stock_id
        if notation_id <= 0 and not stock_id:
            return None
        return DentalChartItem(
            notation_id=notation_id,
            tooth_id=0,
            tooth_plan="E",
            remarks=item.remarks or "",
            unit_price=float(item.unit_price or 0.0),
            notation_status=1,
            bill_status=0,
            notation_title=notation_title,
            stock_name=stock_name,
            stock_id=stock_id,
        )

    def _chart_items_from_settlement(
        self,
        db: ClinicDatabase,
        items: list[_SettlementItem],
    ) -> list[DentalChartItem]:
        charts: list[DentalChartItem] = []
        for itm in items:
            chart = self._chart_item_from_settlement(db, itm)
            if chart:
                charts.append(chart)
        return charts

    def _get_var_float(self, var: tk.Variable) -> float:
        try:
            value = var.get()
        except tk.TclError:
            return 0.0
        if value in ("", None):
            return 0.0
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    def _recalculate_settlement_totals(self) -> None:
        subtotal = sum(item.qty * item.unit_price for item in self.settlement_items)
        discount = self._get_var_float(self.settlement_discount_var)
        rounding = self._get_var_float(self.settlement_rounding_var)
        total = subtotal - discount + rounding
        self.settlement_subtotal_var.set(fmt_money(subtotal))
        self.settlement_total_var.set(fmt_money(total))

    def _payment_method_labels(self) -> List[str]:
        return [label for _code, label in self.payment_methods]

    def _payment_label_for_code(self, code: str) -> str:
        self._ensure_payment_methods_loaded()
        code = (code or "").strip()
        # Prefer the label from code->label map; fall back to the code itself
        return self.payment_code_map.get(code, code)


    def _ensure_payment_methods_loaded(self) -> None:
        if self.payment_methods:
            return
        db = self._ensure_login_database()
        if not db:
            return
        try:
            methods = db.payment_methods_list()  # expected: List[Tuple[code, label]]
        except Exception as exc:
            messagebox.showerror("Payment Methods", f"Failed to load payment methods:\n{exc}")
            self.payment_methods = []
            self.payment_method_map = {}
            self.payment_code_map = {}
            return

        # Build consistent structures:
        # - self.payment_methods: [(code, label), ...]
        # - self.payment_method_map: {label -> code, code -> code}  (for resolving selection to a CODE)
        # - self.payment_code_map:   {code -> label}               (for showing a LABEL from a CODE)
        converted: list[tuple[str, str]] = []
        method_map: dict[str, str] = {}
        code_map: dict[str, str] = {}

        for code, label in methods:
            code = (code or "").strip()
            label = (label or code or "").strip()
            if not label:
                continue
            converted.append((code or label, label))
            # label -> code (user selection text to short code)
            if label:
                method_map[label] = code or label
            # code -> code (if user somehow selects/types code directly)
            if code:
                method_map[code] = code
                code_map[code] = label

        self.payment_methods = converted
        self.payment_method_map = method_map   # label->code, code->code
        self.payment_code_map = code_map       # code->label


    def _refresh_payment_methods(self) -> None:
        combo = getattr(self, "settlement_payment_combo", None)
        if combo is None:
            return
        current = self.settlement_payment_var.get()
        self._ensure_payment_methods_loaded()
        labels = self._payment_method_labels()
        combo["values"] = labels
        if current and current in labels:
            combo.set(current)
        elif labels:
            combo.set(labels[0])
            self.settlement_payment_var.set(labels[0])

    def _ensure_stock_categories(self) -> list[str]:
        if self.stock_categories is not None:
            return self.stock_categories
        db = self._ensure_login_database()
        if not db:
            self.stock_categories = []
            return []
        try:
            categories = db.stock_categories()
        except Exception as exc:
            messagebox.showerror("Stock Catalogue", f"Failed to load stock categories\n{exc}")
            categories = []
        self.stock_categories = categories
        return categories

    def _stock_items_for_category(self, category: str, db: ClinicDatabase | None = None) -> list[tuple[str, str, float]]:
        cat = (category or '').strip()
        if not cat:
            return []
        if cat in self.stock_items_cache:
            return self.stock_items_cache[cat]
        if db is None:
            db = self._ensure_login_database()
        if not db:
            return []
        try:
            items = db.stock_items_by_category(cat)
        except Exception as exc:
            messagebox.showerror("Stock Catalogue", f"Failed to load items for {cat}:\n{exc}")
            items = []
        self.stock_items_cache[cat] = items
        return items

    def _stock_category_for_item(self, stock_id: str, db: ClinicDatabase | None = None) -> str:
        sid = (stock_id or '').strip()
        if not sid:
            return ''
        categories = self._ensure_stock_categories()
        for category in categories:
            items = self._stock_items_for_category(category, db)
            for item_id, _name, _price in items:
                if item_id == sid:
                    return category
        return ''

    def _update_settlement_item_buttons(self) -> None:
        has_selection = bool(self.settlement_items_tree.selection())
        state = tk.NORMAL if has_selection else tk.DISABLED
        self.settlement_edit_btn.configure(state=state)
        self.settlement_delete_btn.configure(state=state)
        if self.settlement_items:
            self.settlement_pay_btn.configure(state=tk.NORMAL)
        else:
            self.settlement_pay_btn.configure(state=tk.DISABLED)

    def _settlement_selected_item(self) -> Tuple[str, _SettlementItem] | Tuple[None, None]:
        sel = self.settlement_items_tree.selection()
        if not sel:
            return None, None
        iid = sel[0]
        item = self.settlement_item_map.get(iid)
        if not item:
            return None, None
        return iid, item

    def _settlement_add_item(self) -> None:
        item = self._open_settlement_item_dialog()
        if not item:
            return
        self.settlement_items.append(item)
        self._render_settlement_items()
        self._update_settlement_item_buttons()

    def _settlement_edit_item(self) -> None:
        _iid, existing = self._settlement_selected_item()
        if not existing:
            return
        updated = self._open_settlement_item_dialog(existing=existing)
        if not updated:
            return
        try:
            index = self.settlement_items.index(existing)
        except ValueError:
            return
        self.settlement_items[index] = updated
        self._render_settlement_items()
        self._update_settlement_item_buttons()

    def _settlement_remove_item(self) -> None:
        iid, existing = self._settlement_selected_item()
        if not existing:
            return
        try:
            self.settlement_items.remove(existing)
        except ValueError:
            return
        if iid:
            self.settlement_items_tree.delete(iid)
            self.settlement_item_map.pop(iid, None)
        self._recalculate_settlement_totals()
        self._update_settlement_item_buttons()

    def _open_settlement_item_dialog(self, existing: Optional[_SettlementItem] = None) -> Optional[_SettlementItem]:
        db = self._ensure_login_database()
        if not db:
            messagebox.showerror("Settlement", "Database connection is not available.")
            return None
        categories = self._ensure_stock_categories()
        if not categories:
            messagebox.showinfo("Settlement", "Stock catalogue is not available.")
            return None

        dialog = tk.Toplevel(self)
        dialog.title("Edit Item" if existing else "Add Item")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        ttk.Label(dialog, text="Category").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        default_category = self._stock_category_for_item(existing.stock_id, db) if existing else categories[0]
        category_var = tk.StringVar(value=default_category)
        category_combo = ttk.Combobox(dialog, textvariable=category_var, state="readonly", values=categories, width=32)
        category_combo.grid(row=0, column=1, sticky="ew", padx=12, pady=(12, 4))

        ttk.Label(dialog, text="Item").grid(row=1, column=0, sticky="w", padx=12, pady=4)
        item_var = tk.StringVar()
        item_combo = ttk.Combobox(dialog, textvariable=item_var, state="readonly", width=48)
        item_combo.grid(row=1, column=1, sticky="ew", padx=12, pady=4)

        ttk.Label(dialog, text="Quantity").grid(row=2, column=0, sticky="w", padx=12, pady=4)
        qty_var = tk.IntVar(value=existing.qty if existing else 1)
        ttk.Entry(dialog, textvariable=qty_var, width=8).grid(row=2, column=1, sticky="w", padx=12, pady=4)

        ttk.Label(dialog, text="Unit Price").grid(row=3, column=0, sticky="w", padx=12, pady=4)
        unit_price_value = tk.DoubleVar(value=existing.unit_price if existing else 0.0)
        unit_price_entry = ttk.Entry(dialog, textvariable=unit_price_value, width=14, justify="right")
        unit_price_entry.grid(row=3, column=1, sticky="w", padx=12, pady=4)

        ttk.Label(dialog, text="Remarks").grid(row=4, column=0, sticky="w", padx=12, pady=4)
        remark_var = tk.StringVar(value=existing.remarks if existing else "")
        ttk.Entry(dialog, textvariable=remark_var, width=48).grid(row=4, column=1, sticky="ew", padx=12, pady=4)

        buttons = ttk.Frame(dialog)
        buttons.grid(row=5, column=0, columnspan=2, sticky="e", pady=(12, 12), padx=12)

        current_options: list[tuple[str, str, float]] = []

        def load_items_for_category(*_args) -> None:
            category = category_var.get().strip()
            options = self._stock_items_for_category(category, db) if category else []
            current_options.clear()
            current_options.extend(options)
            display = [name for _sid, name, _price in options]
            item_combo['values'] = display
            target_stock = existing.stock_id if existing else ''
            if target_stock:
                for idx, (sid, name, price) in enumerate(options):
                    if sid == target_stock:
                        item_combo.current(idx)
                        item_var.set(name)
                        unit_price_value.set(price)
                        return
            if display:
                item_combo.current(0)
                item_var.set(display[0])
                unit_price_value.set(options[0][2])
            else:
                item_var.set('')
                unit_price_value.set(0.0)

        def on_item_selected(*_args) -> None:
            name = item_var.get().strip()
            for stock_id, stock_name, price in current_options:
                if stock_name == name:
                    unit_price_value.set(price)
                    return
            unit_price_value.set(0.0)

        def on_cancel() -> None:
            dialog.destroy()

        result: dict[str, _SettlementItem] = {}

        def _current_unit_price() -> Optional[float]:
            try:
                return float(unit_price_value.get())
            except (tk.TclError, ValueError):
                return None

        def on_ok() -> None:
            name = item_var.get().strip()
            category = category_var.get().strip()
            if not category or not name:
                messagebox.showwarning("Items", "Select a category and item first.", parent=dialog)
                return
            unit_price = _current_unit_price()
            if unit_price is None:
                messagebox.showwarning("Items", "Enter a valid unit price.", parent=dialog)
                unit_price_entry.focus_set()
                return
            stock_id = ''
            for sid, stock_name, price in current_options:
                if stock_name == name:
                    stock_id = sid
                    break
            if not stock_id:
                messagebox.showwarning("Items", "Unable to resolve the selected item.", parent=dialog)
                return
            try:
                qty = max(1, int(qty_var.get()))
            except Exception:
                messagebox.showwarning("Items", "Quantity must be a positive integer.", parent=dialog)
                return
            remarks = remark_var.get().strip()
            notation_id = 0
            try:
                notation = db.notation_for_stock(stock_id)
            except Exception:
                notation = None
            if notation:
                notation_id = notation.notation_id
            result['item'] = _SettlementItem(
                stock_id=stock_id,
                description=name,
                qty=qty,
                unit_price=unit_price,
                notation_id=notation_id,
                tooth_label='',
                remarks=remarks,
                source=category,
            )
            dialog.destroy()

        category_combo.bind("<<ComboboxSelected>>", load_items_for_category)
        item_combo.bind("<<ComboboxSelected>>", on_item_selected)

        ttk.Button(buttons, text="Cancel", command=on_cancel, style="Ghost.TButton").pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="Save", command=on_ok, style="Primary.TButton").pack(side="right")

        load_items_for_category()
        if not item_var.get() and current_options:
            item_combo.current(0)
            item_var.set(current_options[0][1])
            unit_price_value.set(current_options[0][2])

        dialog.wait_window(dialog)
        return result.get('item')


    def _update_settlement_note_display(self, note: Optional[MedicReportDetail]) -> None:
        text_widget = getattr(self, "settlement_note_text", None)
        if not text_widget:
            return
        text_widget.configure(state="normal")
        text_widget.delete("1.0", "end")
        if not note:
            if self.settlement_receipt_text:
                header = self.settlement_receipt_header or "Receipt items"
                self.settlement_note_header_var.set(header)
                text_widget.insert("1.0", self.settlement_receipt_text)
            else:
                self.settlement_note_header_var.set("No visit note loaded.")
                text_widget.insert("1.0", "No visit note recorded for this appointment.")
            text_widget.configure(state="disabled")
            return
        header_parts = []
        if note.generated_on:
            header_parts.append(f"Generated: {self._format_datetime(note.generated_on)}")
        if note.created_by:
            header_parts.append(f"By: {note.created_by}")
        if note.appointment_date:
            header_parts.append(f"Appointment: {self._format_datetime(note.appointment_date)}")
        header = " | ".join(header_parts)
        self.settlement_note_header_var.set(header)
        sections = []
        if note.diagnosis:
            sections.append(("Diagnosis", _clean_note_text(note.diagnosis)))
        if note.treatment:
            sections.append(("Treatment", _clean_note_text(note.treatment)))
        if note.history:
            sections.append(("History", _clean_note_text(note.history)))
        if note.examination:
            sections.append(("Examination", _clean_note_text(note.examination)))
        if note.finding:
            sections.append(("Findings", _clean_note_text(note.finding)))
        if note.advice:
            sections.append(("Advice", _clean_note_text(note.advice)))
        if note.next_action:
            sections.append(("Next Action", _clean_note_text(note.next_action)))
        if not sections:
            text_widget.insert("1.0", "No narrative details recorded in this visit note.")
        else:
            lines = []
            for title, body in sections:
                if body:
                    lines.append(f"{title}:\n{body}\n")
            text_widget.insert("1.0", "\n".join(lines).strip())
        text_widget.configure(state="disabled")

    def _open_selected_visit_note(self) -> None:
        if not self.settlement_selected_appointment:
            messagebox.showinfo("Visit Note", "Select a patient first.")
            return
        self._open_visit_note_editor(self.settlement_selected_appointment)

    def _convert_settlement_items_to_receipt(self, db: ClinicDatabase) -> Optional[List[ReceiptDraftItem]]:
        receipt_items: List[ReceiptDraftItem] = []
        for idx, item in enumerate(self.settlement_items, start=1):
            subtotal = float(item.qty * item.unit_price)
            stock_id = (item.stock_id or "").strip()
            if not stock_id and item.notation_id:
                try:
                    note = db.notation_by_id(item.notation_id)
                except Exception:
                    note = None
                if note and note.stock_id:
                    stock_id = note.stock_id.strip()
            if stock_id:
                try:
                    details = db.stock_item_details(stock_id)
                except Exception as exc:
                    messagebox.showerror(
                        "Settlement",
                        f"Failed to look up stock item {stock_id}:\n{exc}",
                    )
                    return None
                if not details:
                    stock_id = ""
            if not stock_id:
                messagebox.showwarning(
                    "Settlement",
                    (
                        f"Treatment '{item.description}' does not reference a valid stock item.\n"
                        "Edit the entry and choose a procedure code that exists in stock items."
                    ),
                )
                return None
            receipt_items.append(
                ReceiptDraftItem(
                    stock_id=stock_id,
                    description=item.description,
                    qty=int(item.qty),
                    unit_price=float(item.unit_price),
                    subtotal=subtotal,
                    remark=item.remarks,
                )
            )
        return receipt_items

    def _complete_settlement(self) -> None:
        if not self.settlement_selected_appointment or not self.settlement_selected_appointment.patient_id:
            messagebox.showwarning("Settlement", "Select an appointment before marking as paid.")
            return
        if not self.settlement_items:
            messagebox.showwarning("Settlement", "Add at least one item before marking as paid.")
            return

        # Read selection from the combobox
        payment_label = (self.settlement_payment_var.get() or "").strip()
        if not payment_label:
            messagebox.showwarning("Settlement", "Choose a payment method.", parent=self)
            return

        # Ensure payment maps are loaded
        self._ensure_payment_methods_loaded()

        # Convert the UI value to a short CODE (what the DB expects).
        # We support these cases:
        #  - Combobox shows "CODE - Label"  -> take the left side as code
        #  - Combobox shows just Label      -> map Label -> Code via payment_method_map
        #  - User typed a Code directly     -> keep as-is
        payment_code = None
        if " - " in payment_label:
            payment_code = payment_label.split(" - ", 1)[0].strip()

        if not payment_code:
            # payment_method_map should map Label -> Code (and also Code -> Code if you set it that way)
            m = getattr(self, "payment_method_map", {}) or {}
            payment_code = (m.get(payment_label) or "").strip()

        if not payment_code:
            # assume user typed the code directly
            payment_code = payment_label

        if not payment_code:
            messagebox.showwarning("Settlement", "Selected payment method is not recognised.", parent=self)
            return

        db = self._ensure_login_database()
        if not db:
            messagebox.showerror("Settlement", "Database connection is unavailable.")
            return

        receipt_drafts = self._convert_settlement_items_to_receipt(db)
        if not receipt_drafts:
            return

        # from here on, use payment_code when calling your DAL methods
        # (your existing code that saves the receipt can stay the same if it already takes payment_code)
        # e.g., db.create_or_replace_receipt(..., payment_code=payment_code)
        # (leave your existing save logic as-is, just ensure it uses payment_code variable above)


        subtotal = sum(item.subtotal for item in receipt_drafts)
        discount = self._get_var_float(self.settlement_discount_var)
        rounding = self._get_var_float(self.settlement_rounding_var)
        total = subtotal - discount + rounding

        note = self.settlement_current_note
        remark_parts = []
        if note and note.diagnosis:
            remark_parts.append(_clean_note_text(note.diagnosis))
        if note and note.treatment:
            remark_parts.append(_clean_note_text(note.treatment))
        if not remark_parts:
            remark_parts = [item.description for item in self.settlement_items if item.description]
        remark = "; ".join(part for part in remark_parts if part)
        mr_id = note.report_id if note else 0
        issued_source = getattr(self.settlement_selected_appointment, "scheduled", None)
        issued_dt = issued_source if isinstance(issued_source, dt.datetime) else dt.datetime.now()
        department = getattr(self.settlement_selected_appointment, "department_type", None) or "Clinic"
        method_label = self._payment_label_for_code(payment_code)

        receipt_error: Optional[str] = None
        receipt_id: Optional[str] = None
        existing_receipt = self.settlement_current_receipt

        try:
            if existing_receipt:
                db.replace_receipt(
                    existing_receipt.rcpt_id,
                    issued=issued_dt,
                    patient_id=self.settlement_selected_appointment.patient_id,
                    items=receipt_drafts,
                    subtotal=subtotal,
                    discount=discount,
                    rounding=rounding,
                    consult_fees=0.0,
                    remark=remark,
                    payment_code=payment_code,
                    username=self.session_user or existing_receipt.done_by or "",
                    department=existing_receipt.department_type or department,
                    mr_id=mr_id,
                )
                receipt_id = existing_receipt.rcpt_id
                self.settlement_current_receipt = replace(
                    existing_receipt,
                    issued=issued_dt,
                    subtotal=subtotal,
                    total=total,
                    payment_code=payment_code,
                    remark=remark,
                    discount=discount,
                    rounding=rounding,
                    consult_fees=0.0,
                    department_type=existing_receipt.department_type or department,
                    done_by=self.session_user or existing_receipt.done_by or "cashier",
                    settled_by=self.session_user or existing_receipt.settled_by or "cashier",
                    mr_id=mr_id,
                )
            else:
                receipt_id = db.create_receipt(
                    patient_id=self.settlement_selected_appointment.patient_id,
                    issued=issued_dt,
                    username=self.session_user or "cashier",
                    payment_code=payment_code,
                    items=receipt_drafts,
                    subtotal=subtotal,
                    discount=discount,
                    rounding=rounding,
                    consult_fees=0.0,
                    remark=remark,
                    mr_id=mr_id,
                    department=department,
                )
                self.settlement_current_receipt = Receipt(
                    rcpt_id=receipt_id,
                    issued=issued_dt,
                    patient_id=self.settlement_selected_appointment.patient_id,
                    total=total,
                    subtotal=subtotal,
                    gst=0.0,
                    payment_code=payment_code,
                    remark=remark,
                    discount=discount,
                    rounding=rounding,
                    consult_fees=0.0,
                    done_by=self.session_user or "cashier",
                    department_type=department,
                    settled_by=self.session_user or "cashier",
                    tax_total=0.0,
                    mr_id=mr_id,
                )
        except Exception as exc:
            receipt_error = str(exc)

        status_error: Optional[str] = None
        try:
            db.record_appointment_status(
                self.settlement_selected_appointment.patient_id,
                self.settlement_selected_appointment.scheduled,
                99,
                self.session_user or "",
            )
        except Exception as exc:
            status_error = str(exc)

        if receipt_id and not receipt_error:
            self.settlement_message_var.set(
                f"Receipt {receipt_id} ({method_label}) saved. Appointment marked as CLOSED."
            )
            messagebox.showinfo("Settlement", f"Receipt {receipt_id} saved successfully.")
            self.settlement_payment_var.set(method_label)
        if receipt_error:
            messagebox.showwarning("Settlement", f"Failed to save receipt\n{receipt_error}")
        if status_error:
            messagebox.showwarning("Settlement", f"Receipt saved, but failed to update appointment status\n{status_error}")

        self._populate_settlement_items(
            self.settlement_selected_appointment,
            self.settlement_current_note,
            db,
        )


    def _render_calendar(self) -> None:
        month_label = self._calendar_month.strftime("%B %Y")
        for host in self._calendar_hosts:
            host.month_var.set(month_label)
            self._render_calendar_into(host.grid)
        self._update_schedule_header()

    def _update_schedule_header(self) -> None:
        try:
            self.schedule_month_var.set(self._calendar_month.strftime("%B %Y"))
            self.schedule_selected_var.set(self.schedule_date.strftime("%A, %d %B %Y"))
        except Exception:
            pass

    def _change_schedule_month(self, months: int) -> None:
        base = self._calendar_month
        new_month = base.month + months
        new_year = base.year + (new_month - 1) // 12
        new_month = ((new_month - 1) % 12) + 1
        max_day = pycal.monthrange(new_year, new_month)[1]
        new_day = min(self.schedule_date.day, max_day)
        candidate = dt.date(new_year, new_month, new_day)
        self._set_shared_date(candidate)

    def _on_calendar_day(self, day: dt.date) -> None:
        self._set_shared_date(day)


    def _load_schedule_for(self, target: dt.date) -> None:
        self.schedule_date = target
        self._update_schedule_header()
        self.schedule_summary_var.set("Loading...")
        self.schedule_message_var.set("")
        self.appt_tree.delete(*self.appt_tree.get_children())
        self.schedule_index = {}
        db = self._ensure_login_database()
        if not db:
            self.schedule_summary_var.set("No database connection.")
            self.schedule_message_var.set("Connect to MySQL to view appointments.")
            return
        try:
            appointments = db.appointments_for_date(target)
        except Exception as exc:
            self.schedule_summary_var.set("Failed to load appointments.")
            self.schedule_message_var.set(str(exc))
            return

        self.schedule_appointments = appointments
        if not appointments:
            self.schedule_summary_var.set("No appointments.")
            self.schedule_message_var.set("No appointments scheduled for this date.")
            self.timeline_tree.delete(*self.timeline_tree.get_children())
            return

        self.schedule_summary_var.set(f"{len(appointments)} appointment(s)")
        self.schedule_message_var.set("")

        for appt in appointments:
            time_label = appt.scheduled.strftime("%H:%M")
            values = (
                time_label,
                appt.patient_name or appt.patient_id or "-",
                appt.reason or "-",
                appt.resource or "-",
                appt.status or "-",
                appt.location or "-",
                appt.queue_number or "-",
            )
            tag = self._schedule_tag_for_color(appt.status_color)
            iid = self.appt_tree.insert("", "end", values=values, tags=(tag,))
            self.schedule_index[iid] = appt
        self.timeline_tree.delete(*self.timeline_tree.get_children())

    def _schedule_tag_for_color(self, colour: str) -> str:
        hex_colour = self._normalise_hex(colour)
        tag_name = f"status::{hex_colour}"
        if tag_name not in self._calendar_color_tags:
            fg = self._schedule_foreground(hex_colour)
            try:
                self.appt_tree.tag_configure(tag_name, background=hex_colour, foreground=fg)
            except tk.TclError:
                # Fallback when platform does not allow custom colours
                self.appt_tree.tag_configure(tag_name, background="#FFFFFF", foreground="#1F2933")
            self._calendar_color_tags[tag_name] = hex_colour
        return tag_name

    @staticmethod
    def _normalise_hex(colour: str) -> str:
        text = (colour or "").strip()
        if not text:
            return "#FFFFFF"
        if text.startswith("#"):
            text = text.upper()
            if len(text) == 7:
                return text
            if len(text) == 4:
                return "#" + "".join(ch * 2 for ch in text[1:])
            return "#FFFFFF"
        if re.fullmatch(r"[0-9A-Fa-f]{6}", text):
            return "#" + text.upper()
        return "#FFFFFF"

    @staticmethod
    def _schedule_foreground(colour: str) -> str:
        try:
            colour = colour.lstrip("#")
            r = int(colour[0:2], 16)
            g = int(colour[2:4], 16)
            b = int(colour[4:6], 16)
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            return "#1F2933" if luminance > 0.6 else "#FFFFFF"
        except Exception:
            return "#1F2933"

    def _selected_schedule_appointment(self) -> Optional[AppointmentDetail]:
        sel = self.appt_tree.selection()
        if not sel:
            return None
        return self.schedule_index.get(sel[0])

    def _open_patient_profile(self, _event=None) -> None:
        appt = self._selected_schedule_appointment()
        if not appt:
            messagebox.showinfo("Appointments", "Select an appointment to open the patient profile.")
            return
        if not appt.patient_id:
            messagebox.showwarning("Appointments", "This appointment is not linked to a patient record.")
            return
        self._show_patient_profile(appt.patient_id, appt.patient_name)

    def _schedule_selection_changed(self) -> None:
        appt = self._selected_schedule_appointment()
        if appt and appt.patient_id:
            self.open_patient_btn.state(["!disabled"])
            self.new_visit_btn.state(["!disabled"])
        else:
            self.open_patient_btn.state(["disabled"])
            self.new_visit_btn.state(["disabled"])
        if not appt or not appt.patient_id:
            self.timeline_tree.delete(*self.timeline_tree.get_children())
            return
        db = self._ensure_login_database()
        if not db:
            self.timeline_tree.delete(*self.timeline_tree.get_children())
            self.timeline_tree.insert("", "end", values=("No DB connection", "", "", "", "", ""))
            self.open_patient_btn.state(["disabled"])
            self.new_visit_btn.state(["disabled"])
            return
        try:
            reports = db.medic_reports_for_patient(appt.patient_id, limit=50)
        except Exception as exc:
            self.timeline_tree.delete(*self.timeline_tree.get_children())
            self.timeline_tree.insert("", "end", values=("Error loading history", "", "", "", "", str(exc)))
            return
        self.timeline_tree.delete(*self.timeline_tree.get_children())
        if not reports:
            self.timeline_tree.insert("", "end", values=("No visit notes found", "", "", "", "", ""))
            return
        for rep in reports:
            generated = self._format_datetime(rep.generated_on)
            apt_date = self._format_datetime(rep.appointment_date)
            diagnosis = _clean_note_text(rep.diagnosis) or "-"
            treatment = _clean_note_text(rep.treatment) or "-"
            history_snippet = _clean_note_text(rep.history or rep.notes_preview) or "-"
            if len(history_snippet) > 120:
                history_snippet = history_snippet[:117] + "..."
            if len(diagnosis) > 48:
                diagnosis = diagnosis[:45] + "..."
            if len(treatment) > 48:
                treatment = treatment[:45] + "..."
            self.timeline_tree.insert(
                "",
                "end",
                values=(generated, apt_date, rep.created_by or "-", diagnosis, treatment, history_snippet),
            )

    def _open_visit_from_schedule(self, _event: Optional[object] = None) -> None:
        appt = self._selected_schedule_appointment()
        if not appt:
            messagebox.showinfo("Visit Note", "Select an appointment first.")
            return
        self._open_visit_note_editor(appt)

    def _open_history_from_schedule(self, _event: Optional[object] = None) -> None:
        self._open_history_timeline()

    def _new_visit_note(self) -> None:
        appt = self._selected_schedule_appointment()
        if not appt:
            messagebox.showinfo("Visit Note", "Select an appointment first.")
            return
        if not appt.patient_id:
            messagebox.showwarning("Visit Note", "This appointment is not linked to a patient record.")
            return
        self._open_visit_note_editor(appt, force_new=True)

    def _open_history_timeline(self) -> None:
        appt = self._selected_schedule_appointment()
        if not appt or not appt.patient_id:
            messagebox.showinfo("Visit History", "Select a patient appointment first.")
            return
        db = self._ensure_login_database()
        if not db:
            messagebox.showerror("Visit History", "Database connection is not available.")
            return
        try:
            reports = db.medic_reports_for_patient(appt.patient_id, limit=200)
        except Exception as exc:
            messagebox.showerror("Visit History", f"Failed to load visit notes:\n{exc}")
            return
        if not reports:
            messagebox.showinfo("Visit History", "No visit notes recorded for this patient yet.")
            return

        window = tk.Toplevel(self)
        title = appt.patient_name or appt.patient_id or "Patient"
        window.title(f"Visit Timeline – {title}")
        try:
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            width = max(960, int(screen_w * 0.75))
            height = max(640, int(screen_h * 0.75))
            window.geometry(f"{width}x{height}")
            window.minsize(int(screen_w * 0.6), int(screen_h * 0.6))
        except Exception:
            window.geometry("1100x720")
            window.minsize(900, 600)
        window.transient(self)
        window.lift()

        heading = ttk.Frame(window, padding=(16, 16, 16, 8))
        heading.pack(fill="x")
        ttk.Label(heading, text=title, font=("Segoe UI Semibold", 13)).pack(anchor="w")
        ttk.Label(
            heading,
            text=f"{len(reports)} visit note(s)",
            foreground="#475467",
        ).pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(window, padding=(16, 0, 16, 16))
        body.pack(fill="both", expand=True)
        body.rowconfigure(1, weight=1)
        body.columnconfigure(0, weight=1)

        columns = ("generated", "appointment", "author", "diagnosis", "treatment", "history")
        tree = ttk.Treeview(body, columns=columns, show="headings", height=12, selectmode="browse")
        for name, label, width, anchor in (
            ("generated", "Generated", 160, tk.W),
            ("appointment", "Appointment", 160, tk.W),
            ("author", "Created By", 130, tk.W),
            ("diagnosis", "Diagnosis", 180, tk.W),
            ("treatment", "Treatment", 180, tk.W),
            ("history", "History", 240, tk.W),
        ):
            tree.heading(name, text=label)
            tree.column(name, width=width, anchor=anchor, stretch=False)
        tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=tree_scroll.set)

        preview_box = ttk.Labelframe(body, text="Summary", padding=12)
        preview_box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(12, 0))
        preview_box.columnconfigure(0, weight=1)
        preview_box.rowconfigure(0, weight=1)
        preview = tk.Text(preview_box, wrap="word", height=8, state="disabled")
        preview.grid(row=0, column=0, sticky="nsew")
        preview_scroll = ttk.Scrollbar(preview_box, orient="vertical", command=preview.yview)
        preview_scroll.grid(row=0, column=1, sticky="ns")
        preview.configure(yscrollcommand=preview_scroll.set)

        def show_preview(selected_id: str) -> None:
            note = note_index.get(selected_id)
            preview.configure(state="normal")
            preview.delete("1.0", "end")
            if note:
                sections = [
                    ("Generated", self._format_datetime(note.generated_on)),
                    ("Appointment", self._format_datetime(note.appointment_date)),
                    ("Created By", note.created_by or "-"),
                ]
                diagnosis = _clean_note_text(note.diagnosis)
                treatment = _clean_note_text(note.treatment)
                history = _clean_note_text(note.history)
                examination = _clean_note_text(note.examination)
                finding = _clean_note_text(note.finding)
                advice = _clean_note_text(note.advice)
                next_action = _clean_note_text(note.next_action)
                if diagnosis:
                    sections.append(("Diagnosis", diagnosis))
                if treatment:
                    sections.append(("Treatment", treatment))
                if history:
                    sections.append(("History", history))
                if examination:
                    sections.append(("Examination", examination))
                if finding:
                    sections.append(("Findings", finding))
                if advice:
                    sections.append(("Advice", advice))
                if next_action:
                    sections.append(("Next Action", next_action))
                text_lines: list[str] = []
                for label, content in sections:
                    text_lines.append(f"{label}:")
                    text_lines.append(content.strip())
                    text_lines.append("")
                preview.insert("1.0", "\n".join(text_lines).strip())
            preview.configure(state="disabled")

        note_index: dict[str, MedicReportSummary] = {}
        for rep in reports:
            iid = tree.insert(
                "",
                "end",
                values=(
                    self._format_datetime(rep.generated_on),
                    self._format_datetime(rep.appointment_date),
                    rep.created_by or "-",
                    (rep.diagnosis or "")[:64],
                    (rep.treatment or "")[:64],
                    (rep.history or rep.notes_preview or "")[:80],
                ),
            )
            note_index[iid] = rep
        if note_index:
            first_id = next(iter(note_index))
            tree.selection_set(first_id)
            show_preview(first_id)

        tree.bind("<<TreeviewSelect>>", lambda _e: show_preview(tree.selection()[0]) if tree.selection() else None)

        ttk.Button(window, text="Close", command=window.destroy, style="Ghost.TButton").pack(
            anchor="e", padx=16, pady=(0, 16)
        )

    @staticmethod
    def _parse_datetime(value: str) -> dt.datetime:
        text = (value or "").strip()
        if not text:
            raise ValueError("Date/time cannot be blank. Use YYYY-MM-DD HH:MM format.")
        patterns = ("%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M")
        for fmt in patterns:
            try:
                return dt.datetime.strptime(text, fmt)
            except ValueError:
                continue
        raise ValueError("Invalid date/time. Use format YYYY-MM-DD HH:MM.")

    def _open_visit_note_editor(self, appointment: AppointmentDetail, force_new: bool = False) -> None:
        if not appointment.patient_id:
            messagebox.showwarning("Visit Note", "This appointment is not linked to a patient record.")
            return
        db = self._ensure_login_database()
        if not db:
            messagebox.showerror("Visit Note", "Database connection is not available.")
            return
        existing: Optional[MedicReportDetail] = None
        if not force_new:
            try:
                existing = db.medic_report_for_appointment(appointment.patient_id, appointment.scheduled)
            except Exception as exc:
                messagebox.showerror("Visit Note", f"Failed to load visit note: {exc}")
                return
        categories = self._ensure_stock_categories()
        if not categories:
            messagebox.showinfo("Visit Note", "Stock catalogue is not available.")
            return
        prefill = self._prefill_visit_note_items
        prefill_patient = self._prefill_visit_note_patient_id
        self._prefill_visit_note_items = None
        self._prefill_visit_note_patient_id = None
        if appointment.patient_id and prefill_patient and appointment.patient_id != prefill_patient:
            prefill = None
        self._render_visit_note_dialog(db, appointment, categories, existing, prefill)

    def _render_visit_note_dialog(
        self,
        db: ClinicDatabase,
        appointment: AppointmentDetail,
        categories: list[str],
        existing: Optional[MedicReportDetail],
        prefill: Optional[list[VisitNoteItem]] = None,
    ) -> None:
        editing = existing is not None
        dialog = tk.Toplevel(self)
        dialog.title("Edit Visit Note" if editing else "New Visit Note")
        try:
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            width = max(1150, int(screen_w * 0.8))
            height = max(840, int(screen_h * 0.85))
            dialog.geometry(f"{width}x{height}")
            dialog.minsize(int(screen_w * 0.7), int(screen_h * 0.8))
        except Exception:
            dialog.geometry("1200x860")
            dialog.minsize(1024, 780)
        dialog.transient(self)
        dialog.grab_set()
        dialog.lift()

        patient_label = f"{appointment.patient_name or appointment.patient_id} ({appointment.patient_id})"
        header = ttk.Frame(dialog, padding=(16, 16, 16, 8))
        header.pack(fill="x")
        ttk.Label(header, text=("Edit Visit Note" if editing else "Create Visit Note"), font=("Segoe UI Semibold", 13)).grid(row=0, column=0, sticky="w")
        ttk.Label(header, text=patient_label, foreground="#475467").grid(row=1, column=0, sticky="w", pady=(4, 0))

        form = ttk.Frame(dialog, padding=(16, 0, 16, 12))
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        generated_default = (existing.generated_on.strftime("%Y-%m-%d %H:%M") if editing else dt.datetime.now().strftime("%Y-%m-%d %H:%M"))
        generated_var = tk.StringVar(value=generated_default)
        appointment_var = tk.StringVar(value=appointment.scheduled.strftime("%Y-%m-%d %H:%M"))
        diagnosis_var = tk.StringVar(value=_clean_note_text(existing.diagnosis) if editing else "")
        treatment_var = tk.StringVar(value=_clean_note_text(existing.treatment) if editing else "")

        ttk.Label(form, text="Generated (YYYY-MM-DD HH:MM)").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=generated_var, width=24).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(form, text="Appointment Time").grid(row=0, column=2, sticky="w", padx=(16, 0), pady=4)
        ttk.Entry(form, textvariable=appointment_var, width=24).grid(row=0, column=3, sticky="w", pady=4)

        ttk.Label(form, text="Diagnosis").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=diagnosis_var).grid(row=1, column=1, columnspan=3, sticky="we", pady=4)
        ttk.Label(form, text="Planned Treatment").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(form, textvariable=treatment_var).grid(row=2, column=1, columnspan=3, sticky="we", pady=4)

        history_box = ttk.Labelframe(dialog, text="History / Notes", padding=12)
        history_box.pack(fill="both", expand=False, padx=16, pady=(0, 12))
        history_text = tk.Text(history_box, height=4, wrap="word")
        history_text.pack(fill="both", expand=True)
        if editing:
            history_text.insert("1.0", _clean_note_text(existing.history))

        examination_box = ttk.Labelframe(dialog, text="Examination", padding=12)
        examination_box.pack(fill="both", expand=False, padx=16, pady=(0, 12))
        examination_text = tk.Text(examination_box, height=3, wrap="word")
        examination_text.pack(fill="both", expand=True)
        if editing:
            examination_text.insert("1.0", _clean_note_text(existing.examination))

        findings_box = ttk.Labelframe(dialog, text="Findings", padding=12)
        findings_box.pack(fill="both", expand=False, padx=16, pady=(0, 12))
        findings_text = tk.Text(findings_box, height=3, wrap="word")
        findings_text.pack(fill="both", expand=True)
        if editing:
            findings_text.insert("1.0", _clean_note_text(existing.finding))

        advice_box = ttk.Labelframe(dialog, text="Advice / Instructions", padding=12)
        advice_box.pack(fill="both", expand=False, padx=16, pady=(0, 12))
        advice_text = tk.Text(advice_box, height=3, wrap="word")
        advice_text.pack(fill="both", expand=True)
        if editing:
            advice_text.insert("1.0", _clean_note_text(existing.advice))

        next_box = ttk.Labelframe(dialog, text="Next Action / Plan", padding=12)
        next_box.pack(fill="both", expand=False, padx=16, pady=(0, 12))
        next_text = tk.Text(next_box, height=3, wrap="word")
        next_text.pack(fill="both", expand=True)
        if editing:
            next_text.insert("1.0", _clean_note_text(existing.next_action))

        visit_items: list[VisitNoteItem] = []
        visit_item_map: dict[str, VisitNoteItem] = {}
        current_stock_options: list[tuple[str, str, float]] = []
        current_receipt: Optional[Receipt] = None
        current_receipt_id: Optional[str] = None

        items_box = ttk.Labelframe(dialog, text="Items", padding=12)
        items_box.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        items_box.columnconfigure(1, weight=1)
        items_box.columnconfigure(3, weight=1)
        items_box.columnconfigure(5, weight=1)

        category_var = tk.StringVar()
        item_var = tk.StringVar()
        unit_price_value = tk.DoubleVar(value=0.0)
        qty_var = tk.IntVar(value=1)
        amount_label_var = tk.StringVar(value=fmt_money(0))

        ttk.Label(items_box, text="Category").grid(row=0, column=0, sticky="w", pady=(0, 4))
        category_combo = ttk.Combobox(
            items_box,
            textvariable=category_var,
            state="readonly",
            values=categories,
            width=32,
        )
        category_combo.grid(row=0, column=1, columnspan=5, sticky="we", pady=(0, 4))

        ttk.Label(items_box, text="Item").grid(row=1, column=0, sticky="w", pady=4)
        item_combo = ttk.Combobox(
            items_box,
            textvariable=item_var,
            state="readonly",
            width=48,
        )
        item_combo.grid(row=1, column=1, columnspan=5, sticky="we", pady=4)

        ttk.Label(items_box, text="Unit Price").grid(row=2, column=0, sticky="w", pady=4)
        unit_price_entry = ttk.Entry(items_box, textvariable=unit_price_value, width=12, justify="right")
        unit_price_entry.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(items_box, text="Quantity").grid(row=2, column=2, sticky="w", pady=4)
        qty_entry = ttk.Entry(items_box, textvariable=qty_var, width=6)
        qty_entry.grid(row=2, column=3, sticky="w", pady=4)
        qty_entry.bind("<KeyRelease>", lambda *_: update_amount_label())
        qty_entry.bind("<FocusOut>", lambda *_: update_amount_label())

        ttk.Label(items_box, text="Amount").grid(row=2, column=4, sticky="w", pady=4)
        amount_label = ttk.Label(items_box, textvariable=amount_label_var)
        amount_label.grid(row=2, column=5, sticky="w", pady=4)

        add_btn = ttk.Button(items_box, text="Add Item")
        add_btn.grid(row=2, column=6, sticky="e", padx=(8, 0), pady=4)

        tree_columns = ("item", "category", "qty", "unit", "amount")
        items_tree = ttk.Treeview(
            items_box,
            columns=tree_columns,
            show="headings",
            height=6,
            selectmode="browse",
        )
        for name, label, width, anchor in (
            ("item", "Item", 280, tk.W),
            ("category", "Category", 160, tk.W),
            ("qty", "Qty", 60, tk.CENTER),
            ("unit", "Unit Price", 100, tk.E),
            ("amount", "Amount", 110, tk.E),
        ):
            items_tree.heading(name, text=label)
            items_tree.column(name, width=width, anchor=anchor, stretch=False)
        items_tree.grid(row=3, column=0, columnspan=7, sticky="nsew", pady=(10, 0))
        items_box.rowconfigure(3, weight=1)
        tree_scroll = ttk.Scrollbar(items_box, orient="vertical", command=items_tree.yview)
        tree_scroll.grid(row=3, column=7, sticky="ns", pady=(10, 0))
        items_tree.configure(yscrollcommand=tree_scroll.set)

        remove_btn = ttk.Button(items_box, text="Remove Selected")
        remove_btn.grid(row=4, column=0, columnspan=7, sticky="e", pady=(8, 0))

        def close_dialog() -> None:
            dialog.grab_release()
            dialog.destroy()

        def _current_unit_price() -> Optional[float]:
            try:
                return float(unit_price_value.get())
            except (tk.TclError, ValueError):
                return None

        def update_amount_label(*_args) -> None:
            qty_text = qty_entry.get().strip()
            try:
                qty = int(qty_text or "0")
            except ValueError:
                qty = 0
            if qty < 1:
                qty = 1
            if qty_var.get() != qty:
                qty_var.set(qty)
            unit_price = _current_unit_price()
            amount = (unit_price or 0.0) * qty
            amount_label_var.set(fmt_money(amount))

        unit_price_value.trace_add("write", lambda *_: update_amount_label())

        def load_items_for_category(*_args) -> None:
            category = category_var.get().strip()
            options = self._stock_items_for_category(category, db) if category else []
            current_stock_options.clear()
            current_stock_options.extend(options)
            display_values = [name for _sid, name, _price in options]
            item_combo["values"] = display_values
            if display_values:
                item_combo.current(0)
                on_item_selected()
            else:
                item_var.set("")
                unit_price_value.set(0.0)
                update_amount_label()

        def on_item_selected(*_args) -> None:
            name = item_var.get().strip()
            for stock_id, stock_name, price in current_stock_options:
                if stock_name == name:
                    unit_price_value.set(price)
                    update_amount_label()
                    return
            unit_price_value.set(0.0)
            update_amount_label()

        def refresh_tree() -> None:
            items_tree.delete(*items_tree.get_children())
            visit_item_map.clear()
            for item in visit_items:
                iid = items_tree.insert(
                    "",
                    "end",
                    values=(
                        item.name,
                        item.category,
                        item.qty,
                        fmt_money(item.unit_price),
                        fmt_money(item.unit_price * item.qty),
                    ),
                )
                visit_item_map[iid] = item
            update_amount_label()

        def add_visit_item() -> None:
            category = category_var.get().strip()
            name = item_var.get().strip()
            if not category or not name:
                messagebox.showwarning("Items", "Select a category and item before adding.")
                return
            stock_id = ""
            unit_price = _current_unit_price()
            if unit_price is None:
                messagebox.showwarning("Items", "Enter a valid unit price.")
                unit_price_entry.focus_set()
                return
            for sid, stock_name, price in current_stock_options:
                if stock_name == name:
                    stock_id = sid
                    break
            if not stock_id:
                messagebox.showwarning("Items", "Unable to resolve the selected item.")
                return
            try:
                qty = int(qty_var.get())
                if qty < 1:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Items", "Quantity must be a positive number.")
                return
            visit_items.append(VisitNoteItem(stock_id=stock_id, name=name, category=category, unit_price=unit_price, qty=qty))
            refresh_tree()

        def remove_selected_item() -> None:
            sel = items_tree.selection()
            if not sel:
                return
            iid = sel[0]
            item = visit_item_map.pop(iid, None)
            if item and item in visit_items:
                visit_items.remove(item)
            refresh_tree()

        def load_receipt_items_for_note() -> bool:
            nonlocal current_receipt, current_receipt_id
            nonlocal current_receipt, current_receipt_id
            if not existing:
                return False
            try:
                receipts = db.receipts_for_medic_report(existing.report_id)
            except Exception as exc:
                messagebox.showwarning("Visit Note", f"Failed to load receipt items:\n{exc}")
                return False
            if not receipts:
                return False
            receipt = max(receipts, key=lambda r: r.issued)
            try:
                line_items = db.get_receipt_items(receipt.rcpt_id)
            except Exception as exc:
                messagebox.showwarning("Visit Note", f"Failed to load receipt items: {exc}")
                return False
            visit_items.clear()
            visit_item_map.clear()
            new_items: list[VisitNoteItem] = []
            for entry in line_items:
                stock_id = entry.item_id
                category = self._stock_category_for_item(stock_id, db)
                new_items.append(
                    VisitNoteItem(
                        stock_id=stock_id,
                        name=entry.name or stock_id or "Item",
                        category=category,
                        unit_price=entry.unit_price or 0.0,
                        qty=max(1, entry.qty),
                    )
                )
            visit_items.extend(new_items)
            self._prefill_visit_note_items = list(new_items)
            self._prefill_visit_note_patient_id = appointment.patient_id or ""
            current_receipt = receipt
            current_receipt_id = receipt.rcpt_id
            return bool(visit_items)

        def populate_from_existing() -> None:
            loaded = False
            if editing and existing:
                loaded = load_receipt_items_for_note()
            if not loaded and editing and existing and existing.chart_items:
                visit_items.clear()
                for chart in existing.chart_items:
                    stock_id = chart.stock_id or ""
                    if not stock_id and chart.notation_id:
                        try:
                            notation = db.notation_by_id(chart.notation_id)
                        except Exception:
                            notation = None
                        if notation and notation.stock_id:
                            stock_id = notation.stock_id
                    if not stock_id:
                        continue
                    category = self._stock_category_for_item(stock_id, db)
                    visit_items.append(
                        VisitNoteItem(
                            stock_id=stock_id,
                            name=chart.stock_name or chart.notation_title or stock_id or "Item",
                            category=category,
                            unit_price=float(chart.unit_price or 0.0),
                            qty=1,
                        )
                    )
                if visit_items:
                    self._prefill_visit_note_items = list(visit_items)
                    self._prefill_visit_note_patient_id = appointment.patient_id or ""
            elif not loaded and prefill:
                visit_items.extend(prefill)
                if visit_items:
                    self._prefill_visit_note_items = list(visit_items)
                    self._prefill_visit_note_patient_id = appointment.patient_id or ""
            refresh_tree()

        def close_dialog() -> None:
            dialog.grab_release()
            dialog.destroy()

        def visit_items_chart_entries() -> list[DentalChartItem]:
            entries: list[DentalChartItem] = []
            for item in visit_items:
                notation = None
                if item.stock_id:
                    try:
                        notation = db.notation_for_stock(item.stock_id)
                    except Exception:
                        notation = None
                if not notation:
                    continue
                remarks = f"Qty {max(1, item.qty)}"
                entries.append(
                    DentalChartItem(
                        notation_id=notation.notation_id,
                        tooth_id=0,
                        tooth_plan="E",
                        remarks=remarks,
                        unit_price=item.unit_price,
                        notation_status=1,
                        bill_status=0,
                        notation_title=notation.title or notation.stock_name or item.name,
                        stock_name=notation.stock_name or notation.title or item.name,
                        stock_id=notation.stock_id or item.stock_id,
                    )
                )
            return entries

        def build_receipt_items() -> Optional[list[ReceiptDraftItem]]:
            drafts: list[ReceiptDraftItem] = []
            for item in visit_items:
                qty = max(1, item.qty)
                if not item.stock_id:
                    messagebox.showwarning(
                        "Visit Note",
                        f"Item '{item.name}' is missing a stock code. Select a valid stock item before saving.",
                    )
                    return None
                try:
                    stock_details = db.stock_item_details(item.stock_id)
                except Exception as exc:
                    messagebox.showwarning(
                        "Visit Note",
                        f"Unable to verify stock item '{item.name}':\n{exc}",
                    )
                    return None
                if not stock_details:
                    messagebox.showwarning(
                        "Visit Note",
                        f"Stock item '{item.name}' no longer exists in the catalogue. Please choose another item.",
                    )
                    return None
                stock_id, stock_name, stock_price = stock_details
                unit_price = item.unit_price if item.unit_price is not None else stock_price
                subtotal = float(unit_price) * qty
                name = stock_name or item.name
                drafts.append(
                    ReceiptDraftItem(
                        stock_id=stock_id,
                        description=name,
                        qty=qty,
                        unit_price=float(unit_price),
                        subtotal=subtotal,
                        remark="",
                    )
                )
            return drafts

        def save_visit() -> None:
            nonlocal appointment, current_receipt, current_receipt_id
            if not visit_items:
                messagebox.showwarning("Visit Note", "Add at least one item before saving.")
                return
            try:
                generated_dt = self._parse_datetime(generated_var.get())
                appointment_dt = self._parse_datetime(appointment_var.get())
            except ValueError as exc:
                messagebox.showerror("Visit Note", str(exc))
                return
            history_plain = history_text.get("1.0", "end").strip()
            examination_plain = examination_text.get("1.0", "end").strip()
            findings_plain = findings_text.get("1.0", "end").strip()
            advice_plain = advice_text.get("1.0", "end").strip()
            next_plain = next_text.get("1.0", "end").strip()
            diagnosis_plain = diagnosis_var.get().strip()
            treatment_plain = treatment_var.get().strip()

            history_rtf = _text_to_rtf(history_plain)
            examination_rtf = _text_to_rtf(examination_plain)
            findings_rtf = _text_to_rtf(findings_plain)
            advice_rtf = _text_to_rtf(advice_plain)
            next_rtf = _text_to_rtf(next_plain)
            diagnosis_rtf = _text_to_rtf(diagnosis_plain)
            treatment_rtf = _text_to_rtf(treatment_plain)

            chart_entries = visit_items_chart_entries()
            try:
                if editing and existing:
                    db.update_medical_record(
                        existing.report_id,
                        generated_on=generated_dt,
                        appointment_on=appointment_dt,
                        username=self.session_user or existing.created_by or "",
                        history=history_rtf,
                        diagnosis=diagnosis_rtf,
                        treatment=treatment_rtf,
                        examination=examination_rtf,
                        finding=findings_rtf,
                        advice=advice_rtf,
                        next_action=next_rtf,
                        chart_items=chart_entries,
                    )
                    report_id = existing.report_id
                else:
                    report_id = db.create_medical_record(
                        patient_id=appointment.patient_id,
                        generated_on=generated_dt,
                        appointment_on=appointment_dt,
                        username=self.session_user or "",
                        history=history_rtf,
                        diagnosis=diagnosis_rtf,
                        treatment=treatment_rtf,
                        examination=examination_rtf,
                        finding=findings_rtf,
                        advice=advice_rtf,
                        next_action=next_rtf,
                        chart_items=chart_entries,
                    )
            except Exception as exc:
                messagebox.showerror("Visit Note", f"Failed to save visit note: {exc}")
                return

            receipt_drafts = build_receipt_items()
            if receipt_drafts is None:
                return
            subtotal_value = sum(d.subtotal for d in receipt_drafts)
            payment_code = current_receipt.payment_code if current_receipt else "AA"
            discount_value = current_receipt.discount if current_receipt else 0.0
            rounding_value = current_receipt.rounding if current_receipt else 0.0
            consult_value = current_receipt.consult_fees if current_receipt else 0.0
            department_value = (
                current_receipt.department_type
                if current_receipt and current_receipt.department_type
                else getattr(appointment, "department_type", "Clinic") or "Clinic"
            )
            receipt_id: Optional[str] = current_receipt_id
            receipt_error: Optional[str] = None
            had_existing_receipt = current_receipt is not None
            try:
                if current_receipt is not None and current_receipt_id:
                    db.replace_receipt(
                        rcpt_id=current_receipt.rcpt_id,
                        issued=generated_dt,
                        patient_id=appointment.patient_id,
                        items=receipt_drafts,
                        subtotal=subtotal_value,
                        discount=discount_value,
                        rounding=rounding_value,
                        consult_fees=consult_value,
                        remark=diagnosis_plain,
                        payment_code=payment_code,
                        username=self.session_user or current_receipt.done_by or "",
                        department=department_value,
                        mr_id=report_id,
                    )
                    total_value = subtotal_value - discount_value + rounding_value + consult_value
                    current_receipt = replace(
                        current_receipt,
                        issued=generated_dt,
                        subtotal=subtotal_value,
                        total=total_value,
                        payment_code=payment_code,
                        remark=diagnosis_plain,
                        discount=discount_value,
                        rounding=rounding_value,
                        consult_fees=consult_value,
                        department_type=department_value,
                        mr_id=report_id,
                    )
                    receipt_id = current_receipt.rcpt_id
                else:
                    new_receipt_id = db.create_receipt(
                        patient_id=appointment.patient_id,
                        issued=generated_dt,
                        username=self.session_user or "",
                        payment_code=payment_code,
                        items=receipt_drafts,
                        subtotal=subtotal_value,
                        discount=discount_value,
                        rounding=rounding_value,
                        consult_fees=consult_value,
                        remark=diagnosis_plain,
                        mr_id=report_id,
                        department=department_value,
                    )
                    total_value = subtotal_value - discount_value + rounding_value + consult_value
                    current_receipt = Receipt(
                        rcpt_id=new_receipt_id,
                        issued=generated_dt,
                        patient_id=appointment.patient_id,
                        total=total_value,
                        subtotal=subtotal_value,
                        gst=0.0,
                        payment_code=payment_code,
                        remark=diagnosis_plain,
                        discount=discount_value,
                        rounding=rounding_value,
                        consult_fees=consult_value,
                        done_by=self.session_user or "",
                        department_type=department_value,
                        settled_by=self.session_user or "",
                        tax_total=0.0,
                        mr_id=report_id,
                    )
                    receipt_id = new_receipt_id
                current_receipt_id = receipt_id
            except Exception as exc:
                receipt_error = str(exc)

            status_error: Optional[str] = None
            if appointment.status_id != 99:
                try:
                    original_appt = appointment
                    db.record_appointment_status(
                        appointment.patient_id,
                        appointment_dt,
                        3,
                        self.session_user or "",
                    )
                    updated_appt = replace(
                        original_appt,
                        status_id=3,
                        status="SETTLEMENT",
                        scheduled=appointment_dt,
                    )
                    appointment = updated_appt
                    self.settlement_selected_appointment = updated_appt
                    if self.schedule_appointments:
                        self.schedule_appointments = [
                            updated_appt if appt is original_appt else appt for appt in self.schedule_appointments
                        ]
                    for iid, appt in list(self.schedule_index.items()):
                        if appt is original_appt or (
                            appt.patient_id == original_appt.patient_id
                            and appt.scheduled == original_appt.scheduled
                        ):
                            self.schedule_index[iid] = updated_appt
                            values = list(self.appt_tree.item(iid, "values"))
                            if len(values) >= 5:
                                values[0] = appointment_dt.strftime("%H:%M")
                                values[4] = "SETTLEMENT"
                                self.appt_tree.item(iid, values=values)
                            break
                except Exception as exc:
                    status_error = str(exc)

            close_dialog()
            info_message = "Visit note updated." if editing else f"Visit note #{report_id} created."
            if not receipt_error and receipt_id:
                info_message += f"\nReceipt {receipt_id} {'updated' if had_existing_receipt else 'created'}."
            messagebox.showinfo("Visit Note", info_message)
            if receipt_error:
                messagebox.showwarning(
                    "Visit Note",
                    f"Visit note saved but failed to {'update' if had_existing_receipt else 'create'} receipt\n{receipt_error}",
                )
            self._schedule_selection_changed()
            if status_error:
                messagebox.showwarning(
                    "Visit Note",
                    f"Visit note saved but failed to update appointment status to SETTLEMENT\n{status_error}",
                )
            else:
                self.after(200, self._safe_refresh_schedule)
                self.after(250, self._safe_refresh_settlement)

        add_btn.configure(command=add_visit_item)
        remove_btn.configure(command=remove_selected_item)
        qty_var.trace_add("write", update_amount_label)
        category_combo.bind("<<ComboboxSelected>>", load_items_for_category)
        item_combo.bind("<<ComboboxSelected>>", on_item_selected)

        populate_from_existing()
        if not visit_items and categories:
            category_combo.current(0)
            load_items_for_category()

        action_bar = ttk.Frame(dialog, padding=(16, 0, 16, 16))
        action_bar.pack(fill="x")
        ttk.Button(action_bar, text="Cancel", command=close_dialog, style="Ghost.TButton").pack(side="right", padx=(8, 0))
        ttk.Button(action_bar, text="Save Visit Note", command=save_visit, style="Primary.TButton").pack(side="right")

        dialog.protocol("WM_DELETE_WINDOW", close_dialog)

        # Force large size and center AFTER layout (Tk can shrink to requested-size otherwise)
        def _force_large_and_center(win: tk.Toplevel) -> None:
            try:
                win.update_idletasks()
                sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
                # requested content size (in case it is larger)
                rw, rh = win.winfo_reqwidth(), win.winfo_reqheight()

                width  = max(int(sw * 0.80), rw, 1150)
                height = max(int(sh * 0.85), rh, 840)

                x = max((sw - width) // 2, 0)
                y = max((sh - height) // 2, 0)

                win.geometry(f"{width}x{height}+{x}+{y}")
                win.minsize(int(sw * 0.70), int(sh * 0.80))
            except Exception:
                pass

        # Apply twice (some late widgets adjust size after first draw)
        _force_large_and_center(dialog)
        dialog.after(150, lambda: _force_large_and_center(dialog))


    def _safe_refresh_schedule(self) -> None:
        try:
            self._load_schedule_for(self.schedule_date)
        except Exception:
            pass

    def _safe_refresh_settlement(self) -> None:
        try:
            self._load_settlement_list()
        except Exception:
            pass

    def _show_patient_profile(self, patient_id: str, display_name: str | None = None) -> None:
        db = self._ensure_login_database()
        if not db:
            messagebox.showerror("Patient Profile", "Database connection is not available.")
            return
        try:
            profile = db.get_patient_profile(patient_id)
            if not profile:
                messagebox.showinfo("Patient Profile", f"No patient record found for {patient_id}.")
                return
            allergies = db.allergies_for_patient(patient_id)
            documents = db.documents_for_patient(patient_id, limit=50)
            deposits, deposit_total = db.deposits_for_patient(patient_id)
        except Exception as exc:
            messagebox.showerror("Patient Profile", f"Failed to load patient data:\n{exc}")
            return

        window = tk.Toplevel(self)
        title_name = display_name or profile.name or profile.patient_id
        window.title(f"Patient Profile - {title_name}")
        window.geometry("760x560")
        window.minsize(680, 500)
        window.transient(self)
        window.lift()
        window.focus_force()

        nb = ttk.Notebook(window)
        nb.pack(fill="both", expand=True, padx=12, pady=12)

        # Details tab
        details_tab = ttk.Frame(nb, padding=12)
        nb.add(details_tab, text="Details")
        detail_grid = ttk.Frame(details_tab)
        detail_grid.pack(fill="x", expand=False)
        detail_grid.columnconfigure(1, weight=1)

        def _value(text: str) -> str:
            t = (text or "").strip()
            return t if t else "-"

        dob_display = self._format_date(profile.date_of_birth)
        if profile.date_of_birth and profile.date_of_birth.year > 1900:
            today = dt.date.today()
            age = today.year - profile.date_of_birth.year - (
                (today.month, today.day) < (profile.date_of_birth.month, profile.date_of_birth.day)
            )
            if age >= 0:
                dob_display = f"{dob_display} (Age {age})"

        address_line = ", ".join(
            filter(
                None,
                [profile.address, profile.city, profile.state, profile.postcode, profile.country],
            )
        )
        company_address = profile.company_address

        detail_rows = [
            ("Patient ID", profile.patient_id),
            ("Name", profile.name),
            ("Preferred Name", profile.preferred_name),
            ("Receipt Name", profile.receipt_name),
            ("Sex", profile.sex),
            ("Date of Birth", dob_display),
            ("Mobile", profile.phone_mobile),
            ("Phone", profile.phone_fixed),
            ("Email", profile.email),
            ("Address", address_line),
            ("Occupation", profile.occupation),
            ("Company", profile.company),
            ("Company Address", company_address),
            ("Company Contact", profile.company_contact),
            ("Billing Type", profile.billing_type),
            ("Emergency Contact", profile.emergency_contact),
            ("Emergency Phone", profile.emergency_phone),
            ("Registered On", self._format_datetime(profile.registered_on)),
            (
                "Last Updated",
                f"{self._format_datetime(profile.last_modified_on)} by {_value(profile.last_modified_by)}",
            ),
        ]

        for idx, (label, value) in enumerate(detail_rows):
            ttk.Label(detail_grid, text=label, foreground="#475467").grid(row=idx, column=0, sticky="nw", pady=4)
            ttk.Label(
                detail_grid,
                text=_value(value),
                wraplength=420,
                justify="left",
            ).grid(row=idx, column=1, sticky="w", pady=4, padx=(12, 0))

        remark_text = (profile.remark or "").strip()
        if remark_text:
            ttk.Separator(details_tab, orient="horizontal").pack(fill="x", pady=(16, 12))
            ttk.Label(details_tab, text="Remarks", font=("Segoe UI Semibold", 11)).pack(anchor="w")
            ttk.Label(details_tab, text=remark_text, wraplength=600, justify="left").pack(anchor="w", pady=(4, 0))

        medical_text = (profile.medical_illness or "").strip()
        if medical_text:
            ttk.Separator(details_tab, orient="horizontal").pack(fill="x", pady=(16, 12))
            ttk.Label(details_tab, text="Medical History", font=("Segoe UI Semibold", 11)).pack(anchor="w")
            ttk.Label(details_tab, text=medical_text, wraplength=600, justify="left").pack(anchor="w", pady=(4, 0))

        # Allergies tab
        allergies_tab = ttk.Frame(nb, padding=12)
        nb.add(allergies_tab, text=f"Allergies ({len(allergies)})")
        if allergies:
            columns = ("substance", "recorded_by", "modified_on")
            allergy_tree = ttk.Treeview(allergies_tab, columns=columns, show="headings", height=10)
            allergy_tree.heading("substance", text="Allergen")
            allergy_tree.heading("recorded_by", text="Recorded By")
            allergy_tree.heading("modified_on", text="Last Updated")
            allergy_tree.column("substance", width=180, anchor="w")
            allergy_tree.column("recorded_by", width=140, anchor="w")
            allergy_tree.column("modified_on", width=160, anchor="w")
            allergy_tree.pack(side="left", fill="both", expand=True)
            scroll = ttk.Scrollbar(allergies_tab, orient="vertical", command=allergy_tree.yview)
            scroll.pack(side="right", fill="y")
            allergy_tree.configure(yscrollcommand=scroll.set)
            for item in allergies:
                allergy_tree.insert(
                    "",
                    "end",
                    values=(
                        item.substance or "-",
                        item.recorded_by or "-",
                        self._format_datetime(item.modified_on),
                    ),
                )
        else:
            ttk.Label(allergies_tab, text="No allergies recorded for this patient.", foreground="#667085").pack(
                anchor="w"
            )

        # Documents tab
        documents_tab = ttk.Frame(nb, padding=12)
        nb.add(documents_tab, text=f"Documents ({len(documents)})")
        if documents:
            columns = ("title", "created", "effective", "author")
            doc_tree = ttk.Treeview(documents_tab, columns=columns, show="headings", height=10)
            doc_tree.heading("title", text="Title")
            doc_tree.heading("created", text="Created")
            doc_tree.heading("effective", text="Effective")
            doc_tree.heading("author", text="Author")
            doc_tree.column("title", width=280, anchor="w")
            doc_tree.column("created", width=120, anchor="w")
            doc_tree.column("effective", width=120, anchor="w")
            doc_tree.column("author", width=140, anchor="w")
            doc_tree.pack(side="left", fill="both", expand=True)
            scroll = ttk.Scrollbar(documents_tab, orient="vertical", command=doc_tree.yview)
            scroll.pack(side="right", fill="y")
            doc_tree.configure(yscrollcommand=scroll.set)
            for doc in documents:
                doc_tree.insert(
                    "",
                    "end",
                    values=(
                        doc.title or "-",
                        self._format_date(doc.created_on),
                        self._format_date(doc.effective_on),
                        doc.created_by or "-",
                    ),
                )
        else:
            ttk.Label(documents_tab, text="No documents attached to this patient.", foreground="#667085").pack(
                anchor="w"
            )

        # Deposits tab
        deposits_tab = ttk.Frame(nb, padding=12)
        nb.add(deposits_tab, text=f"Deposits ({len(deposits)})")
        if deposits:
            columns = ("date", "amount", "transaction", "method", "user")
            dep_tree = ttk.Treeview(deposits_tab, columns=columns, show="headings", height=10)
            for name, label, width in (
                ("date", "Date", 150),
                ("amount", "Amount", 100),
                ("transaction", "Transaction", 160),
                ("method", "Method", 80),
                ("user", "Recorded By", 140),
            ):
                dep_tree.heading(name, text=label)
                dep_tree.column(name, width=width, anchor="w" if name != "amount" else tk.E)
            dep_tree.pack(side="left", fill="both", expand=True)
            scroll = ttk.Scrollbar(deposits_tab, orient="vertical", command=dep_tree.yview)
            scroll.pack(side="right", fill="y")
            dep_tree.configure(yscrollcommand=scroll.set)
            for dep in deposits:
                dep_tree.insert(
                    "",
                    "end",
                    values=(
                        self._format_datetime(dep.created_on),
                        fmt_money(dep.amount),
                        dep.transaction or "-",
                        dep.payment_code or "-",
                        dep.recorded_by or "-",
                    ),
                )
            ttk.Label(
                deposits_tab,
                text=f"Net Deposits: {fmt_money(deposit_total)}",
                font=("Segoe UI Semibold", 10),
            ).pack(anchor="e", pady=(8, 0))
        else:
            ttk.Label(deposits_tab, text="No deposit transactions recorded.", foreground="#667085").pack(anchor="w")

        btns = ttk.Frame(window, padding=(12, 0, 12, 12))
        btns.pack(fill="x")
        ttk.Button(btns, text="Close", command=window.destroy).pack(side="right")

    @staticmethod
    def _format_date(value: dt.date | dt.datetime | None) -> str:
        if not value:
            return "-"
        if isinstance(value, dt.datetime):
            value = value.date()
        try:
            if value.year <= 1900:
                return "-"
            return value.strftime("%d %b %Y")
        except Exception:
            return "-"

    @staticmethod
    def _format_datetime(value: dt.datetime | None) -> str:
        if not value:
            return "-"
        try:
            if value.year <= 1900:
                return "-"
            return value.strftime("%d %b %Y %H:%M")
        except Exception:
            return "-"

    # ---------------- Receipts tab
    def _build_receipts_tab(self) -> None:
        top = ttk.Frame(self.receipts_tab)
        top.pack(fill="x", pady=(0, 10))

        # Date is driven entirely by shared calendar selection.
        self.date_var = self.shared_date_var
        ttk.Label(top, text="IC / Passport (optional)").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.ic_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.ic_var, width=24).grid(row=0, column=1, sticky="w")
        ttk.Button(top, text="Search", style="Primary.TButton", command=self._search).grid(row=0, column=2, padx=(12, 16))

        # Actions
        self.generate_btn = ttk.Button(top, text="Generate PDF", style="Ghost.TButton",
                                    command=self._generate_pdf, state=tk.DISABLED)
        self.print_btn = ttk.Button(top, text="Print", style="Ghost.TButton",
                                    command=self._print_pdf, state=tk.DISABLED)
        self.whatsapp_btn = ttk.Button(top, text="WhatsApp", style="Ghost.TButton",
                                    command=self._whatsapp, state=tk.DISABLED)
        self.email_btn = ttk.Button(top, text="Email", style="Ghost.TButton",
                                    command=self._email, state=tk.DISABLED)
        self.generate_btn.grid(row=0, column=3, padx=(0, 8))
        self.print_btn.grid(row=0, column=4, padx=(0, 8))
        self.whatsapp_btn.grid(row=0, column=5, padx=(0, 8))
        self.email_btn.grid(row=0, column=6)


        container = ttk.Frame(self.receipts_tab)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        calendar_column = ttk.Frame(container)
        calendar_column.grid(row=0, column=0, sticky="nw", padx=(0, 16))
        calendar_section = self._create_calendar_section(calendar_column, show_selected=False)
        calendar_section.pack(fill="x")

        top_frame = ttk.Frame(container)
        top_frame.grid(row=0, column=1, sticky="nsew")
        top_frame.columnconfigure(0, weight=1)
        top_frame.rowconfigure(0, weight=1)

        columns = ("issued", "rcpt", "patient", "ic", "method", "total", "paid", "balance", "payments")
        self.receipt_tree = ttk.Treeview(top_frame, columns=columns, show="headings", height=12, selectmode="browse")
        for name, label, width, anchor in (
            ("issued", "Date / Time", 150, tk.W),
            ("rcpt", "Receipt #", 120, tk.W),
            ("patient", "Patient", 240, tk.W),
            ("ic", "IC / Passport", 140, tk.W),
            ("method", "Payment", 120, tk.W),
            ("total", "Total", 110, tk.E),
            ("paid", "Paid", 110, tk.E),
            ("balance", "Balance", 110, tk.E),
            ("payments", "Payments", 90, tk.CENTER),
        ):
            self.receipt_tree.heading(name, text=label)
            self.receipt_tree.column(name, width=width, anchor=anchor, stretch=False)
        self.receipt_tree.grid(row=0, column=0, sticky="nsew")
        receipts_scroll = ttk.Scrollbar(top_frame, orient="vertical", command=self.receipt_tree.yview)
        receipts_scroll.grid(row=0, column=1, sticky="ns")
        self.receipt_tree.configure(yscrollcommand=receipts_scroll.set)

        bottom_frame = ttk.Frame(container)
        bottom_frame.grid(row=1, column=1, sticky="nsew", pady=(10, 0))
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(2, weight=1)
        bottom_frame.rowconfigure(0, weight=1)

        self.items_tree = ttk.Treeview(
            bottom_frame,
            columns=("desc", "qty", "unit", "amount"),
            show="headings",
            height=6,
            selectmode="browse",
        )
        for name, label, width, anchor in (
            ("desc", "Description", 360, tk.W),
            ("qty", "Qty", 60, tk.CENTER),
            ("unit", "Unit Price", 120, tk.E),
            ("amount", "Amount", 120, tk.E),
        ):
            self.items_tree.heading(name, text=label)
            self.items_tree.column(name, width=width, anchor=anchor, stretch=False)
        self.items_tree.grid(row=0, column=0, sticky="nsew")
        items_scroll = ttk.Scrollbar(bottom_frame, orient="vertical", command=self.items_tree.yview)
        items_scroll.grid(row=0, column=1, sticky="ns")
        self.items_tree.configure(yscrollcommand=items_scroll.set)
        self.items_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_receipt_item_select())

        bar = ttk.Frame(bottom_frame)
        bar.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.receipt_edit_btn = ttk.Button(
            bar,
            text="Edit Item",
            style="Ghost.TButton",
            command=self._edit_receipt_item,
            state=tk.DISABLED,
        )
        self.receipt_edit_btn.pack(side="right", padx=(0, 4))
        ttk.Label(bar, text="Total:").pack(side="left", padx=(2, 6))
        self.items_total_var = tk.StringVar(value=fmt_money(0))
        ttk.Label(bar, textvariable=self.items_total_var).pack(side="left")

        self.pay_tree = ttk.Treeview(
            bottom_frame,
            columns=("seq", "date", "amount", "method", "paid_to_date", "balance"),
            show="headings",
            height=7,
            selectmode="browse",
        )
        for name, label, width, anchor in (
            ("seq", "#", 60, tk.CENTER),
            ("date", "Date", 120, tk.W),
            ("amount", "Amount", 120, tk.E),
            ("method", "Method", 160, tk.W),
            ("paid_to_date", "Paid To Date", 130, tk.E),
            ("balance", "Balance", 120, tk.E),
        ):
            self.pay_tree.heading(name, text=label)
            self.pay_tree.column(name, width=width, anchor=anchor, stretch=False)
        self.pay_tree.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        pay_scroll = ttk.Scrollbar(bottom_frame, orient="vertical", command=self.pay_tree.yview)
        pay_scroll.grid(row=0, column=3, sticky="ns")
        self.pay_tree.configure(yscrollcommand=pay_scroll.set)
        pay_bar = ttk.Frame(bottom_frame)
        pay_bar.grid(row=1, column=2, sticky="ew", pady=(6, 0), padx=(10, 0))
        self.pay_edit_btn = ttk.Button(
            pay_bar,
            text="Edit Payment",
            style="Ghost.TButton",
            command=self._edit_receipt_payment,
            state=tk.DISABLED,
        )
        self.pay_edit_btn.pack(side="right")

    # ---------------- Settings tab
    def _build_settings_tab(self) -> None:
        s = self.settings_tab

        clinic_box = card(s, "Clinic Profile")
        clinic_box.pack(fill="x", pady=(0, 12))

        self.clinic_name = tk.StringVar(value=self.cfg.settings.clinic.name)
        self.clinic_phone = tk.StringVar(value=self.cfg.settings.clinic.phone)
        self.clinic_email = tk.StringVar(value=self.cfg.settings.clinic.email)

        _row(clinic_box, "Clinic Name", ttk.Entry(clinic_box, textvariable=self.clinic_name), row=0, col=0)
        _row(clinic_box, "Phone", ttk.Entry(clinic_box, textvariable=self.clinic_phone), row=0, col=2)
        _row(clinic_box, "Email", ttk.Entry(clinic_box, textvariable=self.clinic_email), row=1, col=0)

        ttk.Label(clinic_box, text="Address").grid(row=2, column=0, sticky="w", pady=(6, 2))
        self.clinic_address = tk.Text(clinic_box, height=4, width=48)
        self.clinic_address.insert("1.0", self.cfg.settings.clinic.address)
        self.clinic_address.grid(row=3, column=0, sticky="we", padx=(0, 12))

        ttk.Label(clinic_box, text="Logo").grid(row=2, column=2, sticky="w", pady=(6, 2))
        self.logo_var = tk.StringVar(value=self.cfg.settings.clinic.logo_path)
        logo_entry = ttk.Entry(clinic_box, textvariable=self.logo_var)
        logo_entry.grid(row=3, column=2, sticky="we")
        ttk.Button(clinic_box, text="📁", width=3, style="Ghost.TButton", command=self._choose_logo).grid(row=3, column=3, padx=(6, 0))

        ttk.Label(clinic_box, text="Receipt Output Folder").grid(row=4, column=2, sticky="w", pady=(6, 2))
        self.output_dir_var = tk.StringVar(value=self.cfg.settings.receipt.output_directory)
        out_entry = ttk.Entry(clinic_box, textvariable=self.output_dir_var)
        out_entry.grid(row=5, column=2, sticky="we")
        ttk.Button(clinic_box, text="📁", width=3, style="Ghost.TButton", command=self._choose_output_dir).grid(row=5, column=3, padx=(6, 0))

        clinic_box.grid_columnconfigure(0, weight=1)
        clinic_box.grid_columnconfigure(2, weight=1)

        ds = card(s, "Data Source")
        ds.pack(fill="x", pady=(0, 12))

        self.data_source_var = tk.StringVar(
            value="Live MySQL connection" if self.cfg.settings.database.source == "mysql" else "Backup file"
        )
        self._source_map = {"Live MySQL connection": "mysql", "Backup file": "backup"}

        ttk.Label(ds, text="Data Source").grid(row=0, column=0, sticky="w")
        self.data_source = ttk.Combobox(ds, state="readonly",
                                        values=list(self._source_map.keys()),
                                        textvariable=self.data_source_var, width=28)
        self.data_source.grid(row=0, column=1, sticky="w", padx=(8, 18))

        self.backup_path_var = tk.StringVar(value=self.cfg.settings.database.backup_path)
        _row(ds, "Backup File", ttk.Entry(ds, textvariable=self.backup_path_var, width=60),
             row=1, col=0, trailing=ttk.Button(ds, text="...", style="Ghost.TButton", command=self._choose_backup))

        self.mysql_host = tk.StringVar(value=self.cfg.settings.database.mysql_host)
        self.mysql_port = tk.IntVar(value=self.cfg.settings.database.mysql_port)
        self.mysql_user = tk.StringVar(value=self.cfg.settings.database.mysql_user)
        self.mysql_pass = tk.StringVar(value=self.cfg.settings.database.mysql_password)
        self.mysql_db = tk.StringVar(value=self.cfg.settings.database.mysql_database)

        _row(ds, "Host", ttk.Entry(ds, textvariable=self.mysql_host), row=2, col=0)
        _row(ds, "Port", ttk.Entry(ds, textvariable=self.mysql_port), row=2, col=2)
        _row(ds, "User", ttk.Entry(ds, textvariable=self.mysql_user), row=3, col=0)
        _row(ds, "Password", ttk.Entry(ds, textvariable=self.mysql_pass, show="•"), row=3, col=2)
        _row(ds, "Database", ttk.Entry(ds, textvariable=self.mysql_db), row=4, col=0)

        ds.grid_columnconfigure(1, weight=1)
        ds.grid_columnconfigure(3, weight=1)

        email_box = card(s, "Email Delivery (Gmail)")
        email_box.pack(fill="x", pady=(0, 12))

        self.email_sender = tk.StringVar(value=self.cfg.settings.email.sender)
        self.email_app_pw = tk.StringVar(value=self.cfg.settings.email.app_password)
        self.email_subject = tk.StringVar(value=self.cfg.settings.email.subject)

        _row(email_box, "Gmail Address", ttk.Entry(email_box, textvariable=self.email_sender, width=60), row=0, col=0)
        _row(email_box, "App Password", ttk.Entry(email_box, textvariable=self.email_app_pw, show="•", width=30), row=1, col=0)
        _row(email_box, "Subject", ttk.Entry(email_box, textvariable=self.email_subject, width=60), row=2, col=0)

        ttk.Label(email_box, text="Body").grid(row=3, column=0, sticky="w", pady=(6, 2))
        self.email_body = tk.Text(email_box, height=6, width=80)
        self.email_body.insert("1.0", self.cfg.settings.email.body)
        self.email_body.grid(row=4, column=0, columnspan=3, sticky="we")
        ttk.Label(email_box, text="You can use {patient_name} and {receipt_id} in the message.",
                  foreground="#667085").grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))

        email_box.grid_columnconfigure(1, weight=1)

        btns = ttk.Frame(s)
        btns.pack(fill="x")
        ttk.Button(btns, text="Save Settings", style="Primary.TButton", command=self._save_settings).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="Reconnect DB", style="Ghost.TButton", command=self._reconnect).pack(side="right")

    # ---------------- events / services
    def _attach_events(self) -> None:
        self.receipt_tree.bind("<<TreeviewSelect>>", lambda _e: self._load_selected_detail())
        self.pay_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_payment_select())
        self.data_source.bind("<<ComboboxSelected>>", lambda _e: self._on_source_changed())

    def _init_backing_services(self) -> None:
        if not self.db:
            try:
                ms = self.cfg.mysql_settings()
                self.db = ClinicDatabase(**ms)
            except Exception as exc:
                messagebox.showerror("Database", f"Failed to connect to MySQL: {exc}")
                self.db = None
        self._ensure_pdf_generator()
        self._load_schedule_for(self.schedule_date)

    # ---------------- receipts logic
    def _set_today(self) -> None:
        self._set_shared_date(dt.date.today())

    def _search(self) -> None:
        if not self.db:
            messagebox.showwarning("Database", "No database connection.")
            return

        target_date = parse_date(self.date_var.get())
        ic = self.ic_var.get().strip() or None

        self.status_var.set("Searching...")
        self.receipt_tree.delete(*self.receipt_tree.get_children())
        self.receipt_index.clear()
        self.items_tree.delete(*self.items_tree.get_children())
        self.pay_tree.delete(*self.pay_tree.get_children())
        self.items_total_var.set(fmt_money(0))
        self.receipt_edit_summary = None
        self.receipt_edit_items = []
        self.receipt_item_map = {}
        self.receipt_payments = []
        self.receipt_payment_map = {}
        self.selected_payment = None
        self._update_action_buttons(None)
        self._update_receipt_item_buttons()
        self._update_receipt_payment_buttons()

        try:
            summaries = self.db.receipts_for_date(target_date, ic)
        except Exception as exc:
            self.status_var.set("")
            messagebox.showerror("Search", f"Failed to query receipts: {exc}")
            return

        id_list = [s.receipt.rcpt_id for s in summaries]
        try:
            pay_map = self.db.partial_payments_for_receipts(id_list)
        except Exception:
            pay_map = {}

        for s in summaries:
            rid = s.receipt.rcpt_id
            issued = s.receipt.issued.strftime("%Y-%m-%d %H:%M")
            patient_name = self._patient_display_name(s.patient)
            total = s.receipt.total or 0.0
            method_label = self._payment_label_for_code(s.receipt.payment_code)

            payments = pay_map.get(rid, [])
            paid = sum(p.amount for p in payments)
            balance = max(total - paid, 0.0)

            iid = self.receipt_tree.insert(
                "", "end",
                values=(issued, rid, patient_name, s.patient.icpassport,
                        method_label, fmt_money(total), fmt_money(paid), fmt_money(balance), str(len(payments))),
            )
            self.receipt_index[iid] = s

        self.status_var.set(f"{len(summaries)} receipt(s)")

    def _load_selected_detail(self) -> None:
        self.items_tree.delete(*self.items_tree.get_children())
        self.pay_tree.delete(*self.pay_tree.get_children())
        self.items_total_var.set(fmt_money(0))
        self.selected_payment = None
        self.receipt_edit_summary = None
        self.receipt_edit_items = []
        self.receipt_item_map = {}
        self.receipt_payments = []
        self.receipt_payment_map = {}
        self._update_receipt_item_buttons()
        self._update_receipt_payment_buttons()

        summary = self._selected_summary()
        self._update_action_buttons(summary)
        if not (self.db and summary):
            return

        self.receipt_edit_summary = summary
        rid = summary.receipt.rcpt_id

        try:
            items: list[ReceiptItem] = self.db.get_receipt_items(rid)
        except Exception as exc:
            messagebox.showerror("Items", f"Failed to load items: {exc}")
            return

        total_items = 0.0
        editable_items: list[_ReceiptEditableItem] = []
        for it in items:
            qty = max(1, int(it.qty or 1))
            unit_price = float(it.unit_price or 0.0)
            amount = float(it.subtotal or (qty * unit_price))
            total_items += amount
            editable = _ReceiptEditableItem(
                stock_id=(it.item_id or "").strip(),
                description=it.name or (it.item_id or "Item"),
                qty=qty,
                unit_price=unit_price,
                remark=it.remark or "",
            )
            editable_items.append(editable)
            iid = self.items_tree.insert(
                "",
                "end",
                values=(editable.description, qty, fmt_money(unit_price), fmt_money(amount)),
            )
            self.receipt_item_map[iid] = len(editable_items) - 1
        self.receipt_edit_items = editable_items
        self.items_total_var.set(fmt_money(total_items))
        self._update_receipt_item_buttons()

        try:
            pmap = self.db.partial_payments_for_receipts([rid]) or {}
        except Exception as exc:
            messagebox.showerror("Payments", f"Failed to load payments: {exc}")
            return

        pays = pmap.get(rid, [])
        receipt_sel = self.receipt_tree.selection()
        tree_iid = receipt_sel[0] if receipt_sel else None
        self._render_receipt_payments(summary, pays, tree_iid=tree_iid)

    def _on_receipt_item_select(self) -> None:
        self._update_receipt_item_buttons()

    def _update_receipt_item_buttons(self) -> None:
        button = getattr(self, "receipt_edit_btn", None)
        tree = getattr(self, "items_tree", None)
        if not button or tree is None:
            return
        selection = tree.selection() if tree else ()
        if selection and selection[0] in self.receipt_item_map and self.receipt_edit_summary:
            button.configure(state=tk.NORMAL)
        else:
            button.configure(state=tk.DISABLED)

    def _render_receipt_payments(
        self,
        summary: ReceiptSummary,
        payments: Sequence[PartialPayment],
        *,
        tree_iid: str | None = None,
        select_sequence: int | None = None,
    ) -> None:
        self.pay_tree.delete(*self.pay_tree.get_children())
        self.receipt_payments = list(payments)
        self.receipt_payment_map = {}
        total_due = float(summary.receipt.total or 0.0)
        running = 0.0
        select_iid: str | None = None
        for idx, payment in enumerate(self.receipt_payments, start=1):
            amount = float(payment.amount or 0.0)
            running += amount
            balance = max(total_due - running, 0.0)
            display_method = payment.method or payment.pay_code
            date_label = payment.date.strftime("%Y-%m-%d")
            iid = self.pay_tree.insert(
                "",
                "end",
                values=(
                    f"#{idx}",
                    date_label,
                    fmt_money(amount),
                    display_method,
                    fmt_money(running),
                    fmt_money(balance),
                ),
            )
            self.receipt_payment_map[iid] = idx - 1
            if select_sequence is not None and idx == select_sequence:
                select_iid = iid
        if select_iid:
            self.pay_tree.selection_set(select_iid)
            self.pay_tree.focus(select_iid)
            self._on_payment_select()
        else:
            self.pay_tree.selection_remove(self.pay_tree.selection())
            self.selected_payment = None
            self._update_receipt_payment_buttons()
        if tree_iid:
            total_paid = running
            balance_value = max(total_due - total_paid, 0.0)
            method_label = self._payment_label_for_code(summary.receipt.payment_code)
            patient_name = self._patient_display_name(summary.patient)
            self.receipt_tree.item(
                tree_iid,
                values=(
                    summary.receipt.issued.strftime("%Y-%m-%d %H:%M"),
                    summary.receipt.rcpt_id,
                    patient_name,
                    summary.patient.icpassport,
                    method_label,
                    fmt_money(summary.receipt.total or 0.0),
                    fmt_money(total_paid),
                    fmt_money(balance_value),
                    str(len(self.receipt_payments)),
                ),
            )
            self.receipt_index[tree_iid] = summary

    def _update_receipt_payment_buttons(self) -> None:
        button = getattr(self, "pay_edit_btn", None)
        tree = getattr(self, "pay_tree", None)
        if not button or tree is None:
            return
        selection = tree.selection() if tree else ()
        if (
            selection
            and selection[0] in self.receipt_payment_map
            and self.receipt_edit_summary
            and self.receipt_payments
        ):
            button.configure(state=tk.NORMAL)
        else:
            button.configure(state=tk.DISABLED)

    def _edit_receipt_item(self) -> None:
        tree = getattr(self, "items_tree", None)
        if tree is None:
            return
        selection = tree.selection()
        if not selection:
            return
        iid = selection[0]
        idx = self.receipt_item_map.get(iid)
        summary = self.receipt_edit_summary or self._selected_summary()
        if idx is None or summary is None or idx >= len(self.receipt_edit_items):
            return
        current = self.receipt_edit_items[idx]
        db = self._ensure_login_database()
        if not db:
            messagebox.showerror("Receipts", "Database connection is unavailable.")
            return
        existing_category = self._stock_category_for_item(current.stock_id, db)
        existing = _SettlementItem(
            stock_id=current.stock_id,
            description=current.description,
            qty=current.qty,
            unit_price=current.unit_price,
            notation_id=0,
            tooth_label="",
            remarks=current.remark,
            source=existing_category or "receipt",
        )
        updated = self._open_settlement_item_dialog(existing=existing)
        if not updated:
            return
        new_item = _ReceiptEditableItem(
            stock_id=updated.stock_id,
            description=updated.description,
            qty=updated.qty,
            unit_price=updated.unit_price,
            remark=updated.remarks,
        )
        receipt_selection = self.receipt_tree.selection()
        if not receipt_selection:
            return
        new_items = list(self.receipt_edit_items)
        new_items[idx] = new_item
        if not self._persist_receipt_item_changes(summary, new_items, receipt_selection[0]):
            return
        self.receipt_edit_items = new_items
        self.receipt_item_map[iid] = idx
        amount = new_item.qty * new_item.unit_price
        tree.item(
            iid,
            values=(
                new_item.description,
                new_item.qty,
                fmt_money(new_item.unit_price),
                fmt_money(amount),
            ),
        )
        tree.selection_set(iid)
        tree.focus(iid)
        self._update_receipt_item_buttons()

    def _edit_receipt_payment(self) -> None:
        tree = getattr(self, "pay_tree", None)
        if tree is None:
            return
        selection = tree.selection()
        if not selection:
            return
        iid = selection[0]
        idx = self.receipt_payment_map.get(iid)
        summary = self.receipt_edit_summary or self._selected_summary()
        if idx is None or summary is None or idx < 0:
            return
        if idx >= len(self.receipt_payments):
            return
        payment = self.receipt_payments[idx]
        amount_prompt = simpledialog.askstring(
            "Edit Payment",
            "Enter paid amount (RM):",
            initialvalue=f"{float(payment.amount or 0.0):.2f}",
            parent=self,
        )
        if amount_prompt is None:
            return
        try:
            new_amount = float(amount_prompt.strip())
        except (TypeError, ValueError):
            messagebox.showwarning("Payments", "Enter a valid numeric amount.", parent=self)
            return
        if new_amount <= 0:
            messagebox.showwarning("Payments", "Amount must be greater than zero.", parent=self)
            return
        if abs(new_amount - float(payment.amount or 0.0)) < 0.005:
            return
        db = self._ensure_login_database()
        if not db:
            messagebox.showerror("Payments", "Database connection is unavailable.")
            return
        try:
            db.update_partial_payment_amount(payment.payment_id, amount=new_amount)
            pay_map = db.partial_payments_for_receipts([summary.receipt.rcpt_id]) or {}
            refreshed = pay_map.get(summary.receipt.rcpt_id, [])
        except Exception as exc:
            messagebox.showerror("Payments", f"Failed to update payment:\n{exc}", parent=self)
            return
        receipt_sel = self.receipt_tree.selection()
        tree_iid = receipt_sel[0] if receipt_sel else None
        self._render_receipt_payments(summary, refreshed, tree_iid=tree_iid, select_sequence=idx + 1)
        self.status_var.set(f"Payment #{idx + 1} updated.")

    def _persist_receipt_item_changes(
        self,
        summary: ReceiptSummary,
        items: list[_ReceiptEditableItem],
        tree_iid: str,
    ) -> bool:
        previous_selection = self.selected_payment.sequence if self.selected_payment else None
        db = self._ensure_login_database()
        if not db:
            messagebox.showerror("Receipts", "Database connection is unavailable.")
            return False
        if not items:
            messagebox.showwarning("Receipts", "Receipts must contain at least one item.")
            return False
        drafts: list[ReceiptDraftItem] = []
        subtotal = 0.0
        for item in items:
            if not (item.stock_id or "").strip():
                messagebox.showwarning("Receipts", "Each item must reference a valid stock code.")
                return False
            try:
                qty = max(1, int(item.qty or 0))
            except Exception:
                messagebox.showwarning("Receipts", "Quantity must be a positive integer.")
                return False
            try:
                unit_price = float(item.unit_price or 0.0)
            except Exception:
                messagebox.showwarning("Receipts", "Unit price must be numeric.")
                return False
            amount = float(qty * unit_price)
            subtotal += amount
            drafts.append(
                ReceiptDraftItem(
                    stock_id=item.stock_id,
                    description=item.description,
                    qty=qty,
                    unit_price=unit_price,
                    subtotal=amount,
                    remark=item.remark,
                )
            )
        discount = float(summary.receipt.discount or 0.0)
        rounding = float(summary.receipt.rounding or 0.0)
        consult = float(summary.receipt.consult_fees or 0.0)
        remark = summary.receipt.remark or ""
        payment_code = summary.receipt.payment_code or ""
        username = self.session_user or summary.receipt.done_by or ""
        department = summary.receipt.department_type or "Clinic"
        mr_id = summary.receipt.mr_id or 0
        try:
            db.replace_receipt(
                summary.receipt.rcpt_id,
                issued=summary.receipt.issued,
                patient_id=summary.receipt.patient_id,
                items=drafts,
                subtotal=subtotal,
                discount=discount,
                rounding=rounding,
                consult_fees=consult,
                remark=remark,
                payment_code=payment_code,
                username=username,
                department=department,
                mr_id=mr_id,
            )
        except Exception as exc:
            messagebox.showerror("Receipts", f"Failed to update receipt:\n{exc}")
            return False
        total = subtotal + consult - discount + rounding
        try:
            pay_map = db.partial_payments_for_receipts([summary.receipt.rcpt_id]) or {}
            pay_list = pay_map.get(summary.receipt.rcpt_id, [])
        except Exception as exc:
            messagebox.showwarning("Receipts", f"Updated receipt but failed to refresh payments:\n{exc}")
            pay_list = []
        updated_receipt = replace(summary.receipt, subtotal=subtotal, total=total)
        new_summary = ReceiptSummary(receipt=updated_receipt, patient=summary.patient)
        self._render_receipt_payments(
            new_summary,
            pay_list,
            tree_iid=tree_iid,
            select_sequence=previous_selection,
        )
        self.receipt_edit_summary = new_summary
        self.items_total_var.set(fmt_money(subtotal))
        self.status_var.set(f"Receipt {summary.receipt.rcpt_id} updated.")
        return True

    def _on_payment_select(self) -> None:
        sel = self.pay_tree.selection()
        if not sel:
            self.selected_payment = None
            self._update_receipt_payment_buttons()
            return

        idx = self.receipt_payment_map.get(sel[0])
        summary = self.receipt_edit_summary or self._selected_summary()
        if idx is None or summary is None or idx < 0:
            self.selected_payment = None
            self._update_receipt_payment_buttons()
            return
        if idx >= len(self.receipt_payments):
            self.selected_payment = None
            self._update_receipt_payment_buttons()
            return
        payments = self.receipt_payments
        total_due = float(summary.receipt.total or 0.0)
        current_payment = payments[idx]
        running = sum(float(p.amount or 0.0) for p in payments[: idx + 1])
        self.selected_payment = _PaymentProgressState(
            sequence=idx + 1,
            amount=float(current_payment.amount or 0.0),
            total_due=total_due,
            paid_on=current_payment.date,
            balance=max(total_due - running, 0.0),
        )
        self._update_receipt_payment_buttons()

    # ---------------- actions
    def _update_action_buttons(self, summary: ReceiptSummary | None) -> None:
        state = tk.NORMAL if summary else tk.DISABLED
        self.generate_btn.configure(state=state)
        self.print_btn.configure(state=state)
        self.email_btn.configure(state=state)
        self.whatsapp_btn.configure(state=state)

    # ---------- login helpers ----------
    def _prompt_login(self) -> bool:
        result = {"proceed": False}

        dialog = tk.Toplevel(self)


        # Make the Visit Note dialog tall enough by default
        try:
            dialog.title("Visit Note")
        except Exception:
            pass
        try:
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            width  = max(960, int(screen_w * 0.75))
            height = max(720, int(screen_h * 1.00))  # taller: ~80% of screen
            dialog.geometry(f"{width}x{height}")
            dialog.minsize(int(screen_w * 0.60), int(screen_h * 0.80))  # safe minimum
        except Exception:
            dialog.geometry("1100x780")
            dialog.minsize(1000, 720)

        # make main area stretch
        try:
            dialog.grid_rowconfigure(0, weight=1)
            dialog.grid_columnconfigure(0, weight=1)
        except Exception:
            pass

        # center on screen (after geometry applied)
        try:
            dialog.update_idletasks()
            sw, sh = dialog.winfo_screenwidth(), dialog.winfo_screenheight()
            ww, hh = dialog.winfo_width(), dialog.winfo_height()
            x = max((sw - ww) // 2, 0)
            y = max((sh - hh) // 2, 0)
            dialog.geometry(f"{ww}x{hh}+{x}+{y}")
        except Exception:
            pass

        dialog.title("Sign In")
        try:
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            width = max(460, int(screen_w * 0.35))
            height = max(260, int(screen_h * 0.3))
            dialog.geometry(f"{width}x{height}")
        except Exception:
            dialog.geometry("520x280")
        dialog.resizable(False, False)
        if self.winfo_viewable():
            dialog.transient(self)
        dialog.grab_set()
        dialog.update_idletasks()
        dialog.lift()
        try:
            dialog.attributes("-topmost", True)
            dialog.after(250, lambda: dialog.attributes("-topmost", False))
        except Exception:
            pass
        dialog.focus_force()

        frame = ttk.Frame(dialog, padding=24)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Clinic Receipt Login", font=("Segoe UI Semibold", 12)).pack(anchor="w")
        ttk.Label(frame, text="Enter your username to continue.").pack(anchor="w", pady=(2, 16))

        username_var = tk.StringVar()
        status_var = tk.StringVar()

        ttk.Label(frame, text="Username").pack(anchor="w")
        entry = ttk.Entry(frame, textvariable=username_var)
        entry.pack(fill="x", pady=(2, 8))
        entry.focus_set()

        status_label = ttk.Label(frame, textvariable=status_var, foreground="#D92D20")
        status_label.pack(anchor="w", pady=(0, 10))

        btn_bar = ttk.Frame(frame)
        btn_bar.pack(fill="x", pady=(8, 0))

        def close_dialog() -> None:
            dialog.grab_release()
            dialog.destroy()

        def skip_login() -> None:
            result["proceed"] = True
            self.session_user = None
            close_dialog()

        def attempt_login() -> None:
            name = username_var.get().strip()
            if not name:
                status_var.set("Please enter a username.")
                return
            db = self._ensure_login_database()
            if not db:
                status_var.set("Unable to reach the database.")
                return
            try:
                if db.username_exists(name):
                    self.session_user = name
                    result["proceed"] = True
                    close_dialog()
                else:
                    status_var.set("Username not found.")
            except Exception as exc:
                status_var.set("Login failed. See details in alert.")
                messagebox.showerror("Login", f"Could not validate username:\n{exc}")

        ttk.Button(btn_bar, text="Continue without login", style="Ghost.TButton", command=skip_login).pack(side="left")
        ttk.Button(btn_bar, text="Login", style="Primary.TButton", command=attempt_login).pack(side="right")
        ttk.Button(btn_bar, text="Exit", command=close_dialog).pack(side="right", padx=(0, 8))

        dialog.protocol("WM_DELETE_WINDOW", close_dialog)
        dialog.bind("<Return>", lambda _e: attempt_login())

        self.wait_window(dialog)
        return result["proceed"]

    def _ensure_login_database(self) -> ClinicDatabase | None:
        if self.db:
            return self.db
        try:
            settings = self.cfg.mysql_settings()
            self.db = ClinicDatabase(**settings)
            return self.db
        except Exception as exc:
            messagebox.showerror("Database", f"Failed to connect to MySQL: {exc}")
            self.db = None
            return None

    # ---------- Logo path resolution (fixed to include clinic_logo.jpg and work for exe/source) ----------
    def _default_logo_path(self) -> Path | None:
        """
        Look for a bundled default logo in common places, whether running from source or a frozen exe.
        Priority file (per your project): assets/clinic_logo.jpg
        """
        bases: list[Path] = []
        # 1) Bundle temp dir first (PyInstaller onefile)
        bases.append(_bundle_base_dir())
        # 2) Folder containing the executable when frozen
        try:
            if getattr(sys, "frozen", False):
                bases.append(Path(sys.executable).resolve().parent)
        except Exception:
            pass
        # 3) Module folder (running from source)
        bases.append(Path(__file__).resolve().parent)
        # 4) Current working directory as a last resort
        bases.append(Path.cwd())

        candidates_rel = (
            "assets/clinic_logo.jpg",   # << primary default per your setup
            "assets/logo.png",
            "assets/clinic.png",
            "assets/logo.jpg",
            "app/assets/logo.png",
            "app/assets/clinic.png",
            "app/assets/logo.jpg",
        )
        for base in bases:
            for rel in candidates_rel:
                p = (base / rel)
                if p.exists():
                    return p
        return None

    def _effective_logo_path(self) -> Path | None:
        """
        Resolve user-configured logo path robustly:
        - Absolute path: use directly if exists
        - Relative path: try relative to (in order)
            1) executable folder (when frozen)
            2) module folder
            3) PyInstaller bundle base (sys._MEIPASS)
            4) current working directory
        Fall back to bundled default logo if none found.
        """
        raw = ""
        try:
            raw = (self.cfg.settings.clinic.logo_path or "").strip()
        except Exception:
            raw = ""

        # 1) Direct absolute
        if raw:
            p = Path(raw)
            if p.is_absolute() and p.exists():
                return p

        # 2) Try relative against several bases
        bases: list[Path] = []
        try:
            if getattr(sys, "frozen", False):
                bases.append(Path(sys.executable).resolve().parent)
        except Exception:
            pass
        bases.extend([Path(__file__).resolve().parent, _bundle_base_dir(), Path.cwd()])

        if raw:
            for b in bases:
                candidate = (b / raw)
                if candidate.exists():
                    return candidate

        # 3) Default bundled logo
        return self._default_logo_path()

    # ---------- filename helpers (uses SELECTED payment when available) ----------
    def _slug(self, s: str) -> str:
        s = (s or "").strip().replace(" ", "_")
        s = re.sub(r"[^A-Za-z0-9_]+", "_", s)
        s = re.sub(r"_+", "_", s).strip("_")
        return s or "RECEIPT"

    def _payment_suffix(self, payments: list[PaymentEntry], summary: ReceiptSummary) -> str:
        total = float(summary.receipt.total or 0.0)
        total_paid = sum(float(p.amount or 0.0) for p in payments)
        # Omit suffix if settled in one payment
        if len(payments) == 1 and abs(total_paid - total) < 0.01:
            return ""
        # Prefer selected payment index; else fallback to latest
        if self.selected_payment and self.selected_payment.sequence >= 1:
            n = self.selected_payment.sequence
        else:
            n = len(payments)
        return f"_p{n}" if n >= 1 else ""

    def _desired_filename(self, summary: ReceiptSummary, payments: list[PaymentEntry]) -> str:
        patient_name = self._patient_display_name(summary.patient)
        ic = (summary.patient.icpassport or "").strip()
        date_part = summary.receipt.issued.strftime("%Y%m%d")
        suffix = self._payment_suffix(payments, summary)
        return f"receipt_{self._slug(patient_name)}_{self._slug(ic)}_{date_part}{suffix}.pdf"

    def _rename_to_desired(self, provisional_path: Path, summary: ReceiptSummary, payments: list[PaymentEntry]) -> Path:
        desired = provisional_path.with_name(self._desired_filename(summary, payments))
        if desired.exists() and desired.resolve() != provisional_path.resolve():
            base = desired.stem
            suffix = desired.suffix
            parent = desired.parent
            n = 2
            while True:
                candidate = parent / f"{base}_{n}{suffix}"
                if not candidate.exists():
                    desired = candidate
                    break
                n += 1
        if provisional_path.resolve() != desired.resolve():
            try:
                provisional_path.rename(desired)
                return desired
            except Exception:
                desired.write_bytes(provisional_path.read_bytes())
                try:
                    provisional_path.unlink(missing_ok=True)
                except TypeError:
                    try:
                        os.remove(provisional_path)
                    except Exception:
                        pass
                return desired
        return provisional_path

    def _ensure_pdf_generator(self) -> None:
        desired = (self.cfg.settings.receipt.output_directory or "receipts").strip()
        self.cfg.update_receipt(output_directory=desired or "receipts")
        out_dir = self.cfg.resolve_output_dir()
        self.pdf = ReceiptPDFGenerator(out_dir)

    def _generate_pdf(self) -> Path | None:
        summary = self._selected_summary()
        if not (self.db and self.pdf and summary):
            return None

        # Items
        items = self.db.get_receipt_items(summary.receipt.rcpt_id)

        # Payments from DB
        pay_map = self.db.partial_payments_for_receipts([summary.receipt.rcpt_id]) or {}
        pays_raw = pay_map.get(summary.receipt.rcpt_id, [])

        # Choose the "current" payment index: prefer selected row, else latest
        if pays_raw:
            if self.selected_payment and 1 <= self.selected_payment.sequence <= len(pays_raw):
                current_idx = self.selected_payment.sequence
            else:
                current_idx = len(pays_raw)

            previous_entries = []
            total_due = float(summary.receipt.total or 0.0)
            total_paid_running = 0.0
            current_rec = None
            current_amt = 0.0

            for idx, rec in enumerate(pays_raw, start=1):
                amt = float(rec.amount or 0.0)
                if idx < current_idx:
                    previous_entries.append((idx, rec.date, amt))
                elif idx == current_idx:
                    current_rec = rec
                    current_amt = amt
                total_paid_running += amt

            total_paid_all = sum(float(r.amount or 0.0) for r in pays_raw)
            balance = max(total_due - total_paid_all, 0.0)

            # Always display a friendly LABEL in the PDF
            method_raw = (current_rec.method or current_rec.pay_code or "").strip() if current_rec else ""
            self._ensure_payment_methods_loaded()
            # If method_raw is a code, map to label; if it's already a label, keep it
            # _payment_label_for_code returns label for codes, and falls back to the input if unknown
            method_label_mapped = self._payment_label_for_code(method_raw) if method_raw else ""
            method_label = method_label_mapped or method_raw

            # Current payment entry label
            if len(pays_raw) == 1 and abs(total_paid_all - total_due) < 0.01:
                entry_label = method_label or "Payment"
            else:
                entry_label = f"Payment #{current_idx}" + (f" ({method_label})" if method_label else "")


            payment_entries: list[PaymentEntry] = [PaymentEntry(method=entry_label, amount=current_amt)]

            progress = PaymentProgress(
                sequence=current_idx,
                current_amount=current_amt,
                total_paid=total_paid_all,
                balance=balance,
                total_due=total_due,
                received_on=current_rec.date if current_rec else summary.receipt.issued,
                method=method_label,
                previous_payments=tuple(previous_entries),
                remark=(summary.receipt.remark or ""),
            )
        else:
            # No partials recorded: treat as single full payment on receipt date
            total_due = float(summary.receipt.total or 0.0)

            # Map header code to a friendly label for display consistency
            self._ensure_payment_methods_loaded()
            code = (summary.receipt.payment_code or "").strip()
            label = self._payment_label_for_code(code) if code else ""
            display_label = label or code or "Payment"

            payment_entries = [PaymentEntry(method=display_label, amount=total_due)]
            progress = PaymentProgress(
                sequence=1,
                current_amount=total_due,
                total_paid=total_due,
                balance=0.0,
                total_due=total_due,
                received_on=summary.receipt.issued,
                method=display_label,
                previous_payments=tuple(),
                remark=(summary.receipt.remark or ""),
            )


        logo_path = self._effective_logo_path()

        # Generate the PDF with Payment Progress
        provisional = self.pdf.generate(
            clinic=self.cfg.settings.clinic,
            patient=summary.patient,
            receipt=summary.receipt,
            items=items,
            payments=payment_entries,
            logo_path=str(logo_path) if logo_path else None,
            payment_progress=progress,
        )

        # Rename to your exact filename pattern
        final_path = self._rename_to_desired(Path(provisional), summary, payment_entries)
        self.status_var.set(f"PDF saved: {final_path.name}")
        return final_path

    def _print_pdf(self) -> None:
        path = self._generate_pdf()
        if not path:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            webbrowser.open(Path(path).as_uri())

    def _whatsapp(self) -> None:
        summary = self._selected_summary()
        if not summary:
            messagebox.showinfo("WhatsApp", "Please select a receipt first.")
            return

        prefill = ""
        try:
            patient_phone = getattr(summary.patient, "phone", "") or getattr(summary.patient, "mobile", "") or ""
            prefill = normalize_msisdn_malaysia(patient_phone)
        except Exception:
            prefill = ""

        number = simpledialog.askstring("WhatsApp Number",
                                        "Enter recipient phone number (digits only, with country code).\n"
                                        "If it starts with 0, we will add '6' in front (Malaysia).",
                                        initialvalue=prefill)
        if not number:
            return

        number = normalize_msisdn_malaysia(number)
        if not number:
            messagebox.showwarning("WhatsApp", "Invalid phone number.")
            return

        # Ensure a fresh PDF exists (honoring selected payment for filename and progress)
        self._generate_pdf()

        name = self._patient_display_name(summary.patient)
        clinic_name = self.cfg.settings.clinic.name or "clinic"
        total_label = fmt_money(summary.receipt.total or 0.0)
        date_label = summary.receipt.issued.strftime("%d-%m-%Y")
        rcpt_id = summary.receipt.rcpt_id

        text = (
            f"Hello {name}, this is {clinic_name}. "
            f"Official receipt {rcpt_id} dated {date_label}. "
            f"Total: {total_label}. "
            f"Thank you."
        )
        url = f"https://api.whatsapp.com/send?phone={number}&text={urlquote(text)}"
        webbrowser.open(url)

    def _email(self) -> None:
        summary = self._selected_summary()
        if not summary:
            return

        to_email = simpledialog.askstring("Email To", "Enter recipient email address:", initialvalue="")
        if not to_email:
            return

        attachment = self._generate_pdf()
        if not attachment:
            return

        sender = self.cfg.settings.email.sender.strip()
        app_pw = self.cfg.settings.email.app_password.strip()
        subject_t = self.cfg.settings.email.subject or "Clinic Receipt"
        body_t = self.cfg.settings.email.body or ""

        try:
            subject = subject_t.format(patient_name=self._patient_display_name(summary.patient),
                                       receipt_id=summary.receipt.rcpt_id)
            body = body_t.format(patient_name=self._patient_display_name(summary.patient),
                                 receipt_id=summary.receipt.rcpt_id)
        except Exception as exc:
            messagebox.showerror("Email", f"Template error: {exc}")
            return

        try:
            msg = EmailMessage()
            msg["From"] = sender
            msg["To"] = to_email
            msg["Subject"] = subject
            msg.set_content(body)
            msg.add_attachment(attachment.read_bytes(), maintype="application",
                               subtype="pdf", filename=attachment.name)

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sender, app_pw)
                smtp.send_message(msg)

            messagebox.showinfo("Email", f"Sent to {to_email}")
        except Exception as exc:
            messagebox.showerror("Email", f"Failed to send email: {exc}")

    # ---------------- settings helpers
    def _on_source_changed(self) -> None:
        src = {"Live MySQL connection": "mysql", "Backup file": "backup"}.get(self.data_source_var.get(), "mysql")
        self.cfg.update_database(source=src)

    def _choose_logo(self) -> None:
        path = filedialog.askopenfilename(title="Select logo", filetypes=[("Image", "*.png;*.jpg;*.jpeg;*.ico")])
        if path:
            self.logo_var.set(path)

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_dir_var.set(path)
            self._ensure_pdf_generator()

    def _choose_backup(self) -> None:
        path = filedialog.askopenfilename(title="Select backup SQL", filetypes=[("SQL", "*.sql")])
        if path:
            self.backup_path_var.set(path)

    def _save_settings(self) -> None:
        self.cfg.update_clinic(
            name=self.clinic_name.get().strip(),
            phone=self.clinic_phone.get().strip(),
            email=self.clinic_email.get().strip(),
            address=self.clinic_address.get("1.0", "end").strip(),
            logo_path=self.logo_var.get().strip(),
        )
        self.cfg.update_database(
            source={"Live MySQL connection": "mysql", "Backup file": "backup"}.get(self.data_source_var.get(), "mysql"),
            backup_path=self.backup_path_var.get().strip(),
            mysql_host=self.mysql_host.get().strip(),
            mysql_port=int(self.mysql_port.get() or 3306),
            mysql_user=self.mysql_user.get().strip(),
            mysql_password=self.mysql_pass.get().strip(),
            mysql_database=self.mysql_db.get().strip(),
        )
        self.cfg.update_email(
            sender=self.email_sender.get().strip(),
            app_password=self.email_app_pw.get().strip(),
            subject=self.email_subject.get().strip(),
        )
        self.cfg.settings.email.body = self.email_body.get("1.0", "end")
        self.cfg.save()
        self._ensure_pdf_generator()
        messagebox.showinfo("Settings", "Saved.")

    def _reconnect(self) -> None:
        if self.db:
            try:
                self.db.close()
            except Exception:
                pass
        try:
            ms = self.cfg.mysql_settings()
            self.db = ClinicDatabase(**ms)
            messagebox.showinfo("Database", "Successfully connected to MySQL database.")
        except Exception as exc:
            messagebox.showerror("Database", f"Failed to connect: {exc}")
            self.db = None

    # ---------------- misc helpers
    def _selected_summary(self) -> ReceiptSummary | None:
        sel = self.receipt_tree.selection()
        if not sel:
            return None
        return self.receipt_index.get(sel[0])

    def _patient_display_name(self, patient: Patient) -> str:
        for candidate in (patient.receipt_name, patient.preferred_name, patient.name):
            if candidate and candidate.strip():
                return candidate.strip()
        return "Patient"

def _row(s: ttk.Frame, label: str, widget: tk.Widget, *, row: int, col: int, trailing: tk.Widget | None = None):
    ttk.Label(s, text=label).grid(row=row, column=col, sticky="w", pady=4)
    widget.grid(row=row, column=col + 1, sticky="we", padx=8, pady=4)
    if trailing:
        trailing.grid(row=row, column=col + 2, padx=(0, 0))
    s.grid_columnconfigure(col + 1, weight=1)

def run_app(config_path: Path) -> None:
    app = ReceiptApp(config_path)
    try:
        if getattr(app, "_ready", False):
            app.mainloop()
    finally:
        db = getattr(app, "db", None)
        if db:
            try:
                db.close()
            except Exception:
                pass






