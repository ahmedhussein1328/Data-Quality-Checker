"""Desktop GUI for the Data Quality Checker application."""

from __future__ import annotations

import ctypes
import errno
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# Allow running directly as `python src/gui.py` as well as `python main.py`
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from src.csv_reader import read_file
from src.data_validator import EXPECTED_COLUMNS, validate_data
from src.duplicate_checker import find_duplicates
from src.missing_checker import find_missing_values
from src.report_generator import generate_report


COLORS = {
    "navy": "#142136",
    "navy_hover": "#1E314F",
    "navy_active": "#243B5E",
    "navy_divider": "#2C4263",
    "navy_muted": "#94A3B8",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "accent_soft": "#EFF6FF",
    "bg": "#F3F5F9",
    "white": "#FFFFFF",
    "surface_alt": "#F8FAFC",
    "border": "#E2E8F0",
    "text": "#1E293B",
    "muted": "#64748B",
    "green": "#15803D",
    "green_bg": "#DCFCE7",
    "red": "#DC2626",
    "red_bg": "#FEE2E2",
    "orange": "#D97706",
    "orange_bg": "#FEF3C7",
    "neutral_bg": "#E2E8F0",
}

NAV_ITEMS = [
    ("dashboard", "Dashboard", "Overview of dataset metrics and quick actions"),
    ("preview", "Data Preview", "Inspect loaded CSV or Excel records in a structured table"),
    ("results", "Validation Results", "Review duplicate, missing-value, and data validation findings"),
    ("reports", "Reports", "Export validation findings to a multi-sheet Excel workbook"),
    ("info", "Application Info", "Supported file types, validation rules, and report structure"),
]


def is_missing_value(value) -> bool:
    """Return True if a scalar cell value represents a missing/null entry."""
    try:
        result = pd.isna(value)
        return bool(result) if not hasattr(result, "__len__") else False
    except (TypeError, ValueError):
        return False


def format_cell_value(value, empty_placeholder: str = "") -> str:
    """Format a DataFrame cell safely for display in a Treeview."""
    if is_missing_value(value):
        return empty_placeholder
    return str(value)


def format_row_context(row: pd.Series, exclude_col: str | None = None) -> str:
    """Build a concise key=value summary of a row for issue tables."""
    parts = []
    for col, val in row.items():
        if col == exclude_col:
            continue
        display = "<empty>" if is_missing_value(val) else str(val)
        parts.append(f"{col}: {display}")
    return " | ".join(parts)


def configure_ttk_styles(root: tk.Misc) -> None:
    """Configure consistent ttk styling for tables, tabs, and inputs."""
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure(
        "Treeview",
        background=COLORS["white"],
        foreground=COLORS["text"],
        fieldbackground=COLORS["white"],
        rowheight=30,
        borderwidth=0,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Treeview.Heading",
        background="#EEF2F7",
        foreground=COLORS["text"],
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padding=(10, 8),
    )
    style.map(
        "Treeview.Heading",
        background=[("active", "#E2E8F0")],
    )
    style.map(
        "Treeview",
        background=[("selected", "#DBEAFE")],
        foreground=[("selected", COLORS["text"])],
    )

    style.configure("TNotebook", background=COLORS["white"], borderwidth=0)
    style.configure(
        "TNotebook.Tab",
        background="#EDF2F7",
        foreground=COLORS["muted"],
        padding=(16, 9),
        font=("Segoe UI", 9, "bold"),
        borderwidth=0,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", COLORS["white"]), ("active", "#E2E8F0")],
        foreground=[("selected", COLORS["accent"]), ("active", COLORS["text"])],
    )

    style.configure(
        "TCombobox",
        fieldbackground=COLORS["white"],
        background=COLORS["white"],
        foreground=COLORS["text"],
        padding=(6, 4),
    )


class ScrollableTable(tk.Frame):
    """Reusable ttk.Treeview with vertical/horizontal scrollbars and empty state."""

    def __init__(self, parent: tk.Misc, empty_text: str = "No records to display."):
        super().__init__(parent, bg=COLORS["white"])
        self._empty_text = empty_text

        self.table_container = tk.Frame(self, bg=COLORS["white"])
        self.table_container.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(self.table_container, show="headings", selectmode="browse")
        self.v_scroll = ttk.Scrollbar(self.table_container, orient="vertical", command=self.tree.yview)
        self.h_scroll = ttk.Scrollbar(self.table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")
        self.table_container.rowconfigure(0, weight=1)
        self.table_container.columnconfigure(0, weight=1)

        self.tree.tag_configure("even", background=COLORS["white"])
        self.tree.tag_configure("odd", background=COLORS["surface_alt"])

        self.empty_frame = tk.Frame(
            self,
            bg=COLORS["surface_alt"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )
        self.empty_icon = tk.Label(
            self.empty_frame,
            text="✓",
            bg=COLORS["surface_alt"],
            fg=COLORS["muted"],
            font=("Segoe UI", 22, "bold"),
        )
        self.empty_icon.pack(pady=(28, 6))
        self.empty_label = tk.Label(
            self.empty_frame,
            text=empty_text,
            bg=COLORS["surface_alt"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
            justify="center",
        )
        self.empty_label.pack(pady=(0, 28), padx=24)

    def clear(self) -> None:
        """Remove all rows and columns from the table."""
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()

    def show_empty(self, message: str | None = None, icon: str = "—", icon_color: str | None = None) -> None:
        """Display the centered empty-state panel instead of the table."""
        self.clear()
        self.table_container.pack_forget()
        self.empty_label.configure(text=message or self._empty_text)
        self.empty_icon.configure(text=icon, fg=icon_color or COLORS["muted"])
        self.empty_frame.pack(fill="both", expand=True)

    def populate(
        self,
        columns: list[str],
        rows: list[tuple],
        column_widths: dict[str, int] | None = None,
        stretch_last: bool = True,
    ) -> None:
        """Populate the Treeview with columns and row tuples."""
        self.empty_frame.pack_forget()
        self.table_container.pack(fill="both", expand=True)

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns

        widths = column_widths or {}
        for idx, col in enumerate(columns):
            self.tree.heading(col, text=col, anchor="w")
            default_w = max(110, min(240, len(col) * 11 + 32))
            width = widths.get(col, default_w)
            stretch = stretch_last and (idx == len(columns) - 1)
            self.tree.column(col, width=width, minwidth=80, anchor="w", stretch=stretch)

        for idx, values in enumerate(rows):
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert("", "end", values=values, tags=(tag,))


class DataQualityApp:
    """Main desktop GUI controller for Data Quality Checker."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Data Quality Checker")
        self.root.geometry("1220x780")
        self.root.minsize(980, 660)
        self.root.configure(bg=COLORS["bg"])

        self.df: pd.DataFrame | None = None
        self.file_path: Path | None = None
        self.results: dict | None = None
        self.last_report_path: Path | None = None
        self.active_section: str = "dashboard"
        self.preview_limit_var = tk.StringVar(value="500")

        self.nav_buttons: dict[str, tk.Frame] = {}
        self.nav_indicators: dict[str, tk.Frame] = {}
        self.nav_labels: dict[str, tk.Label] = {}
        self.sections: dict[str, tk.Frame] = {}
        self.stat_cards: dict[str, dict[str, tk.Widget]] = {}

        configure_ttk_styles(self.root)
        self._build_shell()
        self._build_dashboard_section()
        self._build_preview_section()
        self._build_results_section()
        self._build_reports_section()
        self._build_info_section()

        self.show_section("dashboard")
        self._refresh_all_views()

    # ------------------------------------------------------------------
    # Shell Layout (Sidebar + Header + Content Container)
    # ------------------------------------------------------------------

    def _build_shell(self) -> None:
        self.sidebar = tk.Frame(self.root, bg=COLORS["navy"], width=235)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        right_area = tk.Frame(self.root, bg=COLORS["bg"])
        right_area.pack(side="right", fill="both", expand=True)

        self._build_sidebar()
        self._build_header(right_area)

        self.content_host = tk.Frame(right_area, bg=COLORS["bg"])
        self.content_host.pack(fill="both", expand=True, padx=24, pady=20)
        self.content_host.rowconfigure(0, weight=1)
        self.content_host.columnconfigure(0, weight=1)

    def _build_sidebar(self) -> None:
        brand = tk.Frame(self.sidebar, bg=COLORS["navy"])
        brand.pack(fill="x", padx=20, pady=(24, 22))

        logo_row = tk.Frame(brand, bg=COLORS["navy"])
        logo_row.pack(anchor="w")

        tk.Label(
            logo_row,
            text="DQ",
            bg=COLORS["accent"],
            fg=COLORS["white"],
            font=("Segoe UI", 12, "bold"),
            padx=9,
            pady=5,
        ).pack(side="left")

        title_box = tk.Frame(logo_row, bg=COLORS["navy"])
        title_box.pack(side="left", padx=(10, 0))
        tk.Label(
            title_box,
            text="Data Quality",
            bg=COLORS["navy"],
            fg=COLORS["white"],
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w")
        tk.Label(
            title_box,
            text="Checker",
            bg=COLORS["navy"],
            fg="#93C5FD",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        tk.Frame(self.sidebar, bg=COLORS["navy_divider"], height=1).pack(fill="x", padx=18, pady=(0, 16))

        tk.Label(
            self.sidebar,
            text="NAVIGATION",
            bg=COLORS["navy"],
            fg=COLORS["navy_muted"],
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", padx=20, pady=(0, 8))

        icons = {
            "dashboard": "▦",
            "preview": "▤",
            "results": "✓",
            "reports": "⇩",
            "info": "ⓘ",
        }

        for key, label, _ in NAV_ITEMS:
            row = tk.Frame(self.sidebar, bg=COLORS["navy"], cursor="hand2")
            row.pack(fill="x", padx=12, pady=2)

            indicator = tk.Frame(row, bg=COLORS["navy"], width=4)
            indicator.pack(side="left", fill="y")

            btn_label = tk.Label(
                row,
                text=f"  {icons[key]}   {label}",
                bg=COLORS["navy"],
                fg="#CBD5E1",
                anchor="w",
                padx=10,
                pady=10,
                font=("Segoe UI", 10),
                cursor="hand2",
            )
            btn_label.pack(side="left", fill="x", expand=True)

            for widget in (row, indicator, btn_label):
                widget.bind("<Button-1>", lambda _e, k=key: self.show_section(k))
                widget.bind("<Enter>", lambda _e, k=key: self._on_nav_hover(k, True))
                widget.bind("<Leave>", lambda _e, k=key: self._on_nav_hover(k, False))

            self.nav_buttons[key] = row
            self.nav_indicators[key] = indicator
            self.nav_labels[key] = btn_label

        bottom_box = tk.Frame(self.sidebar, bg="#1B2B45")
        bottom_box.pack(side="bottom", fill="x", padx=14, pady=16)

        tk.Label(
            bottom_box,
            text="DATASET STATUS",
            bg="#1B2B45",
            fg=COLORS["navy_muted"],
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))

        self.sidebar_status_label = tk.Label(
            bottom_box,
            text="●  No file loaded",
            bg="#1B2B45",
            fg="#E2E8F0",
            font=("Segoe UI", 9),
            wraplength=180,
            justify="left",
        )
        self.sidebar_status_label.pack(anchor="w", padx=12, pady=(0, 10))

    def _build_header(self, parent: tk.Misc) -> None:
        header = tk.Frame(parent, bg=COLORS["white"], highlightbackground=COLORS["border"], highlightthickness=1)
        header.pack(fill="x")

        inner = tk.Frame(header, bg=COLORS["white"])
        inner.pack(fill="x", padx=24, pady=14)

        left = tk.Frame(inner, bg=COLORS["white"])
        left.pack(side="left", fill="x", expand=True)

        self.header_title = tk.Label(
            left,
            text="Dashboard",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 17, "bold"),
        )
        self.header_title.pack(anchor="w")

        self.header_subtitle = tk.Label(
            left,
            text="",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
        )
        self.header_subtitle.pack(anchor="w", pady=(2, 0))

        self.status_badge = tk.Label(
            inner,
            text="READY",
            bg=COLORS["neutral_bg"],
            fg=COLORS["text"],
            font=("Segoe UI", 8, "bold"),
            padx=12,
            pady=6,
        )
        self.status_badge.pack(side="right")

    def _on_nav_hover(self, key: str, entering: bool) -> None:
        if key == self.active_section:
            return
        bg = COLORS["navy_hover"] if entering else COLORS["navy"]
        fg = COLORS["white"] if entering else "#CBD5E1"
        self.nav_buttons[key].configure(bg=bg)
        self.nav_indicators[key].configure(bg=bg)
        self.nav_labels[key].configure(bg=bg, fg=fg)

    def show_section(self, key: str) -> None:
        """Switch the main content area to the selected section."""
        self.active_section = key
        for nav_key, _, subtitle in NAV_ITEMS:
            is_active = nav_key == key
            bg = COLORS["navy_active"] if is_active else COLORS["navy"]
            fg = COLORS["white"] if is_active else "#CBD5E1"
            ind_bg = "#60A5FA" if is_active else COLORS["navy"]
            font = ("Segoe UI", 10, "bold") if is_active else ("Segoe UI", 10)

            self.nav_buttons[nav_key].configure(bg=bg)
            self.nav_indicators[nav_key].configure(bg=ind_bg)
            self.nav_labels[nav_key].configure(bg=bg, fg=fg, font=font)

            if is_active:
                title = next(label for k, label, _ in NAV_ITEMS if k == key)
                self.header_title.configure(text=title)
                self.header_subtitle.configure(text=subtitle)

        self.sections[key].tkraise()

    # ------------------------------------------------------------------
    # Reusable UI Builders
    # ------------------------------------------------------------------

    @staticmethod
    def _make_card(parent: tk.Misc) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=COLORS["white"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
        )

    @staticmethod
    def _make_button(parent: tk.Misc, text: str, command, style_kind: str = "primary") -> tk.Button:
        palette = {
            "primary": (COLORS["accent"], COLORS["white"], COLORS["accent_hover"]),
            "secondary": ("#EEF2F7", COLORS["text"], "#E2E8F0"),
            "success": (COLORS["green"], COLORS["white"], "#166534"),
        }
        bg, fg, active_bg = palette[style_kind]
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            disabledforeground="#94A3B8",
            relief="flat",
            bd=0,
            padx=16,
            pady=9,
            cursor="hand2",
            font=("Segoe UI", 9, "bold"),
        )

    def _create_stat_card(self, parent: tk.Misc, key: str, title: str, subtitle: str) -> tk.Frame:
        card = self._make_card(parent)

        accent_bar = tk.Frame(card, bg=COLORS["border"], height=4)
        accent_bar.pack(fill="x")

        body = tk.Frame(card, bg=COLORS["white"])
        body.pack(fill="both", expand=True, padx=18, pady=14)

        tk.Label(
            body,
            text=title,
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w")

        value_lbl = tk.Label(
            body,
            text="—",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 24, "bold"),
        )
        value_lbl.pack(anchor="w", pady=(6, 4))

        caption_lbl = tk.Label(
            body,
            text=subtitle,
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
        )
        caption_lbl.pack(anchor="w")

        self.stat_cards[key] = {
            "accent": accent_bar,
            "value": value_lbl,
            "caption": caption_lbl,
            "default_caption": subtitle,
        }
        return card

    # ------------------------------------------------------------------
    # Section 1: Dashboard
    # ------------------------------------------------------------------

    def _build_dashboard_section(self) -> None:
        frame = tk.Frame(self.content_host, bg=COLORS["bg"])
        frame.grid(row=0, column=0, sticky="nsew")
        self.sections["dashboard"] = frame

        # Dataset selection panel
        dataset_card = self._make_card(frame)
        dataset_card.pack(fill="x", pady=(0, 16))

        ds_inner = tk.Frame(dataset_card, bg=COLORS["white"])
        ds_inner.pack(fill="x", padx=20, pady=16)

        ds_info = tk.Frame(ds_inner, bg=COLORS["white"])
        ds_info.pack(side="left", fill="x", expand=True)

        tk.Label(
            ds_info,
            text="Dataset Selection",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        self.dash_filename_label = tk.Label(
            ds_info,
            text="No dataset file selected",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10, "bold"),
        )
        self.dash_filename_label.pack(anchor="w", pady=(6, 2))

        self.dash_status_label = tk.Label(
            ds_info,
            text="Click 'Browse File' to load a CSV or Excel (.xlsx, .xls) dataset.",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
        )
        self.dash_status_label.pack(anchor="w")

        btn_group = tk.Frame(ds_inner, bg=COLORS["white"])
        btn_group.pack(side="right", padx=(16, 0))

        self.browse_btn = self._make_button(btn_group, "Browse File", self.browse_file, style_kind="secondary")
        self.browse_btn.pack(side="left", padx=(0, 10))

        self.run_check_btn = self._make_button(btn_group, "Run Quality Check", self.run_quality_check, style_kind="primary")
        self.run_check_btn.pack(side="left")

        # Row 1: Dataset dimension cards (2 equal cards)
        row1 = tk.Frame(frame, bg=COLORS["bg"])
        row1.pack(fill="x", pady=(0, 12))
        row1.columnconfigure(0, weight=1, uniform="row1_cards")
        row1.columnconfigure(1, weight=1, uniform="row1_cards")

        self._create_stat_card(row1, "total_rows", "Total Rows", "Records loaded from file").grid(
            row=0, column=0, sticky="nsew", padx=(0, 12)
        )
        self._create_stat_card(row1, "total_columns", "Total Columns", "Detected dataset fields").grid(
            row=0, column=1, sticky="nsew"
        )

        # Row 2: Quality check cards (3 equal cards)
        row2 = tk.Frame(frame, bg=COLORS["bg"])
        row2.pack(fill="x", pady=(0, 16))
        for col in range(3):
            row2.columnconfigure(col, weight=1, uniform="row2_cards")

        self._create_stat_card(row2, "duplicates", "Duplicate Records", "Repeated rows found").grid(
            row=0, column=0, sticky="nsew", padx=(0, 12)
        )
        self._create_stat_card(row2, "missing", "Missing Values", "Empty or null cells").grid(
            row=0, column=1, sticky="nsew", padx=(0, 12)
        )
        self._create_stat_card(
            row2,
            "validation_issues",
            "Data Validation Issues",
            "Type and format errors detected",
        ).grid(row=0, column=2, sticky="nsew")

        # Bottom summary / workflow card
        overview_card = self._make_card(frame)
        overview_card.pack(fill="both", expand=True)

        ov_inner = tk.Frame(overview_card, bg=COLORS["white"])
        ov_inner.pack(fill="both", expand=True, padx=20, pady=18)

        tk.Label(
            ov_inner,
            text="Quality Check Summary",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w")

        self.dash_overview_text = tk.Label(
            ov_inner,
            text="",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            justify="left",
            wraplength=850,
        )
        self.dash_overview_text.pack(anchor="w", pady=(10, 14))

        self.dash_quick_actions = tk.Frame(ov_inner, bg=COLORS["white"])
        self.dash_quick_actions.pack(anchor="w")

        self.quick_preview_btn = self._make_button(
            self.dash_quick_actions,
            "Open Data Preview",
            lambda: self.show_section("preview"),
            style_kind="secondary",
        )
        self.quick_preview_btn.pack(side="left", padx=(0, 10))

        self.quick_results_btn = self._make_button(
            self.dash_quick_actions,
            "View Validation Results",
            lambda: self.show_section("results"),
            style_kind="secondary",
        )
        self.quick_results_btn.pack(side="left", padx=(0, 10))

        self.quick_reports_btn = self._make_button(
            self.dash_quick_actions,
            "Go to Reports",
            lambda: self.show_section("reports"),
            style_kind="primary",
        )
        self.quick_reports_btn.pack(side="left")

    # ------------------------------------------------------------------
    # Section 2: Data Preview
    # ------------------------------------------------------------------

    def _build_preview_section(self) -> None:
        frame = tk.Frame(self.content_host, bg=COLORS["bg"])
        frame.grid(row=0, column=0, sticky="nsew")
        self.sections["preview"] = frame

        card = self._make_card(frame)
        card.pack(fill="both", expand=True)

        toolbar = tk.Frame(card, bg=COLORS["white"])
        toolbar.pack(fill="x", padx=18, pady=(16, 12))

        left = tk.Frame(toolbar, bg=COLORS["white"])
        left.pack(side="left", fill="x", expand=True)

        self.preview_title_label = tk.Label(
            left,
            text="No Dataset Loaded",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"),
        )
        self.preview_title_label.pack(anchor="w")

        self.preview_info_label = tk.Label(
            left,
            text="Select a CSV or Excel file to preview its columns and rows.",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
        )
        self.preview_info_label.pack(anchor="w", pady=(2, 0))

        controls = tk.Frame(toolbar, bg=COLORS["white"])
        controls.pack(side="right")

        tk.Label(
            controls,
            text="Max preview rows:",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
        ).pack(side="left", padx=(0, 8))

        self.preview_limit_combo = ttk.Combobox(
            controls,
            textvariable=self.preview_limit_var,
            values=("100", "250", "500", "1000", "2500"),
            state="readonly",
            width=7,
        )
        self.preview_limit_combo.pack(side="left", padx=(0, 10))
        self.preview_limit_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_preview())

        self._make_button(controls, "Browse File", self.browse_file, style_kind="secondary").pack(side="left")

        tk.Frame(card, bg=COLORS["border"], height=1).pack(fill="x")

        table_wrap = tk.Frame(card, bg=COLORS["white"])
        table_wrap.pack(fill="both", expand=True, padx=16, pady=16)

        self.preview_table = ScrollableTable(
            table_wrap,
            empty_text="No dataset loaded yet.\nUse 'Browse File' to open a CSV or Excel file.",
        )
        self.preview_table.pack(fill="both", expand=True)

    # ------------------------------------------------------------------
    # Section 3: Validation Results
    # ------------------------------------------------------------------

    def _build_results_section(self) -> None:
        frame = tk.Frame(self.content_host, bg=COLORS["bg"])
        frame.grid(row=0, column=0, sticky="nsew")
        self.sections["results"] = frame

        banner_card = self._make_card(frame)
        banner_card.pack(fill="x", pady=(0, 14))

        banner_inner = tk.Frame(banner_card, bg=COLORS["white"])
        banner_inner.pack(fill="x", padx=18, pady=14)

        banner_text = tk.Frame(banner_inner, bg=COLORS["white"])
        banner_text.pack(side="left", fill="x", expand=True)

        self.results_banner_title = tk.Label(
            banner_text,
            text="Validation Not Run Yet",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"),
        )
        self.results_banner_title.pack(anchor="w")

        self.results_banner_subtitle = tk.Label(
            banner_text,
            text="Load a dataset file and click 'Run Quality Check' to analyze data quality.",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            wraplength=760,
            justify="left",
        )
        self.results_banner_subtitle.pack(anchor="w", pady=(3, 0))

        self.results_run_btn = self._make_button(
            banner_inner,
            "Run Quality Check",
            self.run_quality_check,
            style_kind="primary",
        )
        self.results_run_btn.pack(side="right", padx=(12, 0))

        notebook_card = self._make_card(frame)
        notebook_card.pack(fill="both", expand=True)

        self.results_notebook = ttk.Notebook(notebook_card)
        self.results_notebook.pack(fill="both", expand=True, padx=12, pady=12)

        self.category_tabs: dict[str, dict] = {}
        categories = [
            ("duplicates", "Duplicate Detection"),
            ("missing", "Missing Values"),
            ("validation", "Data Validation"),
        ]

        for key, title in categories:
            tab = tk.Frame(self.results_notebook, bg=COLORS["white"])
            self.results_notebook.add(tab, text=f"  {title}  ")

            status_bar = tk.Frame(tab, bg=COLORS["surface_alt"], highlightbackground=COLORS["border"], highlightthickness=1)
            status_bar.pack(fill="x", padx=8, pady=(12, 10))

            badge = tk.Label(
                status_bar,
                text="PENDING",
                bg=COLORS["neutral_bg"],
                fg=COLORS["text"],
                font=("Segoe UI", 8, "bold"),
                padx=10,
                pady=4,
            )
            badge.pack(side="left", padx=(12, 10), pady=10)

            summary_lbl = tk.Label(
                status_bar,
                text="Run a quality check to view results for this category.",
                bg=COLORS["surface_alt"],
                fg=COLORS["text"],
                font=("Segoe UI", 9),
                justify="left",
                wraplength=760,
            )
            summary_lbl.pack(side="left", fill="x", expand=True, pady=10, padx=(0, 12))

            extra_note_lbl = tk.Label(
                tab,
                text="",
                bg=COLORS["white"],
                fg=COLORS["muted"],
                font=("Segoe UI", 9),
                justify="left",
                wraplength=850,
            )
            extra_note_lbl.pack(anchor="w", padx=10, pady=(0, 6))

            table = ScrollableTable(tab, empty_text="Run a quality check to inspect records.")
            table.pack(fill="both", expand=True, padx=8, pady=(0, 8))

            self.category_tabs[key] = {
                "frame": tab,
                "title": title,
                "badge": badge,
                "summary": summary_lbl,
                "note": extra_note_lbl,
                "table": table,
            }

    # ------------------------------------------------------------------
    # Section 4: Reports
    # ------------------------------------------------------------------

    def _build_reports_section(self) -> None:
        frame = tk.Frame(self.content_host, bg=COLORS["bg"])
        frame.grid(row=0, column=0, sticky="nsew")
        self.sections["reports"] = frame

        export_card = self._make_card(frame)
        export_card.pack(fill="x", pady=(0, 16))

        ex_inner = tk.Frame(export_card, bg=COLORS["white"])
        ex_inner.pack(fill="x", padx=20, pady=18)

        left = tk.Frame(ex_inner, bg=COLORS["white"])
        left.pack(side="left", fill="x", expand=True)

        tk.Label(
            left,
            text="Export Excel Report",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w")

        self.reports_status_label = tk.Label(
            left,
            text="Load a dataset file and run the quality check before exporting.",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            wraplength=650,
            justify="left",
        )
        self.reports_status_label.pack(anchor="w", pady=(6, 0))

        self.export_btn = self._make_button(ex_inner, "Export Excel Report", self.export_report, style_kind="primary")
        self.export_btn.pack(side="right", padx=(16, 0))

        details_card = self._make_card(frame)
        details_card.pack(fill="both", expand=True)

        det_inner = tk.Frame(details_card, bg=COLORS["white"])
        det_inner.pack(fill="both", expand=True, padx=20, pady=18)

        tk.Label(
            det_inner,
            text="Generated Workbook Contents",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        sheets = [
            ("Summary", "High-level counts of rows, columns, duplicate rows, missing values, and validation issues."),
            ("Original Data", "Complete copy of the loaded dataset without modifications."),
            ("Duplicates", "All detected duplicate records."),
            ("Missing Counts", "Per-column count of missing values."),
            ("Missing Rows", "Rows containing one or more missing values."),
            ("Data Validation", "Unified table of row numbers, columns, invalid values, rules, and issue details."),
        ]

        for sheet_name, desc in sheets:
            row = tk.Frame(det_inner, bg=COLORS["white"])
            row.pack(fill="x", pady=3)
            tk.Label(
                row,
                text=sheet_name,
                bg="#EFF6FF",
                fg=COLORS["accent"],
                font=("Segoe UI", 8, "bold"),
                width=16,
                anchor="w",
                padx=8,
                pady=3,
            ).pack(side="left", padx=(0, 10))
            tk.Label(
                row,
                text=desc,
                bg=COLORS["white"],
                fg=COLORS["muted"],
                font=("Segoe UI", 9),
                anchor="w",
            ).pack(side="left", fill="x", expand=True)

        tk.Frame(det_inner, bg=COLORS["border"], height=1).pack(fill="x", pady=(16, 14))

        self.last_export_label = tk.Label(
            det_inner,
            text="No report has been exported in this session yet.",
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            wraplength=780,
            justify="left",
        )
        self.last_export_label.pack(anchor="w")

        self.open_report_actions = tk.Frame(det_inner, bg=COLORS["white"])
        self.open_report_actions.pack(anchor="w", pady=(10, 0))

        self.open_file_btn = self._make_button(
            self.open_report_actions,
            "Open Report",
            self._open_last_report,
            style_kind="success",
        )
        self.open_folder_btn = self._make_button(
            self.open_report_actions,
            "Open Containing Folder",
            self._open_last_report_folder,
            style_kind="secondary",
        )

    # ------------------------------------------------------------------
    # Section 5: Application Information
    # ------------------------------------------------------------------

    def _build_info_section(self) -> None:
        frame = tk.Frame(self.content_host, bg=COLORS["bg"])
        frame.grid(row=0, column=0, sticky="nsew")
        self.sections["info"] = frame

        about_card = self._make_card(frame)
        about_card.pack(fill="x", pady=(0, 16))

        ab_inner = tk.Frame(about_card, bg=COLORS["white"])
        ab_inner.pack(fill="x", padx=20, pady=18)

        tk.Label(
            ab_inner,
            text="Data Quality Checker — Desktop Edition",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w")
        tk.Label(
            ab_inner,
            text=(
                "A local Windows desktop utility for inspecting CSV and Excel (.xlsx, .xls) datasets, "
                "detecting data quality issues, and exporting structured Excel audit reports without modifying source files."
            ),
            bg=COLORS["white"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
            wraplength=860,
            justify="left",
        ).pack(anchor="w", pady=(6, 10))

        tk.Label(
            ab_inner,
            text="Supported Inputs:  .csv, .xlsx, .xls   •   Stack:  Python • Tkinter/ttk • Pandas • OpenPyXL • xlrd",
            bg="#EFF6FF",
            fg=COLORS["accent"],
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=6,
        ).pack(anchor="w")

        rules_card = self._make_card(frame)
        rules_card.pack(fill="both", expand=True)

        ru_inner = tk.Frame(rules_card, bg=COLORS["white"])
        ru_inner.pack(fill="both", expand=True, padx=20, pady=18)

        tk.Label(
            ru_inner,
            text="Implemented Validation Checks",
            bg=COLORS["white"],
            fg=COLORS["text"],
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        rules = [
            ("1. Duplicate Detection", "Detects exact duplicate rows across all columns (retaining the first occurrence)."),
            ("2. Missing Values", "Counts empty/NaN cells per column and identifies all rows containing missing values."),
            (
                "3. Unified Data Validation (Types & Formats)",
                "Validates non-missing values once per cell without duplicate reporting: "
                "customer_id (integer), name (string), age (non-negative integer), "
                "email (string), and join_date (valid date in YYYY-MM-DD format).",
            ),
        ]

        for heading, detail in rules:
            block = tk.Frame(ru_inner, bg=COLORS["surface_alt"], highlightbackground=COLORS["border"], highlightthickness=1)
            block.pack(fill="x", pady=5)
            tk.Label(
                block,
                text=heading,
                bg=COLORS["surface_alt"],
                fg=COLORS["text"],
                font=("Segoe UI", 9, "bold"),
            ).pack(anchor="w", padx=12, pady=(8, 2))
            tk.Label(
                block,
                text=detail,
                bg=COLORS["surface_alt"],
                fg=COLORS["muted"],
                font=("Segoe UI", 9),
                wraplength=820,
                justify="left",
            ).pack(anchor="w", padx=12, pady=(0, 8))

    # ------------------------------------------------------------------
    # Actions & Backend Integration
    # ------------------------------------------------------------------

    def browse_file(self) -> None:
        """Prompt the user to choose a CSV or Excel file and load it via src.csv_reader.read_file."""
        initial_dir = _PROJECT_ROOT / "data"
        selected = filedialog.askopenfilename(
            title="Select a CSV or Excel Dataset",
            initialdir=str(initial_dir if initial_dir.exists() else _PROJECT_ROOT),
            filetypes=[
                ("Supported Data Files", "*.csv *.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("Excel Workbooks (*.xlsx)", "*.xlsx"),
                ("Legacy Excel Files (*.xls)", "*.xls"),
                ("All Files", "*.*"),
            ],
        )
        if not selected:
            return

        try:
            df = read_file(selected)
        except FileNotFoundError:
            self._set_status_badge("ERROR", COLORS["red_bg"], COLORS["red"])
            messagebox.showerror("File Not Found", "The selected file could not be found.")
            return
        except ValueError as err:
            self._set_status_badge("INVALID FILE", COLORS["red_bg"], COLORS["red"])
            messagebox.showerror("Invalid or Unsupported File", str(err))
            return
        except PermissionError:
            self._set_status_badge("ERROR", COLORS["red_bg"], COLORS["red"])
            messagebox.showerror(
                "Permission Denied",
                "Cannot read the selected file. Please close any program (such as Microsoft Excel) that has it open.",
            )
            return
        except OSError as err:
            self._set_status_badge("ERROR", COLORS["red_bg"], COLORS["red"])
            messagebox.showerror("File Error", f"Unable to open the selected file.\n\nDetails: {err.strerror or err}")
            return
        except Exception as err:
            self._set_status_badge("ERROR", COLORS["red_bg"], COLORS["red"])
            messagebox.showerror("Unexpected Error", f"Failed to load the selected file:\n\n{err}")
            return

        self.df = df
        self.file_path = Path(selected)
        self.results = None
        self.last_report_path = None

        self._set_status_badge("FILE LOADED", COLORS["accent_soft"], COLORS["accent"])
        self._refresh_all_views()

    def run_quality_check(self) -> None:
        """Execute duplicate, missing-value, and unified data validation checks."""
        if self.df is None:
            messagebox.showwarning(
                "No Dataset Selected",
                "Please select a CSV or Excel file first using 'Browse File' before running quality checks.",
            )
            return

        try:
            duplicate_rows, duplicate_count = find_duplicates(self.df)
            missing_count, rows_with_missing, total_missing = find_missing_values(self.df)
            validation_issues, validation_issue_count = validate_data(self.df)
        except Exception as err:
            self._set_status_badge("ERROR", COLORS["red_bg"], COLORS["red"])
            messagebox.showerror(
                "Validation Error",
                f"An error occurred while validating the dataset:\n\n{err}",
            )
            return

        total_issues = int(duplicate_count) + int(total_missing) + int(validation_issue_count)

        self.results = {
            "duplicate_rows": duplicate_rows,
            "duplicate_count": int(duplicate_count),
            "missing_count": missing_count,
            "rows_with_missing": rows_with_missing,
            "total_missing": int(total_missing),
            "validation_issues": validation_issues,
            "validation_issue_count": int(validation_issue_count),
            "total_issues": total_issues,
        }

        if total_issues == 0:
            self._set_status_badge("ALL CHECKS PASSED", COLORS["green_bg"], COLORS["green"])
        else:
            self._set_status_badge(f"{total_issues} ISSUE(S) FOUND", COLORS["red_bg"], COLORS["red"])

        self._refresh_all_views()

    def export_report(self) -> None:
        """Export validation results to an Excel workbook using src.report_generator.generate_report."""
        if self.df is None:
            messagebox.showwarning(
                "No Dataset Selected",
                "Please select a CSV or Excel file and run the quality check before exporting a report.",
            )
            return

        if self.results is None:
            messagebox.showwarning(
                "Quality Check Required",
                "Please click 'Run Quality Check' before exporting the Excel report.",
            )
            return

        reports_dir = _PROJECT_ROOT / "reports"
        default_stem = self.file_path.stem if self.file_path else "data"
        default_name = f"{default_stem}_quality_report.xlsx"

        destination = filedialog.asksaveasfilename(
            title="Save Excel Quality Report",
            initialdir=str(reports_dir if reports_dir.exists() else _PROJECT_ROOT),
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not destination:
            return

        try:
            output_path = generate_report(
                self.df,
                self.results["duplicate_rows"],
                self.results["duplicate_count"],
                self.results["missing_count"],
                self.results["rows_with_missing"],
                self.results["total_missing"],
                self.results["validation_issues"],
                self.results["validation_issue_count"],
                destination,
            )
        except PermissionError:
            messagebox.showerror(
                "File Is Open or Locked",
                "Cannot save the Excel report because the destination file is currently open or locked "
                "(for example, in Microsoft Excel).\n\nPlease close the file or choose a different name and try again.",
            )
            return
        except OSError as err:
            if err.errno in (errno.EACCES, errno.EPERM) or getattr(err, "winerror", None) == 32:
                messagebox.showerror(
                    "File Is Open or Locked",
                    "Cannot write to the selected Excel file because it is currently in use by another program.\n\n"
                    "Please close the file and try again.",
                )
            else:
                messagebox.showerror(
                    "Report Export Error",
                    f"Could not write the Excel report to disk.\n\nDetails: {err.strerror or err}",
                )
            return
        except Exception as err:
            messagebox.showerror(
                "Report Export Error",
                f"An unexpected error occurred while generating the Excel report:\n\n{err}",
            )
            return

        self.last_report_path = Path(output_path)
        self._refresh_reports()
        messagebox.showinfo(
            "Report Exported",
            f"Excel report generated successfully!\n\nSaved to:\n{self.last_report_path}",
        )

    def _open_last_report(self) -> None:
        if not self.last_report_path or not self.last_report_path.exists():
            messagebox.showinfo("Report Unavailable", "The exported report file could not be found.")
            return
        try:
            os.startfile(str(self.last_report_path))
        except Exception as err:
            messagebox.showerror("Open Report Error", f"Could not open the report file:\n\n{err}")

    def _open_last_report_folder(self) -> None:
        if not self.last_report_path or not self.last_report_path.parent.exists():
            messagebox.showinfo("Folder Unavailable", "The report folder could not be found.")
            return
        try:
            os.startfile(str(self.last_report_path.parent))
        except Exception as err:
            messagebox.showerror("Open Folder Error", f"Could not open the folder:\n\n{err}")

    # ------------------------------------------------------------------
    # View Refresh Logic
    # ------------------------------------------------------------------

    def _set_status_badge(self, text: str, bg: str, fg: str) -> None:
        self.status_badge.configure(text=text, bg=bg, fg=fg)

    def _refresh_all_views(self) -> None:
        self._refresh_sidebar_status()
        self._refresh_dashboard()
        self._refresh_preview()
        self._refresh_results()
        self._refresh_reports()

    def _refresh_sidebar_status(self) -> None:
        if self.df is None or self.file_path is None:
            self.sidebar_status_label.configure(text="●  No file loaded", fg="#E2E8F0")
        elif self.results is None:
            self.sidebar_status_label.configure(
                text=f"●  {self.file_path.name}\n    Loaded (not checked)",
                fg="#93C5FD",
            )
        elif self.results["total_issues"] == 0:
            self.sidebar_status_label.configure(
                text=f"●  {self.file_path.name}\n    All checks passed",
                fg="#86EFAC",
            )
        else:
            self.sidebar_status_label.configure(
                text=f"●  {self.file_path.name}\n    {self.results['total_issues']} issue(s) found",
                fg="#FCA5A5",
            )

    def _refresh_dashboard(self) -> None:
        if self.df is None or self.file_path is None:
            self.dash_filename_label.configure(text="No dataset file selected", fg=COLORS["muted"])
            self.dash_status_label.configure(text="Click 'Browse File' to load a CSV or Excel (.xlsx, .xls) dataset.")
            for key, card in self.stat_cards.items():
                card["value"].configure(text="—", fg=COLORS["text"])
                card["accent"].configure(bg=COLORS["border"])
                card["caption"].configure(text=card["default_caption"], fg=COLORS["muted"])

            self.dash_overview_text.configure(
                text=(
                    "Workflow:\n"
                    "1. Click 'Browse File' to select a CSV or Excel (.xlsx, .xls) dataset.\n"
                    "2. Inspect the loaded records in 'Data Preview'.\n"
                    "3. Click 'Run Quality Check' to run duplicate, missing-value, and unified type/format checks.\n"
                    "4. Review detailed findings in 'Validation Results' or export an Excel workbook in 'Reports'."
                ),
                fg=COLORS["muted"],
            )
            self.quick_preview_btn.configure(state="disabled")
            self.quick_results_btn.configure(state="disabled")
            self.quick_reports_btn.configure(state="disabled")
            return

        row_count = len(self.df)
        col_count = len(self.df.columns)
        self.dash_filename_label.configure(text=f"{self.file_path.name}   ({self.file_path})", fg=COLORS["text"])

        self._update_stat_card("total_rows", f"{row_count:,}", COLORS["accent"], "Records loaded from file", COLORS["muted"])
        self._update_stat_card("total_columns", f"{col_count:,}", COLORS["accent"], "Detected dataset fields", COLORS["muted"])

        self.quick_preview_btn.configure(state="normal")

        if self.results is None:
            self.dash_status_label.configure(
                text=f"Loaded {row_count:,} rows × {col_count:,} columns. Click 'Run Quality Check' to validate."
            )
            for key in ("duplicates", "missing", "validation_issues"):
                self._update_stat_card(key, "—", COLORS["border"], "Click 'Run Quality Check'", COLORS["muted"])

            schema_note = self._schema_notice_text()
            overview = f"Dataset '{self.file_path.name}' is loaded and ready for validation."
            if schema_note:
                overview += f"\n\n{schema_note}"
            self.dash_overview_text.configure(text=overview, fg=COLORS["text"])
            self.quick_results_btn.configure(state="normal")
            self.quick_reports_btn.configure(state="disabled")
            return

        dup = self.results["duplicate_count"]
        miss = self.results["total_missing"]
        val_issues = self.results["validation_issue_count"]
        total_issues = self.results["total_issues"]

        self.dash_status_label.configure(
            text=f"Validated {row_count:,} rows × {col_count:,} columns • Total issues detected: {total_issues:,}"
        )

        self._apply_issue_card("duplicates", dup, "No duplicate records", f"{dup:,} duplicate row(s)", COLORS["orange"])
        self._apply_issue_card("missing", miss, "No missing values", f"{miss:,} empty cell(s)", COLORS["orange"])
        self._apply_issue_card(
            "validation_issues",
            val_issues,
            "All checked types & formats valid",
            f"{val_issues:,} validation issue(s)",
            COLORS["red"],
        )

        if total_issues == 0:
            overview = (
                f"All quality checks passed on '{self.file_path.name}' with 0 issues detected. "
                "You can export a full Excel summary in 'Reports'."
            )
            self.dash_overview_text.configure(text=overview, fg=COLORS["green"])
        else:
            overview = (
                f"Quality check finished for '{self.file_path.name}' and detected {total_issues:,} total issue(s): "
                f"{dup:,} duplicate(s), {miss:,} missing value(s), and {val_issues:,} data validation issue(s).\n"
                "Open 'Validation Results' to inspect affected rows and values."
            )
            schema_note = self._schema_notice_text()
            if schema_note:
                overview += f"\n\n{schema_note}"
            self.dash_overview_text.configure(text=overview, fg=COLORS["text"])

        self.quick_results_btn.configure(state="normal")
        self.quick_reports_btn.configure(state="normal")

    def _update_stat_card(self, key: str, value_text: str, accent_color: str, caption_text: str, caption_color: str) -> None:
        card = self.stat_cards[key]
        card["value"].configure(text=value_text, fg=COLORS["text"] if value_text == "—" else accent_color)
        card["accent"].configure(bg=accent_color)
        card["caption"].configure(text=caption_text, fg=caption_color)

    def _apply_issue_card(self, key: str, count: int, ok_caption: str, issue_caption: str, issue_color: str) -> None:
        if count == 0:
            self._update_stat_card(key, "0", COLORS["green"], ok_caption, COLORS["green"])
        else:
            self._update_stat_card(key, f"{count:,}", issue_color, issue_caption, issue_color)

    def _refresh_preview(self) -> None:
        if self.df is None or self.file_path is None:
            self.preview_title_label.configure(text="No Dataset Loaded")
            self.preview_info_label.configure(text="Select a CSV or Excel file to preview its columns and rows.")
            self.preview_table.show_empty(
                "No dataset loaded yet.\nClick 'Browse File' to open a CSV or Excel file.",
                icon="▤",
                icon_color=COLORS["muted"],
            )
            return

        try:
            limit = int(self.preview_limit_var.get())
        except ValueError:
            limit = 500

        total_rows = len(self.df)
        total_cols = len(self.df.columns)
        shown_df = self.df.head(limit)
        shown_count = len(shown_df)

        self.preview_title_label.configure(text=self.file_path.name)
        if total_rows > shown_count:
            self.preview_info_label.configure(
                text=f"Showing first {shown_count:,} of {total_rows:,} rows  •  {total_cols:,} columns"
            )
        else:
            self.preview_info_label.configure(
                text=f"Showing all {shown_count:,} rows  •  {total_cols:,} columns"
            )

        data_cols = [str(col) for col in shown_df.columns]
        columns = ["Row #"] + data_cols
        rows = []
        for idx, record in shown_df.iterrows():
            rows.append((int(idx) + 2, *[format_cell_value(v) for v in record.tolist()]))

        widths = {"Row #": 80}
        self.preview_table.populate(columns, rows, column_widths=widths, stretch_last=False)

    def _refresh_results(self) -> None:
        if self.df is None:
            self.results_banner_title.configure(text="No Dataset Loaded", fg=COLORS["text"])
            self.results_banner_subtitle.configure(
                text="Select a CSV or Excel file first, then click 'Run Quality Check' to view validation results."
            )
            self.results_run_btn.configure(state="normal")
            for tab_info in self.category_tabs.values():
                self.results_notebook.tab(tab_info["frame"], text=f"  {tab_info['title']}  ")
                tab_info["badge"].configure(text="NO DATA", bg=COLORS["neutral_bg"], fg=COLORS["text"])
                tab_info["summary"].configure(text="Load a dataset file to run this check.")
                tab_info["note"].configure(text="")
                tab_info["table"].show_empty("No dataset loaded.", icon="—")
            return

        if self.results is None:
            self.results_banner_title.configure(text=f"Ready to Validate: {self.file_path.name}", fg=COLORS["text"])
            self.results_banner_subtitle.configure(
                text="Click 'Run Quality Check' to execute duplicate, missing-value, and unified type/format checks."
            )
            self.results_run_btn.configure(state="normal")
            for tab_info in self.category_tabs.values():
                self.results_notebook.tab(tab_info["frame"], text=f"  {tab_info['title']}  ")
                tab_info["badge"].configure(text="PENDING", bg=COLORS["accent_soft"], fg=COLORS["accent"])
                tab_info["summary"].configure(text="Waiting for 'Run Quality Check' to be executed.")
                tab_info["note"].configure(text="")
                tab_info["table"].show_empty(
                    "Click 'Run Quality Check' to inspect this category.",
                    icon="▶",
                    icon_color=COLORS["accent"],
                )
            return

        total_issues = self.results["total_issues"]
        schema_note = self._schema_notice_text()
        if total_issues == 0:
            self.results_banner_title.configure(text="All Quality Checks Passed", fg=COLORS["green"])
            subtitle = f"No issues were detected across {len(self.df):,} rows in '{self.file_path.name}'."
        else:
            self.results_banner_title.configure(text=f"{total_issues:,} Data Quality Issue(s) Detected", fg=COLORS["red"])
            subtitle = f"Inspected {len(self.df):,} rows and {len(self.df.columns):,} columns in '{self.file_path.name}'."
        if schema_note:
            subtitle += f"  ({schema_note})"
        self.results_banner_subtitle.configure(text=subtitle)

        self._populate_duplicates_tab()
        self._populate_missing_tab()
        self._populate_validation_tab()

    def _populate_duplicates_tab(self) -> None:
        tab = self.category_tabs["duplicates"]
        dup_rows: pd.DataFrame = self.results["duplicate_rows"]
        dup_count: int = self.results["duplicate_count"]

        self.results_notebook.tab(tab["frame"], text=f"  Duplicate Detection ({dup_count})  ")
        tab["note"].configure(text="")

        if dup_count == 0:
            tab["badge"].configure(text="PASSED", bg=COLORS["green_bg"], fg=COLORS["green"])
            tab["summary"].configure(text="0 duplicate records detected. All rows in the dataset are unique.")
            tab["table"].show_empty("No duplicate records found.", icon="✓", icon_color=COLORS["green"])
            return

        tab["badge"].configure(text=f"{dup_count} ISSUE(S)", bg=COLORS["orange_bg"], fg=COLORS["orange"])
        tab["summary"].configure(
            text=f"Detected {dup_count:,} duplicate record(s) (matching an earlier row across all columns)."
        )

        data_cols = [str(c) for c in dup_rows.columns]
        columns = ["Row #", "Affected Column"] + data_cols
        rows = []
        for idx, record in dup_rows.iterrows():
            rows.append((int(idx) + 2, "All Columns", *[format_cell_value(v) for v in record.tolist()]))

        tab["table"].populate(columns, rows, column_widths={"Row #": 80, "Affected Column": 130}, stretch_last=False)

    def _populate_missing_tab(self) -> None:
        tab = self.category_tabs["missing"]
        missing_count: pd.Series = self.results["missing_count"]
        rows_with_missing: pd.DataFrame = self.results["rows_with_missing"]
        total_missing: int = self.results["total_missing"]

        self.results_notebook.tab(tab["frame"], text=f"  Missing Values ({total_missing})  ")

        if total_missing == 0:
            tab["badge"].configure(text="PASSED", bg=COLORS["green_bg"], fg=COLORS["green"])
            tab["summary"].configure(text="0 missing values detected across all columns.")
            tab["note"].configure(text="")
            tab["table"].show_empty("No missing values found.", icon="✓", icon_color=COLORS["green"])
            return

        affected_cols = [f"{col}: {int(cnt)}" for col, cnt in missing_count.items() if cnt > 0]
        tab["badge"].configure(text=f"{total_missing} ISSUE(S)", bg=COLORS["orange_bg"], fg=COLORS["orange"])
        tab["summary"].configure(
            text=f"Detected {total_missing:,} missing value(s) across {len(rows_with_missing):,} row(s)."
        )
        tab["note"].configure(text="Missing count by column:  " + "   •   ".join(affected_cols))

        columns = ["Row #", "Affected Column", "Invalid Value", "Row Context"]
        rows = []
        for idx, record in rows_with_missing.iterrows():
            for col in rows_with_missing.columns:
                if is_missing_value(record[col]):
                    rows.append((int(idx) + 2, str(col), "<Missing>", format_row_context(record, exclude_col=col)))

        tab["table"].populate(
            columns,
            rows,
            column_widths={"Row #": 80, "Affected Column": 160, "Invalid Value": 130, "Row Context": 520},
            stretch_last=True,
        )

    def _populate_validation_tab(self) -> None:
        tab = self.category_tabs["validation"]
        issues_df: pd.DataFrame = self.results["validation_issues"]
        count: int = self.results["validation_issue_count"]

        self.results_notebook.tab(tab["frame"], text=f"  Data Validation ({count})  ")

        checked_cols = [c for c in EXPECTED_COLUMNS if c in self.df.columns]
        missing_cols = [c for c in EXPECTED_COLUMNS if c not in self.df.columns]

        note_parts = []
        if checked_cols:
            note_parts.append("Checked columns: " + ", ".join(checked_cols))
        if missing_cols:
            note_parts.append("Not present in file (skipped): " + ", ".join(missing_cols))
        tab["note"].configure(text="   |   ".join(note_parts))

        if not checked_cols:
            tab["badge"].configure(text="SKIPPED", bg=COLORS["neutral_bg"], fg=COLORS["text"])
            tab["summary"].configure(
                text="None of the configured schema columns (customer_id, name, age, email, join_date) exist in this file."
            )
            tab["table"].show_empty(
                "No expected schema columns were found in this dataset for type/format validation.",
                icon="ⓘ",
                icon_color=COLORS["muted"],
            )
            return

        if count == 0:
            tab["badge"].configure(text="PASSED", bg=COLORS["green_bg"], fg=COLORS["green"])
            tab["summary"].configure(text="0 data type or format issues detected in the checked columns.")
            tab["table"].show_empty(
                "All checked column values match their expected data types and formats.",
                icon="✓",
                icon_color=COLORS["green"],
            )
            return

        affected_cols = issues_df["Column"].nunique()
        tab["badge"].configure(text=f"{count} ISSUE(S)", bg=COLORS["red_bg"], fg=COLORS["red"])
        tab["summary"].configure(
            text=f"Detected {count:,} data validation issue(s) across {affected_cols} column(s) (reported once per cell)."
        )

        columns = ["Row #", "Affected Column", "Validation Rule", "Invalid Value", "Issue Detail", "Row Context"]
        rows = []
        for _, item in issues_df.iterrows():
            row_number = int(item["Row Number"])
            col = str(item["Column"])
            df_idx = row_number - 2
            context = format_row_context(self.df.loc[df_idx], exclude_col=col) if df_idx in self.df.index else ""
            rows.append((
                row_number,
                col,
                str(item["Validation Rule"]),
                format_cell_value(item["Invalid Value"]),
                str(item["Issue"]),
                context,
            ))

        tab["table"].populate(
            columns,
            rows,
            column_widths={
                "Row #": 75,
                "Affected Column": 135,
                "Validation Rule": 155,
                "Invalid Value": 145,
                "Issue Detail": 230,
                "Row Context": 320,
            },
            stretch_last=True,
        )

    def _refresh_reports(self) -> None:
        if self.df is None or self.file_path is None:
            self.reports_status_label.configure(
                text="No dataset file selected. Load a CSV or Excel file and run the quality check before exporting.",
                fg=COLORS["muted"],
            )
        elif self.results is None:
            self.reports_status_label.configure(
                text=f"'{self.file_path.name}' is loaded, but validation has not been run yet. "
                "Click 'Run Quality Check' first to generate report data.",
                fg=COLORS["orange"],
            )
        else:
            self.reports_status_label.configure(
                text=f"Ready to export quality report for '{self.file_path.name}' "
                f"({self.results['total_issues']:,} total issue(s) recorded).",
                fg=COLORS["green"],
            )

        if self.last_report_path and self.last_report_path.exists():
            self.last_export_label.configure(
                text=f"Last exported report:  {self.last_report_path}",
                fg=COLORS["green"],
            )
            self.open_file_btn.pack(side="left", padx=(0, 10))
            self.open_folder_btn.pack(side="left")
        else:
            self.last_export_label.configure(
                text="No report has been exported in this session yet.",
                fg=COLORS["muted"],
            )
            self.open_file_btn.pack_forget()
            self.open_folder_btn.pack_forget()

    def _schema_notice_text(self) -> str:
        if self.df is None:
            return ""
        expected = set(EXPECTED_COLUMNS)
        present = set(self.df.columns)
        missing = sorted(expected - present)
        unexpected = sorted(present - expected)
        parts = []
        if missing:
            parts.append(f"Expected columns not in file: {', '.join(missing)}")
        if unexpected:
            parts.append(f"Additional columns: {', '.join(unexpected)}")
        return " | ".join(parts)


def enable_windows_dpi_awareness() -> None:
    """Enable sharp DPI rendering on Windows displays when supported."""
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


def launch_app() -> None:
    """Create the Tk root window and start the Data Quality Checker desktop app."""
    enable_windows_dpi_awareness()
    root = tk.Tk()
    DataQualityApp(root)
    root.mainloop()


if __name__ == "__main__":
    launch_app()
