"""Desktop GUI for KPI Benchmark & Trial Analysis Tool.

Built with Tkinter & TTK:
- Automatic detection of Excel and CSV files in data/
- Dropdown selectors for Implementation Date and Pre/Post periods
- Instant search and multi-selection of KPIs/counters
- Live on-screen chart preview with official corporate theme and implementation marker
- One-click PowerPoint deck export to output/ folder
"""

from __future__ import annotations

import glob
import os
import subprocess
import sys
import threading
from datetime import datetime
from typing import Dict, List, Optional, Set

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from src.chart_engine import ChartEngine
from src.data_loader import DataLoader
from src.ppt_generator import PowerPointGenerator


class KPIBenchmarkApp:
    """Main GUI Application for KPI Benchmark & Trial Analysis."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("KPI Benchmark & Trial Analysis Tool")
        self.root.geometry("1240x820")
        self.root.minsize(1050, 700)

        # Force window to foreground on Windows
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        try:
            self.root.attributes("-topmost", True)
            self.root.after_idle(self.root.attributes, "-topmost", False)
        except Exception:
            pass

        # Apply Modern Styling
        self._setup_styles()

        # Core Engines
        self.chart_engine = ChartEngine()
        self.ppt_generator = PowerPointGenerator(self.chart_engine)

        # State Variables
        self.loader: Optional[DataLoader] = None
        self.current_file_path = tk.StringVar()
        self.impl_date_var = tk.StringVar()
        self.pre_start_var = tk.StringVar()
        self.pre_end_var = tk.StringVar()
        self.post_start_var = tk.StringVar()
        self.post_end_var = tk.StringVar()
        self.search_filter_var = tk.StringVar()
        self.output_dir_var = tk.StringVar(value=os.path.abspath("output"))
        self.template_path_var = tk.StringVar(value="")
        self.include_summary_table_var = tk.BooleanVar(value=True)

        # KPI Selection State
        self.selected_kpis_set: Set[str] = set()
        self.kpi_check_vars: Dict[str, tk.BooleanVar] = {}
        self.visible_kpis: List[str] = []
        self.current_preview_kpi: Optional[str] = None
        self.canvas_widget = None

        # Build Layout
        self._build_ui()

        # Enforce initial sash position so left sidebar is immediately visible without dragging
        self.root.after(30, self._ensure_sash_position)
        self.root.after(100, self._ensure_sash_position)
        self.root.after(300, self._ensure_sash_position)

        # Auto-detect starter data
        self.root.after(200, self._auto_detect_initial_file)

    def _setup_styles(self) -> None:
        """Sets up custom ttk styling with corporate theme."""
        style = ttk.Style()
        style.theme_use("clam")

        # Palette
        c_primary = "#1B365D"
        c_light_bg = "#F8FAFC"
        c_border = "#CBD5E1"
        c_dark_text = "#1E293B"

        self.root.configure(bg=c_light_bg)

        style.configure(".", background=c_light_bg, foreground=c_dark_text, font=("Segoe UI", 9))
        style.configure("TLabel", background=c_light_bg, foreground=c_dark_text)
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground=c_primary)
        style.configure("SubHeader.TLabel", font=("Segoe UI", 9, "bold"), foreground="#475569")
        style.configure("Badge.TLabel", font=("Segoe UI", 9, "bold"), foreground="#2563EB")
        
        style.configure("TLabelframe", background=c_light_bg, bordercolor=c_border, relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", background=c_light_bg, foreground=c_primary, font=("Segoe UI", 9, "bold"))

        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background="#1B365D",
            foreground="#FFFFFF",
            padding=(10, 6)
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#2563EB"), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#F1F5F9")]
        )

        style.configure("TButton", padding=(6, 3), font=("Segoe UI", 9))

    def _build_ui(self) -> None:
        """Constructs split-pane layout: Left (Configuration & Selection) and Right (Live Preview)."""
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Snap back sash if user accidentally drags it too narrow
        self.main_paned.bind("<ButtonRelease-1>", lambda e: self._on_sash_released())

        # ==========================================
        # LEFT PANEL: Controls & KPI Selection
        # ==========================================
        self.left_frame = ttk.Frame(self.main_paned, width=480)
        self.main_paned.add(self.left_frame, weight=0)
        left_frame = self.left_frame

        # 1. File Selection Card
        file_card = ttk.LabelFrame(left_frame, text=" 1. Data Source (Excel / CSV) ", padding=8)
        file_card.pack(fill=tk.X, pady=(0, 6))

        file_input_row = ttk.Frame(file_card)
        file_input_row.pack(fill=tk.X)

        self.file_entry = ttk.Entry(file_input_row, textvariable=self.current_file_path, font=("Segoe UI", 9))
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        browse_btn = ttk.Button(file_input_row, text="Browse...", command=self._on_browse_file)
        browse_btn.pack(side=tk.LEFT)

        self.file_status_lbl = ttk.Label(file_card, text="No file loaded", font=("Segoe UI", 8), foreground="#64748B")
        self.file_status_lbl.pack(anchor=tk.W, pady=(4, 0))

        # 2. Period & Implementation Date Configuration Card
        period_card = ttk.LabelFrame(left_frame, text=" 2. Period & Implementation Date ", padding=8)
        period_card.pack(fill=tk.X, pady=(0, 6))

        # Implementation Date Row (Highlighted)
        impl_row = ttk.Frame(period_card)
        impl_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(impl_row, text="Implementation Date:", font=("Segoe UI", 9, "bold"), foreground="#DC2626").pack(side=tk.LEFT)
        self.impl_combo = ttk.Combobox(impl_row, textvariable=self.impl_date_var, state="readonly", width=14)
        self.impl_combo.pack(side=tk.RIGHT)
        self.impl_combo.bind("<<ComboboxSelected>>", self._on_impl_date_changed)

        # Pre-Period Row
        pre_row = ttk.Frame(period_card)
        pre_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(pre_row, text="Pre-Period:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.pre_end_combo = ttk.Combobox(pre_row, textvariable=self.pre_end_var, state="readonly", width=12)
        self.pre_end_combo.pack(side=tk.RIGHT)
        self.pre_end_combo.bind("<<ComboboxSelected>>", self._on_period_changed)
        ttk.Label(pre_row, text=" to ", font=("Segoe UI", 8)).pack(side=tk.RIGHT, padx=2)
        self.pre_start_combo = ttk.Combobox(pre_row, textvariable=self.pre_start_var, state="readonly", width=12)
        self.pre_start_combo.pack(side=tk.RIGHT)
        self.pre_start_combo.bind("<<ComboboxSelected>>", self._on_period_changed)

        # Post-Period Row
        post_row = ttk.Frame(period_card)
        post_row.pack(fill=tk.X)
        ttk.Label(post_row, text="Post-Period:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.post_end_combo = ttk.Combobox(post_row, textvariable=self.post_end_var, state="readonly", width=12)
        self.post_end_combo.pack(side=tk.RIGHT)
        self.post_end_combo.bind("<<ComboboxSelected>>", self._on_period_changed)
        ttk.Label(post_row, text=" to ", font=("Segoe UI", 8)).pack(side=tk.RIGHT, padx=2)
        self.post_start_combo = ttk.Combobox(post_row, textvariable=self.post_start_var, state="readonly", width=12)
        self.post_start_combo.pack(side=tk.RIGHT)
        self.post_start_combo.bind("<<ComboboxSelected>>", self._on_period_changed)

        # 3. KPI & Counter Selection Card
        kpi_card = ttk.LabelFrame(left_frame, text=" 3. Select KPIs & Counters for Report ", padding=8)
        kpi_card.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # Search Bar
        search_row = ttk.Frame(kpi_card)
        search_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(search_row, text="Search:", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_filter_var, font=("Segoe UI", 9))
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.search_filter_var.trace_add("write", lambda *args: self._filter_kpis())

        # Action Buttons for Selection
        btn_row = ttk.Frame(kpi_card)
        btn_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(btn_row, text="Select All", command=self._select_all_kpis).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text="Clear All", command=self._clear_all_kpis).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text="Select Filtered", command=self._select_filtered_kpis).pack(side=tk.LEFT)
        
        self.selection_count_lbl = ttk.Label(btn_row, text="Selected: 0", font=("Segoe UI", 9, "bold"), foreground="#2563EB")
        self.selection_count_lbl.pack(side=tk.RIGHT)

        # Scrollable KPI List
        kpi_list_frame = ttk.Frame(kpi_card)
        kpi_list_frame.pack(fill=tk.BOTH, expand=True)

        self.kpi_canvas = tk.Canvas(kpi_list_frame, bg="#FFFFFF", highlightthickness=1, highlightbackground="#CBD5E1")
        self.kpi_scrollbar = ttk.Scrollbar(kpi_list_frame, orient=tk.VERTICAL, command=self.kpi_canvas.yview)
        self.kpi_inner_frame = ttk.Frame(self.kpi_canvas)
        self.kpi_inner_frame.configure(style="TFrame")

        self.kpi_inner_frame.bind(
            "<Configure>",
            lambda e: self.kpi_canvas.configure(scrollregion=self.kpi_canvas.bbox("all"))
        )

        self.kpi_canvas_window = self.kpi_canvas.create_window((0, 0), window=self.kpi_inner_frame, anchor="nw")
        self.kpi_canvas.configure(yscrollcommand=self.kpi_scrollbar.set)

        self.kpi_canvas.bind("<Configure>", self._on_canvas_configure)
        self.kpi_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.kpi_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.kpi_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 4. PowerPoint Generation Card
        gen_card = ttk.LabelFrame(left_frame, text=" 4. Generate PowerPoint Deck ", padding=8)
        gen_card.pack(fill=tk.X)

        out_row = ttk.Frame(gen_card)
        out_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(out_row, text="Output Folder:", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
        self.out_entry = ttk.Entry(out_row, textvariable=self.output_dir_var, font=("Segoe UI", 8))
        self.out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(out_row, text="Browse", command=self._on_browse_output).pack(side=tk.LEFT)

        # Template Row (Optional)
        tmpl_row = ttk.Frame(gen_card)
        tmpl_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(tmpl_row, text="Template (Opt):", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
        self.tmpl_entry = ttk.Entry(tmpl_row, textvariable=self.template_path_var, font=("Segoe UI", 8))
        self.tmpl_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(tmpl_row, text="Browse", command=self._on_browse_template).pack(side=tk.LEFT)

        # Summary Table Checkbox Option
        table_opt_row = ttk.Frame(gen_card)
        table_opt_row.pack(fill=tk.X, pady=(0, 6))
        self.summary_table_cb = ttk.Checkbutton(
            table_opt_row,
            text="Include Summary Table on Slide 1 (Editable Table)",
            variable=self.include_summary_table_var,
        )
        self.summary_table_cb.pack(side=tk.LEFT)

        self.generate_btn = ttk.Button(
            gen_card,
            text="Generate PowerPoint Report (.pptx)",
            style="Primary.TButton",
            command=self._on_generate_report
        )
        self.generate_btn.pack(fill=tk.X, pady=(2, 4))

        self.progress_bar = ttk.Progressbar(gen_card, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(2, 2))

        self.status_lbl = ttk.Label(gen_card, text="Ready", font=("Segoe UI", 8), foreground="#475569")
        self.status_lbl.pack(anchor=tk.W)

        # ==========================================
        # RIGHT PANEL: Live Chart Preview
        # ==========================================
        self.right_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.right_frame, weight=1)
        right_frame = self.right_frame

        preview_header_card = ttk.Frame(right_frame, padding=(10, 6))
        preview_header_card.pack(fill=tk.X)

        self.preview_title_lbl = ttk.Label(
            preview_header_card,
            text="Interactive Chart Preview",
            style="Header.TLabel"
        )
        self.preview_title_lbl.pack(side=tk.LEFT)

        self.preview_hint_lbl = ttk.Label(
            preview_header_card,
            text="(Click any KPI on the left to preview its benchmark chart)",
            font=("Segoe UI", 9),
            foreground="#64748B"
        )
        self.preview_hint_lbl.pack(side=tk.LEFT, padx=10)

        # Canvas container for Matplotlib
        self.preview_container = ttk.Frame(right_frame, padding=6)
        self.preview_container.pack(fill=tk.BOTH, expand=True)

        self._show_initial_preview_placeholder()

    def _ensure_sash_position(self, target_width: int = 480) -> None:
        """Guarantees the left control sidebar is fully visible and not collapsed."""
        try:
            self.root.update_idletasks()
            current_pos = self.main_paned.sashpos(0)
            if current_pos < 380:
                self.main_paned.sashpos(0, target_width)
        except Exception:
            pass

    def _on_sash_released(self) -> None:
        """Snaps sash back if user accidentally dragged it too narrow."""
        try:
            if self.main_paned.sashpos(0) < 360:
                self.main_paned.sashpos(0, 480)
        except Exception:
            pass

    def _on_canvas_configure(self, event):
        """Ensures inner frame matches canvas width."""
        self.kpi_canvas.itemconfig(self.kpi_canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        """Allows mouse scrolling across the KPI list."""
        if self.kpi_canvas.winfo_ismapped():
            self.kpi_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _show_initial_preview_placeholder(self):
        """Displays a clean placeholder message in the preview panel."""
        if self.canvas_widget:
            self.canvas_widget.destroy()
            self.canvas_widget = None

        placeholder = ttk.Label(
            self.preview_container,
            text="No KPI selected for preview.\n\nSelect a data file and click on any KPI to preview its chart.",
            font=("Segoe UI", 12),
            foreground="#94A3B8",
            justify=tk.CENTER
        )
        placeholder.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        self.placeholder_widget = placeholder

    def _auto_detect_initial_file(self) -> None:
        """Looks for .csv or .xlsx files in ./data/ on launch."""
        possible_data_dirs = [
            os.path.abspath("data"),
            os.path.join(os.path.dirname(sys.executable), "data"),
        ]
        if hasattr(sys, "_MEIPASS"):
            possible_data_dirs.append(os.path.join(sys._MEIPASS, "data"))

        found_files = []
        for data_dir in possible_data_dirs:
            if data_dir and os.path.exists(data_dir):
                patterns = [os.path.join(data_dir, "*.csv"), os.path.join(data_dir, "*.xlsx"), os.path.join(data_dir, "*.xls")]
                for pat in patterns:
                    found_files.extend(glob.glob(pat))
                if found_files:
                    break

        if found_files:
            # Prefer CSV or latest modified file
            found_files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
            target_file = found_files[0]
            self.load_file(target_file)

        # Check for templates in ./templates/
        possible_tmpl_dirs = [
            os.path.abspath("templates"),
            os.path.join(os.path.dirname(sys.executable), "templates"),
        ]
        if hasattr(sys, "_MEIPASS"):
            possible_tmpl_dirs.append(os.path.join(sys._MEIPASS, "templates"))

        for tmpl_dir in possible_tmpl_dirs:
            if tmpl_dir and os.path.exists(tmpl_dir):
                for ext in ["*.pptx", "*.potx"]:
                    found_tmpls = glob.glob(os.path.join(tmpl_dir, ext))
                    for t in found_tmpls:
                        try:
                            with open(t, "rb") as f_chk:
                                f_chk.read(32)
                            self.template_path_var.set(t)
                            break
                        except Exception:
                            continue
                    if self.template_path_var.get():
                        break
                if self.template_path_var.get():
                    break

    def _on_browse_file(self) -> None:
        """Opens file dialog for user to select Excel or CSV file."""
        initial_dir = os.path.abspath("data") if os.path.exists("data") else os.getcwd()
        file_selected = filedialog.askopenfilename(
            title="Select KPI Data File",
            initialdir=initial_dir,
            filetypes=[
                ("All Supported Files", "*.xlsx;*.xls;*.csv"),
                ("Excel Files", "*.xlsx;*.xls"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ]
        )
        if file_selected:
            self.load_file(file_selected)

    def _on_browse_output(self) -> None:
        """Selects output directory."""
        dir_selected = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_dir_var.get()
        )
        if dir_selected:
            self.output_dir_var.set(dir_selected)

    def _on_browse_template(self) -> None:
        """Allows user to select a corporate PowerPoint template (.pptx or .potx)."""
        initial_dir = os.path.abspath("templates") if os.path.exists("templates") else os.getcwd()
        file_selected = filedialog.askopenfilename(
            title="Select Corporate PowerPoint Template (.pptx or .potx)",
            initialdir=initial_dir,
            filetypes=[
                ("PowerPoint Presentations & Templates", "*.pptx;*.potx"),
                ("PowerPoint Presentation", "*.pptx"),
                ("PowerPoint Template", "*.potx"),
                ("All Files", "*.*")
            ]
        )
        if file_selected:
            self.template_path_var.set(file_selected)

    def load_file(self, file_path: str) -> None:
        """Loads data from selected file, populates dates and KPIs."""
        try:
            self.status_lbl.config(text=f"Loading {os.path.basename(file_path)}...")
            self.root.update_idletasks()

            self.loader = DataLoader(file_path)
            self.current_file_path.set(file_path)

            file_name = os.path.basename(file_path)
            dates_cnt = len(self.loader.available_dates)
            kpis_cnt = len(self.loader.numeric_kpis)
            self.file_status_lbl.config(
                text=f"✓ Loaded '{file_name}': {dates_cnt} dates ({self.loader.available_dates[0]} to {self.loader.available_dates[-1]}), {kpis_cnt} numeric KPIs",
                foreground="#166534"
            )

            # Populate Date Dropdowns
            dates = self.loader.available_dates
            self.impl_combo["values"] = dates
            self.pre_start_combo["values"] = dates
            self.pre_end_combo["values"] = dates
            self.post_start_combo["values"] = dates
            self.post_end_combo["values"] = dates

            # Auto-suggest periods based on mid-point
            sug = self.loader.get_suggested_periods()
            self.impl_date_var.set(sug["impl_date"])
            self.pre_start_var.set(sug["pre_start"])
            self.pre_end_var.set(sug["pre_end"])
            self.post_start_var.set(sug["post_start"])
            self.post_end_var.set(sug["post_end"])

            # Populate KPI list
            self._populate_kpi_list()

            # Pre-select first 5 key KPIs or any matching keywords
            preferred_keywords = ["rate", "volume", "att", "succ", "cqi", "prb", "cce"]
            auto_selected = []
            for kpi in self.loader.numeric_kpis:
                if any(kw in kpi.lower() for kw in preferred_keywords):
                    auto_selected.append(kpi)
                if len(auto_selected) >= 4:
                    break
            if not auto_selected and self.loader.numeric_kpis:
                auto_selected = self.loader.numeric_kpis[:3]

            for kpi in auto_selected:
                self.selected_kpis_set.add(kpi)
                if kpi in self.kpi_check_vars:
                    self.kpi_check_vars[kpi].set(True)

            self._update_selection_counter()

            # Show preview of first selected KPI
            if auto_selected:
                self.preview_kpi(auto_selected[0])

            self.status_lbl.config(text="Ready to generate report")

        except Exception as e:
            messagebox.showerror("Error Loading File", f"Failed to load file:\n{str(e)}")
            self.file_status_lbl.config(text=f"Error loading file: {str(e)}", foreground="#DC2626")
            self.status_lbl.config(text="Error loading file")

    def _on_impl_date_changed(self, event=None) -> None:
        """Auto-adjusts Pre and Post period suggestions when Implementation Date is picked."""
        impl_date = self.impl_date_var.get()
        if not self.loader or not impl_date:
            return

        sug = self.loader.get_suggested_periods(impl_date)
        self.pre_start_var.set(sug["pre_start"])
        self.pre_end_var.set(sug["pre_end"])
        self.post_start_var.set(sug["post_start"])
        self.post_end_var.set(sug["post_end"])

        # Refresh preview if active
        if self.current_preview_kpi:
            self.preview_kpi(self.current_preview_kpi)

    def _on_period_changed(self, event=None) -> None:
        """Refreshes preview when period boundaries change."""
        if self.current_preview_kpi:
            self.preview_kpi(self.current_preview_kpi)

    def _populate_kpi_list(self) -> None:
        """Constructs scrollable list of checkboxes for KPIs."""
        # Clear existing widgets
        for widget in self.kpi_inner_frame.winfo_children():
            widget.destroy()

        self.kpi_check_vars.clear()
        self.selected_kpis_set.clear()

        if not self.loader:
            return

        for kpi in self.loader.numeric_kpis:
            var = tk.BooleanVar(value=False)
            self.kpi_check_vars[kpi] = var

        self._filter_kpis()

    def _filter_kpis(self) -> None:
        """Filters visible checkboxes according to search text."""
        query = self.search_filter_var.get().strip().lower()

        # Clear existing rows
        for widget in self.kpi_inner_frame.winfo_children():
            widget.destroy()

        if not self.loader:
            return

        self.visible_kpis = [
            kpi for kpi in self.loader.numeric_kpis
            if (not query or query in kpi.lower())
        ]

        for kpi in self.visible_kpis:
            row_frame = tk.Frame(self.kpi_inner_frame, bg="#FFFFFF")
            row_frame.pack(fill=tk.X, padx=2, pady=1)

            # Checkbox
            var = self.kpi_check_vars[kpi]
            cb = ttk.Checkbutton(
                row_frame,
                variable=var,
                command=lambda k=kpi, v=var: self._on_checkbox_toggle(k, v)
            )
            cb.pack(side=tk.LEFT, padx=(2, 4))

            # Clickable Label that activates preview
            lbl = tk.Label(
                row_frame,
                text=kpi,
                font=("Segoe UI", 9),
                bg="#FFFFFF",
                fg="#1E293B",
                anchor="w",
                cursor="hand2"
            )
            lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)
            lbl.bind("<Button-1>", lambda e, k=kpi: self._on_kpi_label_click(k))

    def _on_checkbox_toggle(self, kpi: str, var: tk.BooleanVar) -> None:
        """Tracks selected KPIs and triggers preview."""
        if var.get():
            self.selected_kpis_set.add(kpi)
            self.preview_kpi(kpi)
        else:
            self.selected_kpis_set.discard(kpi)
        self._update_selection_counter()

    def _on_kpi_label_click(self, kpi: str) -> None:
        """Clicking KPI label toggles its checkbox and opens its chart preview."""
        var = self.kpi_check_vars[kpi]
        new_val = not var.get()
        var.set(new_val)
        if new_val:
            self.selected_kpis_set.add(kpi)
        else:
            self.selected_kpis_set.discard(kpi)
        self._update_selection_counter()
        self.preview_kpi(kpi)

    def _update_selection_counter(self) -> None:
        """Updates the selected counter label."""
        total = len(self.loader.numeric_kpis) if self.loader else 0
        cnt = len(self.selected_kpis_set)
        self.selection_count_lbl.config(text=f"Selected: {cnt} of {total}")

    def _select_all_kpis(self) -> None:
        """Selects all available KPIs."""
        if not self.loader:
            return
        for kpi in self.loader.numeric_kpis:
            self.selected_kpis_set.add(kpi)
            if kpi in self.kpi_check_vars:
                self.kpi_check_vars[kpi].set(True)
        self._update_selection_counter()

    def _clear_all_kpis(self) -> None:
        """Clears all selections."""
        self.selected_kpis_set.clear()
        for var in self.kpi_check_vars.values():
            var.set(False)
        self._update_selection_counter()

    def _select_filtered_kpis(self) -> None:
        """Selects all currently visible/filtered KPIs."""
        for kpi in self.visible_kpis:
            self.selected_kpis_set.add(kpi)
            self.kpi_check_vars[kpi].set(True)
        self._update_selection_counter()

    def preview_kpi(self, kpi_name: str) -> None:
        """Renders live chart preview in the right panel."""
        if not self.loader:
            return

        self.current_preview_kpi = kpi_name
        self.preview_title_lbl.config(text=f"Preview: {kpi_name}")
        self.preview_hint_lbl.config(text="")

        if hasattr(self, "placeholder_widget") and self.placeholder_widget:
            self.placeholder_widget.destroy()
            self.placeholder_widget = None

        if self.canvas_widget:
            self.canvas_widget.destroy()
            self.canvas_widget = None

        pre_start = self.pre_start_var.get()
        pre_end = self.pre_end_var.get()
        impl_date = self.impl_date_var.get()
        post_start = self.post_start_var.get()
        post_end = self.post_end_var.get()

        fig = self.chart_engine.create_figure(
            loader=self.loader,
            kpi_name=kpi_name,
            pre_start=pre_start,
            pre_end=pre_end,
            impl_date=impl_date,
            post_start=post_start,
            post_end=post_end,
            figsize=(8.5, 4.8),
            dpi=110,
            show_title_in_chart=True
        )

        canvas = FigureCanvasTkAgg(fig, master=self.preview_container)
        self.canvas_widget = canvas.get_tk_widget()
        self.canvas_widget.pack(fill=tk.BOTH, expand=True)
        canvas.draw()
        self._ensure_sash_position()

    def _on_generate_report(self) -> None:
        """Starts background thread to generate PowerPoint deck."""
        if not self.loader:
            messagebox.showwarning("No Data", "Please load an Excel or CSV file first.")
            return

        if not self.selected_kpis_set:
            messagebox.showwarning("No KPIs Selected", "Please select at least one KPI/counter to include in the report.")
            return

        out_dir = self.output_dir_var.get().strip()
        if not out_dir:
            out_dir = os.path.abspath("output")
            self.output_dir_var.set(out_dir)

        # Disable button during generation
        self.generate_btn.config(state="disabled")
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.selected_kpis_set)
        self.status_lbl.config(text="Starting PowerPoint generation...", foreground="#2563EB")

        # Run in separate thread
        thread = threading.Thread(target=self._run_generation_task, daemon=True)
        thread.start()

    def _run_generation_task(self) -> None:
        """Worker thread for building PowerPoint presentation."""
        try:
            out_dir = self.output_dir_var.get().strip()
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"KPI_Benchmark_Report_{timestamp_str}.pptx"
            full_output_path = os.path.join(out_dir, output_filename)

            selected_list = sorted(list(self.selected_kpis_set))

            def update_progress(curr, total, kpi_name):
                def gui_update():
                    self.progress_bar["value"] = curr
                    self.status_lbl.config(
                        text=f"Generating slide {curr} of {total}: {kpi_name}...",
                        foreground="#2563EB"
                    )
                self.root.after(0, gui_update)

            template_path = self.template_path_var.get().strip() or None
            self.ppt_generator.generate_report(
                loader=self.loader,
                selected_kpis=selected_list,
                pre_start=self.pre_start_var.get(),
                pre_end=self.pre_end_var.get(),
                impl_date=self.impl_date_var.get(),
                post_start=self.post_start_var.get(),
                post_end=self.post_end_var.get(),
                output_path=full_output_path,
                template_path=template_path,
                include_summary_table=self.include_summary_table_var.get(),
                progress_callback=update_progress
            )

            # Completion notification on GUI thread
            def on_complete():
                self.generate_btn.config(state="normal")
                self.progress_bar["value"] = len(selected_list)
                self.status_lbl.config(
                    text=f"✓ Report generated: {output_filename}",
                    foreground="#166534"
                )
                self._show_success_dialog(full_output_path)

            self.root.after(0, on_complete)

        except Exception as e:
            def on_error():
                self.generate_btn.config(state="normal")
                self.status_lbl.config(text="Generation failed", foreground="#DC2626")
                messagebox.showerror("Generation Error", f"Failed to generate PowerPoint deck:\n{str(e)}")

            self.root.after(0, on_error)

    def _show_success_dialog(self, pptx_path: str) -> None:
        """Displays completion modal with options to open file or folder."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Report Generated Successfully")
        dialog.geometry("520x220")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frame,
            text="✓ PowerPoint Report Created Successfully!",
            font=("Segoe UI", 12, "bold"),
            foreground="#166534"
        ).pack(anchor=tk.W, pady=(0, 8))

        msg = f"Saved to:\n{pptx_path}\n\nTotal Slides: {len(self.selected_kpis_set) + 1} (Title + {len(self.selected_kpis_set)} KPI benchmark slides)"
        ttk.Label(frame, text=msg, font=("Segoe UI", 9), foreground="#334155").pack(anchor=tk.W, pady=(0, 15))

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X)

        def open_file():
            try:
                os.startfile(pptx_path)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}")

        def open_folder():
            try:
                folder = os.path.dirname(os.path.abspath(pptx_path))
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open folder: {e}")

        ttk.Button(btn_row, text="Open Presentation", style="Primary.TButton", command=open_file).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="Open Output Folder", command=open_folder).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="Close", command=dialog.destroy).pack(side=tk.RIGHT)


def launch_gui():
    """Starts the application."""
    log_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "kpi_tool_debug.log")
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now()}] launch_gui called, creating tk.Tk()...\n")
            f.flush()
    except Exception:
        pass

    root = tk.Tk()
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now()}] tk.Tk() created, initializing KPIBenchmarkApp...\n")
            f.flush()
    except Exception:
        pass

    app = KPIBenchmarkApp(root)
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now()}] KPIBenchmarkApp initialized, entering root.mainloop()...\n")
            f.flush()
    except Exception:
        pass

    # Close splash screen as soon as UI is ready
    try:
        import pyi_splash
        pyi_splash.close()
    except Exception:
        pass

    root.mainloop()
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now()}] root.mainloop() exited cleanly.\n")
            f.flush()
    except Exception:
        pass


if __name__ == "__main__":
    launch_gui()

