#!/usr/bin/env python3

import csv
import json
from datetime import datetime
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk
import re
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates


BASE_DIR = Path(__file__).resolve().parent
FILTER_STATE_PATH = BASE_DIR / "mwo_filter_state.json"


def list_csv_files() -> list[str]:
    """Return CSV filenames that live alongside this script."""
    return sorted(p.name for p in BASE_DIR.glob("*.csv"))


def load_csv_file(
    filename: str, preview_limit: int = 200
) -> tuple[bool, str, list[str], list[list[str]], list[list[str]]]:
    """
    Attempt to load the given CSV.

    Returns:
        success flag, status message, headers, preview rows, all rows.
    """
    csv_path = BASE_DIR / filename
    if not csv_path.exists():
        return False, f"File not found: {filename}", [], [], []

    try:
        def read_csv(encoding: str) -> tuple[list[str], list[list[str]], list[list[str]], int, int]:
            headers: list[str] = []
            preview_rows: list[list[str]] = []
            all_rows: list[list[str]] = []
            data_rows = 0
            corrupted_rows = 0
            with csv_path.open(newline="", encoding=encoding) as handle:
                reader = csv.reader(handle)
                try:
                    headers = next(reader)
                except StopIteration:
                    return [], [], [], 0, 0

                for row in reader:
                    try:
                        if len(row) != len(headers):
                            raise ValueError("column count mismatch")
                        data_rows += 1
                        all_rows.append(row)
                        if len(preview_rows) < preview_limit:
                            preview_rows.append(row)
                    except Exception:
                        corrupted_rows += 1
                        continue
            return headers, preview_rows, all_rows, data_rows, corrupted_rows

        try:
            headers, preview_rows, all_rows, data_rows, corrupted_rows = read_csv("utf-8")
        except UnicodeDecodeError:
            headers, preview_rows, all_rows, data_rows, corrupted_rows = read_csv("cp1252")

        if not headers and not all_rows:
            return False, f"'{filename}' is empty", [], [], []

        total_rows = data_rows + corrupted_rows
        message = f"Total {total_rows}, Success {data_rows}, Fail {corrupted_rows}"
        return True, message, headers, preview_rows, all_rows
    except Exception as exc:  # noqa: BLE001
        return False, f"Failed to load '{filename}': {exc}", [], [], []


def load_filter_state() -> tuple[dict[str, set[str]], str | None, int | None]:
    if not FILTER_STATE_PATH.exists():
        return {}, None, None
    try:
        with FILTER_STATE_PATH.open("r", encoding="utf-8") as handle:
            raw = json.load(handle)
        if not isinstance(raw, dict):
            return {}, None, None
        state: dict[str, set[str]] = {}
        for key in ["apm", "trailer", "work_type", "hw_sw", "icr", "accident"]:
            values = raw.get(key, [])
            if isinstance(values, list):
                state[key] = {str(v) for v in values}
        window_size = raw.get("window_size")
        if isinstance(window_size, str):
            component_sash = raw.get("component_sash")
            if isinstance(component_sash, int):
                return state, window_size, component_sash
            return state, window_size, None
        component_sash = raw.get("component_sash")
        if isinstance(component_sash, int):
            return state, None, component_sash
        return state, None, None
    except Exception:
        return {}, None, None


def save_filter_state(
    state: dict[str, set[str]], window_size: str | None, component_sash: int | None
) -> None:
    payload = {key: sorted(values) for key, values in state.items()}
    if window_size:
        payload["window_size"] = window_size
    if component_sash is not None:
        payload["component_sash"] = component_sash
    with FILTER_STATE_PATH.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)


def build_ui() -> None:
    root = tk.Tk()
    root.title("MWO CSV Loader")
    root.geometry("1100x780")

    status_var = tk.StringVar(value="Select a CSV and click Load")
    selected_file = tk.StringVar()
    current_headers: list[str] = []
    current_rows: list[list[str]] = []
    pivot_row_field = tk.StringVar()
    pivot_col_field = tk.StringVar()
    fleet_period_var = tk.StringVar(value="all")
    chart_period_var = tk.StringVar(value="")
    model_period_var = tk.StringVar(value="all")
    model_chart_period_var = tk.StringVar(value="")
    hw_sw_chart_var = tk.BooleanVar(value=False)
    work_type_chart_var = tk.BooleanVar(value=False)
    icr_chart_var = tk.BooleanVar(value=False)
    accident_chart_var = tk.BooleanVar(value=False)
    keyword_filter_var = tk.StringVar(value="")
    keyword_filter_var_2 = tk.StringVar(value="")
    keyword_status_var = tk.StringVar(value="Load a CSV to view keyword data")
    keyword_fleet_chart_var = tk.BooleanVar(value=True)
    keyword_vehicle_chart_var = tk.BooleanVar(value=False)
    keyword_sort_state = {"column": "", "reverse": False}
    keyword_model_filter_var = tk.StringVar(value="")
    keyword_component_filter_var = tk.StringVar(value="")
    total_count_var = tk.BooleanVar(value=True)
    alert_col_width_var = tk.StringVar(value="300")
    day_col_width_var = tk.StringVar(value="25")
    apm_col_width_var = tk.StringVar(value="40")
    last_prompted_label = tk.StringVar(value="")
    all_apm_ids: list[str] = []
    all_trailer_ids: list[str] = []
    selected_apm_vars: dict[str, tk.BooleanVar] = {}
    selected_trailer_vars: dict[str, tk.BooleanVar] = {}
    all_work_types: list[str] = []
    all_hw_sw_types: list[str] = []
    all_icr_types: list[str] = []
    selected_work_type_vars: dict[str, tk.BooleanVar] = {}
    selected_hw_sw_vars: dict[str, tk.BooleanVar] = {}
    selected_icr_vars: dict[str, tk.BooleanVar] = {}
    saved_filter_state, saved_window_size, saved_component_sash = load_filter_state()
    saved_filters_applied = False
    alert_popup: tk.Toplevel | None = None
    window_size_var = tk.StringVar(value="")
    accident_yes_var = tk.BooleanVar(value=False)
    accident_no_var = tk.BooleanVar(value=False)
    accident_all_var = tk.BooleanVar(value=True)

    if saved_window_size:
        try:
            root.geometry(saved_window_size)
        except Exception:
            pass
    else:
        try:
            root.state("zoomed")
        except Exception:
            pass

    def refresh_dropdown() -> None:
        files = list_csv_files()
        menu = dropdown["menu"]
        menu.delete(0, "end")
        if not files:
            selected_file.set("")
            menu.add_command(label="No CSV files found", command=lambda: None)
            status_var.set("No CSV files found in this folder")
            return

        selected_file.set(files[0])
        for name in files:
            menu.add_command(label=name, command=lambda n=name: selected_file.set(n))
        status_var.set(f"Found {len(files)} CSV file(s)")

    def handle_load() -> None:
        filename = selected_file.get()
        if not filename:
            status_var.set("No CSV selected to load")
            return
        success, message, headers, preview_rows, all_rows = load_csv_file(filename)
        status_var.set(message)
        status_label.configure(foreground="green" if success else "red")
        if success:
            current_headers.clear()
            current_headers.extend(headers)
            current_rows.clear()
            current_rows.extend(all_rows)
            refresh_apm_filter_options()
            refresh_trailer_filter_options()
            refresh_work_type_filter_options()
            refresh_hw_sw_filter_options()
            refresh_icr_filter_options()
            apply_saved_accident_filter()
            saved_filters_applied = True
            load_info_var.set(message + build_unreadable_summary())
            populate_table(headers, preview_rows)
            refresh_pivot_options(headers)
            build_fleet_counts()
            build_model_counts()
            refresh_keyword_view()
        else:
            clear_table()
            clear_pivot_table()
            clear_fleet_table()
            clear_model_table()
            clear_apm_table()
            clear_keyword_table()
            load_info_var.set("")

    def bind_double_q_close(
        window: tk.Misc, close_action: callable, message: str | None = None
    ) -> None:
        last_q_time = {"time": 0.0}

        def handler(_event: tk.Event) -> None:
            now = time.monotonic()
            if now - last_q_time["time"] <= 1.5:
                close_action()
                return
            last_q_time["time"] = now
            if message:
                status_var.set(message)
                status_label.configure(foreground="orange")

        window.bind("q", handler)

    def auto_load_mwo() -> None:
        files = list_csv_files()
        mwo_files = [name for name in files if "mwo" in name.lower()]
        if not mwo_files:
            return
        selected_file.set(mwo_files[0])
        handle_load()

    main_frame = ttk.Frame(root, padding=16)
    main_frame.pack(fill="both", expand=True)

    ttk.Label(main_frame, text="CSV file in this folder:").pack(anchor="w")

    dropdown = tk.OptionMenu(main_frame, selected_file, "")
    dropdown.configure(width=40)
    dropdown.pack(fill="x", pady=(4, 8))

    button_row = ttk.Frame(main_frame)
    button_row.pack(fill="x", pady=(0, 10))
    ttk.Button(button_row, text="Refresh list", command=refresh_dropdown).pack(
        side="left", padx=(0, 8)
    )
    ttk.Button(button_row, text="Load selected CSV", command=handle_load).pack(
        side="left", padx=(0, 8)
    )
    ttk.Label(button_row, text="Window size:").pack(side="left", padx=(8, 4))
    size_entry = ttk.Entry(button_row, textvariable=window_size_var, width=12, state="readonly")
    size_entry.pack(side="left")
    load_info_var = tk.StringVar(value="")
    ttk.Label(button_row, textvariable=load_info_var).pack(side="left", padx=(4, 0))
    ttk.Button(button_row, text="Close", command=root.destroy).pack(side="right")

    status_label = ttk.Label(main_frame, textvariable=status_var, foreground="gray")
    status_label.pack(anchor="w")

    notebook = ttk.Notebook(main_frame)
    notebook.pack(fill="both", expand=True, pady=(8, 0))

    pivot_tab = ttk.Frame(notebook)
    fleet_tab = ttk.Frame(notebook)
    apm_tab = ttk.Frame(notebook)
    keyword_tab = ttk.Frame(notebook)
    notebook.add(pivot_tab, text="Pivot")
    notebook.add(fleet_tab, text="BY COMPONENT")
    notebook.add(apm_tab, text="BY APM")
    notebook.add(keyword_tab, text="BY KEYWORD")
    model_tab = ttk.Frame(notebook)
    notebook.add(model_tab, text="BY MODEL")

    preview_frame = ttk.LabelFrame(pivot_tab, text="Preview (first 200 rows)")
    preview_frame.pack(fill="both", expand=True, pady=(0, 8))

    table_frame = ttk.Frame(preview_frame)
    table_frame.pack(fill="both", expand=True)

    table = ttk.Treeview(table_frame, show="headings")
    y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
    table.configure(yscrollcommand=y_scroll.set)
    table.pack(side="left", fill="both", expand=True)
    y_scroll.pack(side="right", fill="y")

    def clear_table() -> None:
        table.delete(*table.get_children())
        table.configure(columns=[])

    def compute_column_widths(
        headers: list[str],
        rows: list[list[str]],
        min_width: int = 10,
        max_width: int = 260,
    ) -> list[int]:
        """Return pixel widths based on content length, clamped to sensible bounds."""
        longest = [len(h) for h in headers]
        for row in rows:
            for idx, cell in enumerate(row[: len(headers)]):
                text = "" if cell is None else str(cell)
                longest[idx] = max(longest[idx], len(text))
        return [
            max(min_width, min(max_width, length * 8)) for length in longest
        ]  # ~8px per character

    def populate_table(headers: list[str], rows: list[list[str]]) -> None:
        clear_table()
        table.configure(columns=headers)
        widths = compute_column_widths(headers, rows)
        for idx, name in enumerate(headers):
            table.heading(name, text=name)
            table.column(name, width=widths[idx], anchor="w")
        for row in rows:
            values = row + [""] * (len(headers) - len(row))
            table.insert("", "end", values=values[: len(headers)])

    pivot_controls = ttk.LabelFrame(main_frame, text="Pivot table (counts)")
    pivot_controls.pack(in_=pivot_tab, fill="x", pady=(0, 8))

    selector_row = ttk.Frame(pivot_controls)
    selector_row.pack(fill="x", pady=4)

    ttk.Label(selector_row, text="Row field:").pack(side="left")
    row_field_dropdown = tk.OptionMenu(selector_row, pivot_row_field, "")
    row_field_dropdown.configure(width=30)
    row_field_dropdown.pack(side="left", padx=(4, 12))

    ttk.Label(selector_row, text="Column field:").pack(side="left")
    col_field_dropdown = tk.OptionMenu(selector_row, pivot_col_field, "")
    col_field_dropdown.configure(width=30)
    col_field_dropdown.pack(side="left", padx=(4, 12))

    ttk.Button(selector_row, text="Build pivot", command=lambda: build_pivot()).pack(
        side="left"
    )

    pivot_frame = ttk.LabelFrame(pivot_tab, text="Pivot preview (counts)")
    pivot_frame.pack(fill="both", expand=True)

    pivot_table_frame = ttk.Frame(pivot_frame)
    pivot_table_frame.pack(fill="both", expand=True)

    pivot_table = ttk.Treeview(pivot_table_frame, show="headings")
    pivot_y_scroll = ttk.Scrollbar(
        pivot_table_frame, orient="vertical", command=pivot_table.yview
    )
    pivot_table.configure(yscrollcommand=pivot_y_scroll.set)
    pivot_table.pack(side="left", fill="both", expand=True)
    pivot_y_scroll.pack(side="right", fill="y")

    component_body = ttk.Panedwindow(fleet_tab, orient="horizontal")
    component_body.pack(fill="both", expand=True)
    component_left = ttk.Frame(component_body)
    component_right = ttk.Frame(component_body)
    component_body.add(component_left, weight=3)
    component_body.add(component_right, weight=2)

    fleet_info = ttk.Label(
        component_left,
        text=(
            "Work order totals (rows = components, columns = selected time bucket; "
            "uses 'Component' and Start time columns)"
        ),
        anchor="w",
    )
    fleet_info.pack(fill="x", pady=(4, 4), padx=4)

    fleet_buttons = ttk.Frame(component_left)
    fleet_buttons.pack(fill="x", padx=4, pady=(0, 4))
    ttk.Button(
        fleet_buttons, text="All", command=lambda: set_fleet_period("all")
    ).pack(side="left", padx=2)
    ttk.Button(
        fleet_buttons, text="Day", command=lambda: set_fleet_period("day")
    ).pack(side="left", padx=2)
    ttk.Button(
        fleet_buttons, text="Week", command=lambda: set_fleet_period("week")
    ).pack(side="left", padx=2)
    ttk.Button(
        fleet_buttons, text="Month", command=lambda: set_fleet_period("month")
    ).pack(side="left", padx=2)
    ttk.Button(
        fleet_buttons, text="Quarter", command=lambda: set_fleet_period("quarter")
    ).pack(side="left", padx=2)

    apm_trailer_row = ttk.Frame(component_left)
    apm_trailer_row.pack(fill="both", padx=4, pady=(0, 6))

    apm_filter_frame = ttk.LabelFrame(apm_trailer_row, text="Filter by APM")
    apm_filter_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))
    apm_controls = ttk.Frame(apm_filter_frame)
    apm_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(apm_controls, text="Select all", command=lambda: select_all_apm(True)).pack(
        side="left", padx=(0, 4)
    )
    ttk.Button(apm_controls, text="Clear all", command=lambda: select_all_apm(False)).pack(
        side="left"
    )

    apm_list_canvas = tk.Canvas(apm_filter_frame, height=120)
    apm_list_scroll = ttk.Scrollbar(apm_filter_frame, orient="vertical", command=apm_list_canvas.yview)
    apm_list_inner = ttk.Frame(apm_list_canvas)
    apm_list_inner.bind(
        "<Configure>", lambda e: apm_list_canvas.configure(scrollregion=apm_list_canvas.bbox("all"))
    )
    apm_list_canvas.create_window((0, 0), window=apm_list_inner, anchor="nw")
    apm_list_canvas.configure(yscrollcommand=apm_list_scroll.set)
    apm_list_canvas.pack(side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4))
    apm_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    trailer_filter_frame = ttk.LabelFrame(apm_trailer_row, text="Filter by Trailer")
    trailer_filter_frame.pack(side="left", fill="both", expand=True, padx=(4, 0))
    trailer_controls = ttk.Frame(trailer_filter_frame)
    trailer_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(trailer_controls, text="Select all", command=lambda: select_all_trailer(True)).pack(
        side="left", padx=(0, 4)
    )
    ttk.Button(trailer_controls, text="Clear all", command=lambda: select_all_trailer(False)).pack(
        side="left"
    )

    trailer_list_canvas = tk.Canvas(trailer_filter_frame, height=120)
    trailer_list_scroll = ttk.Scrollbar(trailer_filter_frame, orient="vertical", command=trailer_list_canvas.yview)
    trailer_list_inner = ttk.Frame(trailer_list_canvas)
    trailer_list_inner.bind(
        "<Configure>", lambda e: trailer_list_canvas.configure(scrollregion=trailer_list_canvas.bbox("all"))
    )
    trailer_list_canvas.create_window((0, 0), window=trailer_list_inner, anchor="nw")
    trailer_list_canvas.configure(yscrollcommand=trailer_list_scroll.set)
    trailer_list_canvas.pack(side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4))
    trailer_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    accident_filter_frame = ttk.LabelFrame(component_left, text="Accident")
    accident_filter_frame.pack(fill="x", padx=4, pady=(0, 6))
    accident_controls = ttk.Frame(accident_filter_frame)
    accident_controls.pack(fill="x", pady=(4, 4))

    def set_accident_mode(mode: str) -> None:
        accident_all_var.set(mode == "all")
        accident_yes_var.set(mode == "yes")
        accident_no_var.set(mode == "no")
        on_filter_change()

    ttk.Checkbutton(
        accident_controls,
        text="All",
        variable=accident_all_var,
        command=lambda: set_accident_mode("all"),
    ).pack(side="left", padx=(6, 6))
    ttk.Checkbutton(
        accident_controls,
        text="Yes",
        variable=accident_yes_var,
        command=lambda: set_accident_mode("yes"),
    ).pack(side="left", padx=(6, 6))
    ttk.Checkbutton(
        accident_controls,
        text="No",
        variable=accident_no_var,
        command=lambda: set_accident_mode("no"),
    ).pack(side="left", padx=(6, 6))

    work_hw_icr_row = ttk.Frame(component_left)
    work_hw_icr_row.pack(fill="both", padx=4, pady=(0, 6))

    work_type_filter_frame = ttk.LabelFrame(work_hw_icr_row, text="Filter by Work type")
    work_type_filter_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))
    work_type_controls = ttk.Frame(work_type_filter_frame)
    work_type_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(work_type_controls, text="Select all", command=lambda: select_all_values(selected_work_type_vars, True)).pack(
        side="left", padx=(0, 4)
    )
    ttk.Button(work_type_controls, text="Clear all", command=lambda: select_all_values(selected_work_type_vars, False)).pack(
        side="left"
    )

    work_type_list_canvas = tk.Canvas(work_type_filter_frame, height=80)
    work_type_list_scroll = ttk.Scrollbar(work_type_filter_frame, orient="vertical", command=work_type_list_canvas.yview)
    work_type_list_inner = ttk.Frame(work_type_list_canvas)
    work_type_list_inner.bind(
        "<Configure>", lambda e: work_type_list_canvas.configure(scrollregion=work_type_list_canvas.bbox("all"))
    )
    work_type_list_canvas.create_window((0, 0), window=work_type_list_inner, anchor="nw")
    work_type_list_canvas.configure(yscrollcommand=work_type_list_scroll.set)
    work_type_list_canvas.pack(side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4))
    work_type_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    hw_sw_filter_frame = ttk.LabelFrame(work_hw_icr_row, text="Filter by Hardware/Software")
    hw_sw_filter_frame.pack(side="left", fill="both", expand=True, padx=(4, 4))
    hw_sw_controls = ttk.Frame(hw_sw_filter_frame)
    hw_sw_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(hw_sw_controls, text="Select all", command=lambda: select_all_values(selected_hw_sw_vars, True)).pack(
        side="left", padx=(0, 4)
    )
    ttk.Button(hw_sw_controls, text="Clear all", command=lambda: select_all_values(selected_hw_sw_vars, False)).pack(
        side="left"
    )

    hw_sw_list_canvas = tk.Canvas(hw_sw_filter_frame, height=80)
    hw_sw_list_scroll = ttk.Scrollbar(hw_sw_filter_frame, orient="vertical", command=hw_sw_list_canvas.yview)
    hw_sw_list_inner = ttk.Frame(hw_sw_list_canvas)
    hw_sw_list_inner.bind(
        "<Configure>", lambda e: hw_sw_list_canvas.configure(scrollregion=hw_sw_list_canvas.bbox("all"))
    )
    hw_sw_list_canvas.create_window((0, 0), window=hw_sw_list_inner, anchor="nw")
    hw_sw_list_canvas.configure(yscrollcommand=hw_sw_list_scroll.set)
    hw_sw_list_canvas.pack(side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4))
    hw_sw_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    icr_filter_frame = ttk.LabelFrame(work_hw_icr_row, text="Filter by Inspection/Change/Rework")
    icr_filter_frame.pack(side="left", fill="both", expand=True, padx=(4, 0))
    icr_controls = ttk.Frame(icr_filter_frame)
    icr_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(icr_controls, text="Select all", command=lambda: select_all_values(selected_icr_vars, True)).pack(
        side="left", padx=(0, 4)
    )
    ttk.Button(icr_controls, text="Clear all", command=lambda: select_all_values(selected_icr_vars, False)).pack(
        side="left"
    )

    icr_list_canvas = tk.Canvas(icr_filter_frame, height=80)
    icr_list_scroll = ttk.Scrollbar(icr_filter_frame, orient="vertical", command=icr_list_canvas.yview)
    icr_list_inner = ttk.Frame(icr_list_canvas)
    icr_list_inner.bind(
        "<Configure>", lambda e: icr_list_canvas.configure(scrollregion=icr_list_canvas.bbox("all"))
    )
    icr_list_canvas.create_window((0, 0), window=icr_list_inner, anchor="nw")
    icr_list_canvas.configure(yscrollcommand=icr_list_scroll.set)
    icr_list_canvas.pack(side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4))
    icr_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    filter_actions = ttk.Frame(component_left)
    filter_actions.pack(fill="x", padx=4, pady=(0, 6))
    ttk.Button(
        filter_actions, text="Save filters", command=lambda: save_filters()
    ).pack(side="left", padx=(0, 6))
    ttk.Button(
        filter_actions, text="Reset filters", command=lambda: reset_filters()
    ).pack(side="left")

    width_controls = ttk.Frame(component_left)
    width_controls.pack(fill="x", padx=4, pady=(0, 6))
    ttk.Label(width_controls, text="Component col width:").pack(side="left")
    alert_width_entry = ttk.Entry(width_controls, textvariable=alert_col_width_var, width=8)
    alert_width_entry.pack(side="left", padx=(4, 4))
    ttk.Label(width_controls, textvariable=alert_col_width_var).pack(side="left", padx=(0, 10))

    ttk.Label(width_controls, text="Day col width:").pack(side="left")
    day_width_entry = ttk.Entry(width_controls, textvariable=day_col_width_var, width=8)
    day_width_entry.pack(side="left", padx=(4, 4))
    ttk.Label(width_controls, textvariable=day_col_width_var).pack(side="left", padx=(0, 10))

    ttk.Button(width_controls, text="Apply widths", command=lambda: apply_width_settings()).pack(
        side="left", padx=(10, 0)
    )

    fleet_frame = ttk.LabelFrame(
        component_left, text="Work order totals (rows: components, columns: dates/periods)"
    )
    fleet_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))

    fleet_table_frame = ttk.Frame(fleet_frame)
    fleet_table_frame.pack(fill="both", expand=True)

    fleet_table = ttk.Treeview(fleet_table_frame, show="headings")
    fleet_y_scroll = ttk.Scrollbar(
        fleet_table_frame, orient="vertical", command=fleet_table.yview
    )
    fleet_x_scroll = ttk.Scrollbar(
        fleet_table_frame, orient="horizontal", command=fleet_table.xview
    )
    fleet_table.configure(yscrollcommand=fleet_y_scroll.set, xscrollcommand=fleet_x_scroll.set)
    fleet_table.pack(side="top", fill="both", expand=True)
    fleet_y_scroll.pack(side="right", fill="y")
    fleet_x_scroll.pack(side="bottom", fill="x")
    fleet_table.bind("<<TreeviewSelect>>", lambda e: None)


    chart_period_row = ttk.Frame(component_right)
    chart_period_row.pack(fill="x", padx=4, pady=(4, 4))
    ttk.Label(chart_period_row, text="Chart period:").pack(side="left")
    chart_period_dropdown = tk.OptionMenu(chart_period_row, chart_period_var, "")
    chart_period_dropdown.configure(width=20)
    chart_period_dropdown.pack(side="left", padx=(4, 12))
    def set_chart_mode(mode: str) -> None:
        if mode == "total":
            total_count_var.set(True)
            hw_sw_chart_var.set(False)
            work_type_chart_var.set(False)
            icr_chart_var.set(False)
            accident_chart_var.set(False)
            draw_bar_chart()
            return

        total_count_var.set(False)
        if mode == "hw_sw":
            hw_sw_chart_var.set(True)
            work_type_chart_var.set(False)
            icr_chart_var.set(False)
            accident_chart_var.set(False)
        elif mode == "work_type":
            work_type_chart_var.set(True)
            hw_sw_chart_var.set(False)
            icr_chart_var.set(False)
            accident_chart_var.set(False)
        elif mode == "icr":
            icr_chart_var.set(True)
            hw_sw_chart_var.set(False)
            work_type_chart_var.set(False)
            accident_chart_var.set(False)
        elif mode == "accident":
            accident_chart_var.set(True)
            hw_sw_chart_var.set(False)
            work_type_chart_var.set(False)
            icr_chart_var.set(False)

        if not any([hw_sw_chart_var.get(), work_type_chart_var.get(), icr_chart_var.get(), accident_chart_var.get()]):
            total_count_var.set(True)
        draw_bar_chart()

    ttk.Checkbutton(
        chart_period_row,
        text="Hardware / Software",
        variable=hw_sw_chart_var,
        command=lambda: set_chart_mode("hw_sw"),
    ).pack(side="left")
    ttk.Checkbutton(
        chart_period_row,
        text="Work type",
        variable=work_type_chart_var,
        command=lambda: set_chart_mode("work_type"),
    ).pack(side="left", padx=(6, 0))
    ttk.Checkbutton(
        chart_period_row,
        text="Inspection / Change / Rework",
        variable=icr_chart_var,
        command=lambda: set_chart_mode("icr"),
    ).pack(side="left", padx=(6, 0))
    ttk.Checkbutton(
        chart_period_row,
        text="Accident",
        variable=accident_chart_var,
        command=lambda: set_chart_mode("accident"),
    ).pack(side="left", padx=(6, 0))
    ttk.Checkbutton(
        chart_period_row,
        text="Total count",
        variable=total_count_var,
        command=lambda: set_chart_mode("total"),
    ).pack(side="left", padx=(6, 0))

    chart_frame = ttk.LabelFrame(component_right, text="Component bar chart")
    chart_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))
    chart_body = ttk.Frame(chart_frame)
    chart_body.pack(fill="both", expand=True, padx=4, pady=4)

    chart_canvas_frame = ttk.Frame(chart_body)
    chart_canvas_frame.pack(side="left", fill="both", expand=True)

    id_frame = ttk.Frame(chart_body, width=360)
    id_frame.pack(side="right", fill="both", expand=True)
    ttk.Label(id_frame, text="ID / Component").pack(anchor="nw")
    id_table = ttk.Treeview(id_frame, columns=["ID", "Component"], show="headings", height=10)
    id_table.heading("ID", text="ID")
    id_table.heading("Component", text="Component")
    id_table.column("ID", width=50, anchor="w")
    id_table.column("Component", width=180, anchor="w")
    id_table.pack(fill="both", expand=True, pady=(4, 0))

    # BY APM tab
    apm_info = ttk.Label(
        apm_tab,
        text="Component counts by APM / Trailer: rows = components, columns = Vehicle IDs",
        anchor="w",
    )
    apm_info.pack(fill="x", pady=(4, 4), padx=4)

    apm_width_controls = ttk.Frame(apm_tab)
    apm_width_controls.pack(fill="x", padx=4, pady=(0, 6))
    ttk.Label(apm_width_controls, text="Component col width:").pack(side="left")
    apm_alert_width_entry = ttk.Entry(apm_width_controls, textvariable=alert_col_width_var, width=8)
    apm_alert_width_entry.pack(side="left", padx=(4, 4))
    ttk.Label(apm_width_controls, textvariable=alert_col_width_var).pack(side="left", padx=(0, 10))

    ttk.Label(apm_width_controls, text="APM col width:").pack(side="left")
    apm_day_width_entry = ttk.Entry(apm_width_controls, textvariable=apm_col_width_var, width=8)
    apm_day_width_entry.pack(side="left", padx=(4, 4))
    ttk.Label(apm_width_controls, textvariable=apm_col_width_var).pack(side="left", padx=(0, 10))

    ttk.Button(apm_width_controls, text="Apply widths", command=lambda: apply_apm_width_settings()).pack(
        side="left", padx=(10, 0)
    )

    apm_frame = ttk.LabelFrame(apm_tab, text="Components x APM / Trailer")
    apm_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))

    apm_table_frame = ttk.Frame(apm_frame)
    apm_table_frame.pack(fill="both", expand=True)

    apm_table = ttk.Treeview(apm_table_frame, show="headings")
    apm_y_scroll = ttk.Scrollbar(apm_table_frame, orient="vertical", command=apm_table.yview)
    apm_x_scroll = ttk.Scrollbar(apm_table_frame, orient="horizontal", command=apm_table.xview)
    apm_table.configure(yscrollcommand=apm_y_scroll.set, xscrollcommand=apm_x_scroll.set)
    apm_table.pack(side="top", fill="both", expand=True)
    apm_y_scroll.pack(side="right", fill="y")
    apm_x_scroll.pack(side="bottom", fill="x")
    apm_table.bind("<<TreeviewSelect>>", lambda e: None)

    # BY KEYWORD tab
    keyword_body = ttk.Panedwindow(keyword_tab, orient="horizontal")
    keyword_body.pack(fill="both", expand=True)
    keyword_left = ttk.Frame(keyword_body)
    keyword_right = ttk.Frame(keyword_body)
    keyword_body.add(keyword_left, weight=3)
    keyword_body.add(keyword_right, weight=2)

    keyword_controls = ttk.Frame(keyword_left)
    keyword_controls.pack(fill="x", padx=4, pady=(4, 4))
    ttk.Label(
        keyword_controls, text="Keyword search (Work done):"
    ).pack(side="left")
    keyword_entry = ttk.Entry(keyword_controls, textvariable=keyword_filter_var, width=28)
    keyword_entry.pack(side="left", padx=(6, 10))
    keyword_entry.bind("<Return>", lambda _e: refresh_keyword_view())
    ttk.Button(keyword_controls, text="Apply", command=lambda: refresh_keyword_view()).pack(
        side="left"
    )
    ttk.Button(
        keyword_controls,
        text="Clear",
        command=lambda: [
            keyword_filter_var.set(""),
            keyword_filter_var_2.set(""),
            keyword_model_filter_var.set(""),
            keyword_component_filter_var.set(""),
            refresh_keyword_view(),
        ],
    ).pack(side="left", padx=(6, 0))

    keyword_workdone_controls = ttk.Frame(keyword_left)
    keyword_workdone_controls.pack(fill="x", padx=4, pady=(0, 4))
    ttk.Label(keyword_workdone_controls, text="Keyword search (Work done):").pack(
        side="left"
    )
    keyword_entry_2 = ttk.Entry(
        keyword_workdone_controls, textvariable=keyword_filter_var_2, width=28
    )
    keyword_entry_2.pack(side="left", padx=(6, 0))
    keyword_entry_2.bind("<Return>", lambda _e: refresh_keyword_view())

    keyword_secondary_controls = ttk.Frame(keyword_left)
    keyword_secondary_controls.pack(fill="x", padx=4, pady=(0, 4))
    ttk.Label(keyword_secondary_controls, text="Keyword search (Model):").pack(
        side="left"
    )
    keyword_model_entry = ttk.Entry(
        keyword_secondary_controls, textvariable=keyword_model_filter_var, width=24
    )
    keyword_model_entry.pack(side="left", padx=(6, 0))
    keyword_model_entry.bind("<Return>", lambda _e: refresh_keyword_view())

    keyword_component_controls = ttk.Frame(keyword_left)
    keyword_component_controls.pack(fill="x", padx=4, pady=(0, 6))
    ttk.Label(keyword_component_controls, text="Keyword search (Component):").pack(
        side="left"
    )
    keyword_component_entry = ttk.Entry(
        keyword_component_controls, textvariable=keyword_component_filter_var, width=24
    )
    keyword_component_entry.pack(side="left", padx=(6, 0))
    keyword_component_entry.bind("<Return>", lambda _e: refresh_keyword_view())

    keyword_status_label = ttk.Label(keyword_left, textvariable=keyword_status_var)
    keyword_status_label.pack(anchor="w", padx=4, pady=(0, 4))

    keyword_table_frame = ttk.LabelFrame(
        keyword_left, text="Keyword filtered rows (first 300)"
    )
    keyword_table_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))

    keyword_table_body = ttk.Frame(keyword_table_frame)
    keyword_table_body.pack(fill="both", expand=True)

    keyword_table = ttk.Treeview(keyword_table_body, show="headings")
    keyword_y_scroll = ttk.Scrollbar(
        keyword_table_body, orient="vertical", command=keyword_table.yview
    )
    keyword_x_scroll = ttk.Scrollbar(
        keyword_table_body, orient="horizontal", command=keyword_table.xview
    )
    keyword_table.configure(
        yscrollcommand=keyword_y_scroll.set, xscrollcommand=keyword_x_scroll.set
    )
    keyword_table.pack(side="top", fill="both", expand=True)
    keyword_y_scroll.pack(side="right", fill="y")
    keyword_x_scroll.pack(side="bottom", fill="x")

    def set_keyword_chart_scope(mode: str) -> None:
        keyword_fleet_chart_var.set(mode == "fleet")
        keyword_vehicle_chart_var.set(mode == "vehicle")
        refresh_keyword_view()

    keyword_chart_controls = ttk.Frame(keyword_right)
    keyword_chart_controls.pack(fill="x", padx=4, pady=(4, 0))
    ttk.Label(keyword_chart_controls, text="Chart mode:").pack(side="left")
    ttk.Checkbutton(
        keyword_chart_controls,
        text="By fleet",
        variable=keyword_fleet_chart_var,
        command=lambda: set_keyword_chart_scope("fleet"),
    ).pack(side="left", padx=(6, 0))
    ttk.Checkbutton(
        keyword_chart_controls,
        text="By vehicle",
        variable=keyword_vehicle_chart_var,
        command=lambda: set_keyword_chart_scope("vehicle"),
    ).pack(side="left", padx=(6, 0))

    keyword_chart_frame = ttk.LabelFrame(
        keyword_right, text="Component bar chart (keyword filter)"
    )
    keyword_chart_frame.pack(fill="both", expand=True, padx=2, pady=(4, 6))
    keyword_chart_body = ttk.Frame(keyword_chart_frame)
    keyword_chart_body.pack(fill="both", expand=True, padx=4, pady=4)

    # BY MODEL tab
    model_body = ttk.Panedwindow(model_tab, orient="horizontal")
    model_body.pack(fill="both", expand=True)
    model_left = ttk.Frame(model_body)
    model_right = ttk.Frame(model_body)
    model_body.add(model_left, weight=3)
    model_body.add(model_right, weight=2)

    model_info = ttk.Label(
        model_left,
        text=(
            "Model totals (rows = models, columns = selected time bucket; "
            "uses 'Model' and Start time columns)"
        ),
        anchor="w",
    )
    model_info.pack(fill="x", pady=(4, 4), padx=4)

    model_buttons = ttk.Frame(model_left)
    model_buttons.pack(fill="x", padx=4, pady=(0, 4))
    ttk.Button(
        model_buttons, text="All", command=lambda: set_model_period("all")
    ).pack(side="left", padx=2)
    ttk.Button(
        model_buttons, text="Day", command=lambda: set_model_period("day")
    ).pack(side="left", padx=2)
    ttk.Button(
        model_buttons, text="Week", command=lambda: set_model_period("week")
    ).pack(side="left", padx=2)
    ttk.Button(
        model_buttons, text="Month", command=lambda: set_model_period("month")
    ).pack(side="left", padx=2)
    ttk.Button(
        model_buttons, text="Quarter", command=lambda: set_model_period("quarter")
    ).pack(side="left", padx=2)

    model_apm_trailer_row = ttk.Frame(model_left)
    model_apm_trailer_row.pack(fill="both", padx=4, pady=(0, 6))

    model_apm_filter_frame = ttk.LabelFrame(model_apm_trailer_row, text="Filter by APM")
    model_apm_filter_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))
    model_apm_controls = ttk.Frame(model_apm_filter_frame)
    model_apm_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(
        model_apm_controls,
        text="Select all",
        command=lambda: select_all_apm(True),
    ).pack(side="left", padx=(0, 4))
    ttk.Button(
        model_apm_controls, text="Clear all", command=lambda: select_all_apm(False)
    ).pack(side="left")

    model_apm_list_canvas = tk.Canvas(model_apm_filter_frame, height=110)
    model_apm_list_scroll = ttk.Scrollbar(
        model_apm_filter_frame, orient="vertical", command=model_apm_list_canvas.yview
    )
    model_apm_list_inner = ttk.Frame(model_apm_list_canvas)
    model_apm_list_inner.bind(
        "<Configure>",
        lambda e: model_apm_list_canvas.configure(
            scrollregion=model_apm_list_canvas.bbox("all")
        ),
    )
    model_apm_list_canvas.create_window(
        (0, 0), window=model_apm_list_inner, anchor="nw"
    )
    model_apm_list_canvas.configure(yscrollcommand=model_apm_list_scroll.set)
    model_apm_list_canvas.pack(
        side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4)
    )
    model_apm_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    model_trailer_filter_frame = ttk.LabelFrame(model_apm_trailer_row, text="Filter by Trailer")
    model_trailer_filter_frame.pack(side="left", fill="both", expand=True, padx=(4, 0))
    model_trailer_controls = ttk.Frame(model_trailer_filter_frame)
    model_trailer_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(
        model_trailer_controls,
        text="Select all",
        command=lambda: select_all_trailer(True),
    ).pack(side="left", padx=(0, 4))
    ttk.Button(
        model_trailer_controls, text="Clear all", command=lambda: select_all_trailer(False)
    ).pack(side="left")

    model_trailer_list_canvas = tk.Canvas(model_trailer_filter_frame, height=110)
    model_trailer_list_scroll = ttk.Scrollbar(
        model_trailer_filter_frame, orient="vertical", command=model_trailer_list_canvas.yview
    )
    model_trailer_list_inner = ttk.Frame(model_trailer_list_canvas)
    model_trailer_list_inner.bind(
        "<Configure>",
        lambda e: model_trailer_list_canvas.configure(
            scrollregion=model_trailer_list_canvas.bbox("all")
        ),
    )
    model_trailer_list_canvas.create_window(
        (0, 0), window=model_trailer_list_inner, anchor="nw"
    )
    model_trailer_list_canvas.configure(yscrollcommand=model_trailer_list_scroll.set)
    model_trailer_list_canvas.pack(
        side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4)
    )
    model_trailer_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    model_work_hw_icr_row = ttk.Frame(model_left)
    model_work_hw_icr_row.pack(fill="both", padx=4, pady=(0, 6))

    model_work_type_filter_frame = ttk.LabelFrame(model_work_hw_icr_row, text="Filter by Work type")
    model_work_type_filter_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))
    model_work_type_controls = ttk.Frame(model_work_type_filter_frame)
    model_work_type_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(
        model_work_type_controls,
        text="Select all",
        command=lambda: select_all_values(selected_work_type_vars, True),
    ).pack(side="left", padx=(0, 4))
    ttk.Button(
        model_work_type_controls,
        text="Clear all",
        command=lambda: select_all_values(selected_work_type_vars, False),
    ).pack(side="left")
    ttk.Button(
        model_work_type_controls,
        text="Refresh",
        command=lambda: [refresh_work_type_filter_options(), build_fleet_counts(), build_model_counts()],
    ).pack(side="left", padx=(6, 0))

    model_work_type_list_canvas = tk.Canvas(model_work_type_filter_frame, height=80)
    model_work_type_list_scroll = ttk.Scrollbar(
        model_work_type_filter_frame,
        orient="vertical",
        command=model_work_type_list_canvas.yview,
    )
    model_work_type_list_inner = ttk.Frame(model_work_type_list_canvas)
    model_work_type_list_inner.bind(
        "<Configure>",
        lambda e: model_work_type_list_canvas.configure(
            scrollregion=model_work_type_list_canvas.bbox("all")
        ),
    )
    model_work_type_list_canvas.create_window(
        (0, 0), window=model_work_type_list_inner, anchor="nw"
    )
    model_work_type_list_canvas.configure(yscrollcommand=model_work_type_list_scroll.set)
    model_work_type_list_canvas.pack(
        side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4)
    )
    model_work_type_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    model_hw_sw_filter_frame = ttk.LabelFrame(
        model_work_hw_icr_row, text="Filter by Hardware/Software"
    )
    model_hw_sw_filter_frame.pack(side="left", fill="both", expand=True, padx=(4, 4))
    model_hw_sw_controls = ttk.Frame(model_hw_sw_filter_frame)
    model_hw_sw_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(
        model_hw_sw_controls,
        text="Select all",
        command=lambda: select_all_values(selected_hw_sw_vars, True),
    ).pack(side="left", padx=(0, 4))
    ttk.Button(
        model_hw_sw_controls,
        text="Clear all",
        command=lambda: select_all_values(selected_hw_sw_vars, False),
    ).pack(side="left")
    ttk.Button(
        model_hw_sw_controls,
        text="Refresh",
        command=lambda: [refresh_hw_sw_filter_options(), build_fleet_counts(), build_model_counts()],
    ).pack(side="left", padx=(6, 0))

    model_hw_sw_list_canvas = tk.Canvas(model_hw_sw_filter_frame, height=80)
    model_hw_sw_list_scroll = ttk.Scrollbar(
        model_hw_sw_filter_frame, orient="vertical", command=model_hw_sw_list_canvas.yview
    )
    model_hw_sw_list_inner = ttk.Frame(model_hw_sw_list_canvas)
    model_hw_sw_list_inner.bind(
        "<Configure>",
        lambda e: model_hw_sw_list_canvas.configure(
            scrollregion=model_hw_sw_list_canvas.bbox("all")
        ),
    )
    model_hw_sw_list_canvas.create_window(
        (0, 0), window=model_hw_sw_list_inner, anchor="nw"
    )
    model_hw_sw_list_canvas.configure(yscrollcommand=model_hw_sw_list_scroll.set)
    model_hw_sw_list_canvas.pack(
        side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4)
    )
    model_hw_sw_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    model_icr_filter_frame = ttk.LabelFrame(
        model_work_hw_icr_row, text="Filter by Inspection/Change/Rework"
    )
    model_icr_filter_frame.pack(side="left", fill="both", expand=True, padx=(4, 0))
    model_icr_controls = ttk.Frame(model_icr_filter_frame)
    model_icr_controls.pack(fill="x", pady=(4, 2))
    ttk.Button(
        model_icr_controls,
        text="Select all",
        command=lambda: select_all_values(selected_icr_vars, True),
    ).pack(side="left", padx=(0, 4))
    ttk.Button(
        model_icr_controls,
        text="Clear all",
        command=lambda: select_all_values(selected_icr_vars, False),
    ).pack(side="left")
    ttk.Button(
        model_icr_controls,
        text="Refresh",
        command=lambda: [refresh_icr_filter_options(), build_fleet_counts(), build_model_counts()],
    ).pack(side="left", padx=(6, 0))

    model_icr_list_canvas = tk.Canvas(model_icr_filter_frame, height=80)
    model_icr_list_scroll = ttk.Scrollbar(
        model_icr_filter_frame, orient="vertical", command=model_icr_list_canvas.yview
    )
    model_icr_list_inner = ttk.Frame(model_icr_list_canvas)
    model_icr_list_inner.bind(
        "<Configure>",
        lambda e: model_icr_list_canvas.configure(
            scrollregion=model_icr_list_canvas.bbox("all")
        ),
    )
    model_icr_list_canvas.create_window(
        (0, 0), window=model_icr_list_inner, anchor="nw"
    )
    model_icr_list_canvas.configure(yscrollcommand=model_icr_list_scroll.set)
    model_icr_list_canvas.pack(
        side="left", fill="both", expand=True, padx=(0, 4), pady=(0, 4)
    )
    model_icr_list_scroll.pack(side="right", fill="y", pady=(0, 4))

    model_width_controls = ttk.Frame(model_left)
    model_width_controls.pack(fill="x", padx=4, pady=(0, 6))
    ttk.Label(model_width_controls, text="Model col width:").pack(side="left")
    model_alert_width_entry = ttk.Entry(
        model_width_controls, textvariable=alert_col_width_var, width=8
    )
    model_alert_width_entry.pack(side="left", padx=(4, 4))
    ttk.Label(model_width_controls, textvariable=alert_col_width_var).pack(
        side="left", padx=(0, 10)
    )

    ttk.Label(model_width_controls, text="Day col width:").pack(side="left")
    model_day_width_entry = ttk.Entry(
        model_width_controls, textvariable=day_col_width_var, width=8
    )
    model_day_width_entry.pack(side="left", padx=(4, 4))
    ttk.Label(model_width_controls, textvariable=day_col_width_var).pack(
        side="left", padx=(0, 10)
    )

    ttk.Button(
        model_width_controls, text="Apply widths", command=lambda: apply_model_width_settings()
    ).pack(side="left", padx=(10, 0))

    model_frame = ttk.LabelFrame(
        model_left, text="Model totals (rows: models, columns: dates/periods)"
    )
    model_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))

    model_table_frame = ttk.Frame(model_frame)
    model_table_frame.pack(fill="both", expand=True)

    model_table = ttk.Treeview(model_table_frame, show="headings")
    model_y_scroll = ttk.Scrollbar(
        model_table_frame, orient="vertical", command=model_table.yview
    )
    model_x_scroll = ttk.Scrollbar(
        model_table_frame, orient="horizontal", command=model_table.xview
    )
    model_table.configure(yscrollcommand=model_y_scroll.set, xscrollcommand=model_x_scroll.set)
    model_table.pack(side="top", fill="both", expand=True)
    model_y_scroll.pack(side="right", fill="y")
    model_x_scroll.pack(side="bottom", fill="x")
    model_table.bind("<<TreeviewSelect>>", lambda e: None)

    model_chart_period_row = ttk.Frame(model_right)
    model_chart_period_row.pack(fill="x", padx=4, pady=(4, 4))
    ttk.Label(model_chart_period_row, text="Chart period:").pack(side="left")
    model_chart_period_dropdown = tk.OptionMenu(
        model_chart_period_row, model_chart_period_var, ""
    )
    model_chart_period_dropdown.configure(width=20)
    model_chart_period_dropdown.pack(side="left", padx=(4, 12))

    model_chart_frame = ttk.LabelFrame(model_right, text="Model bar chart")
    model_chart_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))
    model_chart_body = ttk.Frame(model_chart_frame)
    model_chart_body.pack(fill="both", expand=True, padx=4, pady=4)

    model_chart_canvas_frame = ttk.Frame(model_chart_body)
    model_chart_canvas_frame.pack(side="left", fill="both", expand=True)

    model_id_frame = ttk.Frame(model_chart_body, width=360)
    model_id_frame.pack(side="right", fill="both", expand=True)
    ttk.Label(model_id_frame, text="ID / Model").pack(anchor="nw")
    model_id_table = ttk.Treeview(
        model_id_frame, columns=["ID", "Model"], show="headings", height=10
    )
    model_id_table.heading("ID", text="ID")
    model_id_table.heading("Model", text="Model")
    model_id_table.column("ID", width=50, anchor="w")
    model_id_table.column("Model", width=180, anchor="w")
    model_id_table.pack(fill="both", expand=True, pady=(4, 0))

    def refresh_pivot_options(headers: list[str]) -> None:
        def _fill(menu: tk.Misc, variable: tk.StringVar) -> None:
            menu.delete(0, "end")
            if not headers:
                variable.set("")
                menu.add_command(label="Load a CSV first", command=lambda: None)
                return
            variable.set(headers[0])
            for name in headers:
                menu.add_command(label=name, command=lambda n=name: variable.set(n))

        _fill(row_field_dropdown["menu"], pivot_row_field)
        _fill(col_field_dropdown["menu"], pivot_col_field)

    def clear_pivot_table() -> None:
        pivot_table.delete(*pivot_table.get_children())
        pivot_table.configure(columns=[])

    def populate_pivot_table(headers: list[str], rows: list[list[str]]) -> None:
        clear_pivot_table()
        pivot_table.configure(columns=headers)
        widths = compute_column_widths(headers, rows, min_width=10, max_width=200)
        for idx, name in enumerate(headers):
            pivot_table.heading(name, text=name)
            pivot_table.column(name, width=widths[idx], anchor="w")
        for row in rows:
            values = row + [""] * (len(headers) - len(row))
            pivot_table.insert("", "end", values=values[: len(headers)])

    def clear_fleet_table() -> None:
        fleet_table.delete(*fleet_table.get_children())
        fleet_table.configure(columns=[])

    def clear_apm_table() -> None:
        apm_table.delete(*apm_table.get_children())
        apm_table.configure(columns=[])

    def populate_fleet_table(headers: list[str], rows: list[list[str]]) -> None:
        clear_fleet_table()
        fleet_table.configure(columns=headers)
        try:
            first_width = int(alert_col_width_var.get())
        except ValueError:
            first_width = 300
            alert_col_width_var.set(str(first_width))
        try:
            day_width = int(day_col_width_var.get())
        except ValueError:
            day_width = 25
            day_col_width_var.set(str(day_width))
        default_width = 120
        meta_labels = {"YEAR", "MONTH", "DAY", "WEEK", "QUARTER"}
        fleet_table.tag_configure("meta", background="#e6e6e6")
        for idx, name in enumerate(headers):
            if idx == 0:
                width = first_width
            elif fleet_period_var.get() == "day":
                width = day_width
            else:
                width = default_width
            fleet_table.heading(name, text=name)
            fleet_table.column(
                name,
                width=width,
                minwidth=width,
                anchor="w",
                stretch=False,  # prevent auto-stretching; keeps horizontal scroll usable
            )
        for row in rows:
            values = row + [""] * (len(headers) - len(row))
            label = str(values[0]) if values else ""
            tags = ("meta",) if label in meta_labels else ()
            fleet_table.insert("", "end", values=values[: len(headers)], tags=tags)

    def clear_model_table() -> None:
        model_table.delete(*model_table.get_children())
        model_table.configure(columns=[])

    def populate_model_table(headers: list[str], rows: list[list[str]]) -> None:
        clear_model_table()
        model_table.configure(columns=headers)
        try:
            first_width = int(alert_col_width_var.get())
        except ValueError:
            first_width = 300
            alert_col_width_var.set(str(first_width))
        try:
            day_width = int(day_col_width_var.get())
        except ValueError:
            day_width = 25
            day_col_width_var.set(str(day_width))
        default_width = 120
        meta_labels = {"YEAR", "MONTH", "DAY", "WEEK", "QUARTER"}
        model_table.tag_configure("meta", background="#e6e6e6")
        for idx, name in enumerate(headers):
            if idx == 0:
                width = first_width
            elif model_period_var.get() == "day":
                width = day_width
            else:
                width = default_width
            model_table.heading(name, text=name)
            model_table.column(
                name,
                width=width,
                minwidth=width,
                anchor="w",
                stretch=False,
            )
        for row in rows:
            values = row + [""] * (len(headers) - len(row))
            label = str(values[0]) if values else ""
            tags = ("meta",) if label in meta_labels else ()
            model_table.insert("", "end", values=values[: len(headers)], tags=tags)

    def populate_apm_table(headers: list[str], rows: list[list[str]]) -> None:
        clear_apm_table()
        apm_table.configure(columns=headers)
        try:
            first_width = int(alert_col_width_var.get())
        except ValueError:
            first_width = 300
            alert_col_width_var.set(str(first_width))
        try:
            apm_width = int(apm_col_width_var.get())
        except ValueError:
            apm_width = 40
            apm_col_width_var.set(str(apm_width))
        default_width = 120
        for idx, name in enumerate(headers):
            if idx == 0:
                width = first_width
            else:
                width = apm_width or default_width
            apm_table.heading(name, text=name)
            apm_table.column(name, width=width, anchor="w", stretch=False, minwidth=width)
        for row in rows:
            values = row + [""] * (len(headers) - len(row))
            apm_table.insert("", "end", values=values[: len(headers)])

    def clear_keyword_table() -> None:
        keyword_table.delete(*keyword_table.get_children())
        keyword_table.configure(columns=[])

    def sort_keyword_table(column: str, force: bool = False) -> None:
        rows = [
            (keyword_table.item(item_id)["values"], item_id)
            for item_id in keyword_table.get_children()
        ]
        if not rows:
            return
        try:
            col_idx = list(keyword_table["columns"]).index(column)
        except ValueError:
            return

        def key_func(values: list[str]) -> tuple[int, object]:
            value = values[col_idx] if col_idx < len(values) else ""
            text = "" if value is None else str(value).strip()
            try:
                number = float(text)
                return (0, number)
            except ValueError:
                return (1, text.lower())

        if force or keyword_sort_state["column"] != column:
            reverse = False
        else:
            reverse = not keyword_sort_state["reverse"]

        rows.sort(key=lambda pair: key_func(pair[0]), reverse=reverse)
        for idx, (_values, item_id) in enumerate(rows):
            keyword_table.move(item_id, "", idx)
        keyword_sort_state["column"] = column
        keyword_sort_state["reverse"] = reverse

    def populate_keyword_table(headers: list[str], rows: list[list[str]]) -> None:
        clear_keyword_table()
        keyword_table.configure(columns=headers)
        widths = compute_column_widths(headers, rows, min_width=60, max_width=220)
        for idx, name in enumerate(headers):
            keyword_table.heading(
                name,
                text=name,
                command=lambda col=name: sort_keyword_table(col),
            )
            keyword_table.column(name, width=widths[idx], anchor="w")
        for row in rows:
            values = row + [""] * (len(headers) - len(row))
            keyword_table.insert("", "end", values=values[: len(headers)])
        if headers:
            sort_keyword_table(headers[0], force=True)

    def populate_id_table(ids_to_alerts: list[tuple[str, str]]) -> None:
        id_table.delete(*id_table.get_children())
        for id_label, alert_name in ids_to_alerts:
            id_table.insert("", "end", values=[id_label, alert_name])

    def populate_model_id_table(ids_to_models: list[tuple[str, str]]) -> None:
        model_id_table.delete(*model_id_table.get_children())
        for id_label, model_name in ids_to_models:
            model_id_table.insert("", "end", values=[id_label, model_name])

    def refresh_chart_period_menu(periods: list[str]) -> None:
        menu = chart_period_dropdown["menu"]
        menu.delete(0, "end")
        if not periods:
            chart_period_var.set("")
            menu.add_command(label="No periods", command=lambda: None)
        else:
            # Default to first period
            chart_period_var.set(periods[0])
            for period in periods:
                menu.add_command(
                    label=period, command=lambda p=period: set_chart_period(p)
                )

    def refresh_model_chart_period_menu(periods: list[str]) -> None:
        menu = model_chart_period_dropdown["menu"]
        menu.delete(0, "end")
        if not periods:
            model_chart_period_var.set("")
            menu.add_command(label="No periods", command=lambda: None)
        else:
            model_chart_period_var.set(periods[0])
            for period in periods:
                menu.add_command(
                    label=period, command=lambda p=period: set_model_chart_period(p)
                )

    def set_chart_period(period: str) -> None:
        chart_period_var.set(period)
        draw_bar_chart()

    def set_model_chart_period(period: str) -> None:
        model_chart_period_var.set(period)
        draw_model_chart()

    def apply_apm_width_settings() -> None:
        """Re-render APM table with current width settings."""
        populate_apm_table(
            list(apm_table["columns"]),
            [apm_table.item(row_id)["values"] for row_id in apm_table.get_children()],
        )

    def apply_width_settings() -> None:
        """Re-render fleet table with current width settings."""
        populate_fleet_table(list(fleet_table["columns"]), [
            fleet_table.item(row_id)["values"] for row_id in fleet_table.get_children()
        ])
        draw_bar_chart()

    def apply_model_width_settings() -> None:
        """Re-render model table with current width settings."""
        populate_model_table(list(model_table["columns"]), [
            model_table.item(row_id)["values"] for row_id in model_table.get_children()
        ])
        draw_model_chart()
    def parse_period_label(label_txt: str) -> datetime | None:
        """Parse day/week/month/quarter labels into a representative date (start of period)."""
        try:
            return datetime.strptime(label_txt, "%Y-%m-%d")
        except ValueError:
            pass
        m = re.match(r"^(\d{4})-W(\d{2})$", label_txt)
        if m:
            year, week = int(m.group(1)), int(m.group(2))
            try:
                return datetime.strptime(f"{year}-W{week}-1", "%G-W%V-%u")
            except ValueError:
                return None
        try:
            return datetime.strptime(label_txt, "%Y-%m")
        except ValueError:
            pass
        m = re.match(r"^(\d{4})-Q(\d)$", label_txt)
        if m:
            year, quarter = int(m.group(1)), int(m.group(2))
            month = (quarter - 1) * 3 + 1
            try:
                return datetime(year, month, 1)
            except ValueError:
                return None
        return None

    def sort_periods(periods: list[str], counts: list[int]) -> tuple[list[str], list[int]]:
        """Sort period labels chronologically when possible, preserving alignment with counts."""
        combined = list(zip(periods, counts))
        def sort_key(label: str) -> tuple[int, int]:
            dt = parse_period_label(label)
            if dt:
                return (0, dt.toordinal())
            return (1, 0)
        combined.sort(key=lambda pair: sort_key(pair[0]))
        sorted_periods, sorted_counts = zip(*combined) if combined else ([], [])
        return list(sorted_periods), list(sorted_counts)

    def sort_period_labels(periods: list[str]) -> list[str]:
        sorted_periods, _ = sort_periods(periods, [0] * len(periods))
        return sorted_periods

    def compute_date_range_for_labels(
        label_names: list[str], label_field: str
    ) -> tuple[str, str] | None:
        """Return min/max dates for selected labels based on the Start time column."""
        date_idx = find_header_index("Start time")
        label_idx = find_header_index(label_field)
        if date_idx is None or label_idx is None:
            return None
        min_dt: datetime | None = None
        max_dt: datetime | None = None
        rows_source = get_filtered_rows()
        for row in rows_source:
            if label_idx < len(row) and row[label_idx] in label_names:
                dt = parse_date(row[date_idx]) if date_idx < len(row) else None
                if dt:
                    if min_dt is None or dt < min_dt:
                        min_dt = dt
                    if max_dt is None or dt > max_dt:
                        max_dt = dt
        if min_dt and max_dt:
            return min_dt.strftime("%Y-%m-%d"), max_dt.strftime("%Y-%m-%d")
        return None

    def aggregate_selected_counts(
        label_names: list[str], label_field: str, period_kind: str
    ) -> tuple[list[str], list[int]]:
        date_idx = find_header_index("Start time")
        label_idx = find_header_index(label_field)
        if label_idx is None:
            return [], []
        if period_kind != "all" and date_idx is None:
            return [], []
        counts_by_period: dict[str, int] = {}
        rows_source = get_filtered_rows()
        for row in rows_source:
            label_val = row[label_idx] if label_idx < len(row) else ""
            if label_val not in label_names:
                continue
            if period_kind == "all":
                period_key = "Sum"
            else:
                date_val = row[date_idx] if date_idx < len(row) else ""
                dt = parse_date(date_val)
                if not dt:
                    continue
                if period_kind == "day":
                    period_key = dt.strftime("%Y-%m-%d")
                elif period_kind == "week":
                    iso_year, iso_week, _ = dt.isocalendar()
                    period_key = f"{iso_year}-W{iso_week:02d}"
                elif period_kind == "month":
                    period_key = dt.strftime("%Y-%m")
                elif period_kind == "quarter":
                    q = (dt.month - 1) // 3 + 1
                    period_key = f"{dt.year}-Q{q}"
                else:
                    period_key = "Sum"
            counts_by_period[period_key] = counts_by_period.get(period_key, 0) + 1

        period_order = list(counts_by_period.keys())
        if period_kind != "all":
            period_order = sort_period_labels(period_order)
        counts = [counts_by_period.get(period, 0) for period in period_order]
        return period_order, counts

    def build_selected_pivot(
        label_names: list[str],
        label_field: str,
        period_kind: str,
        label_kind: str,
    ) -> tuple[list[str], list[list[str]]]:
        label_idx = find_header_index(label_field)
        date_idx = find_header_index("Start time")
        if label_idx is None:
            return [], []
        if period_kind != "all" and date_idx is None:
            return [], []

        label_order: list[str] = []
        period_order: list[str] = []
        totals: dict[str, dict[str, int]] = {}

        def remember(value: str, collection: list[str]) -> None:
            if value not in collection:
                collection.append(value)

        rows_source = get_filtered_rows()
        for row in rows_source:
            label_val = row[label_idx] if label_idx < len(row) else ""
            if label_val not in label_names:
                continue
            remember(label_val, label_order)

            if period_kind == "all":
                period_key = "Sum"
            else:
                date_val = row[date_idx] if date_idx < len(row) else ""
                dt = parse_date(date_val)
                if not dt:
                    continue
                if period_kind == "day":
                    period_key = dt.strftime("%Y-%m-%d")
                elif period_kind == "week":
                    iso_year, iso_week, _ = dt.isocalendar()
                    period_key = f"{iso_year}-W{iso_week:02d}"
                elif period_kind == "month":
                    period_key = dt.strftime("%Y-%m")
                elif period_kind == "quarter":
                    q = (dt.month - 1) // 3 + 1
                    period_key = f"{dt.year}-Q{q}"
                else:
                    period_key = "Sum"
            remember(period_key, period_order)
            totals.setdefault(label_val, {})
            totals[label_val][period_key] = totals[label_val].get(period_key, 0) + 1

        if period_kind != "all":
            period_order = sort_period_labels(period_order)
        if not period_order:
            period_order = ["Sum"]

        headers = [label_kind if label_kind else "Label"] + period_order
        rows: list[list[str]] = []
        for label_val in label_order:
            row_counts = totals.get(label_val, {})
            rows.append(
                [label_val]
                + [str(row_counts.get(period, 0)) for period in period_order]
            )
        return headers, rows

    def show_label_popup(
        label: str,
        periods: list[str],
        counts: list[int],
        period_kind: str,
        label_kind: str,
    ) -> None:
        nonlocal alert_popup
        if alert_popup is not None and alert_popup.winfo_exists():
            alert_popup.destroy()
        if period_kind != "apm":
            periods, counts = sort_periods(periods, counts)
        alert_popup = tk.Toplevel(root)
        alert_popup.title(f"Selected {label_kind.lower()}: {label}")
        try:
            alert_popup.state("zoomed")
        except Exception:
            alert_popup.geometry("900x640")
        container = ttk.Frame(alert_popup, padding=10)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text=f"{label_kind}: {label}").pack(anchor="w")
        # Date range display
        parsed_dates = [parse_period_label(p) for p in periods if parse_period_label(p)]
        range_text: str | None = None
        if parsed_dates:
            start = min(parsed_dates).strftime("%Y-%m-%d")
            end = max(parsed_dates).strftime("%Y-%m-%d")
            range_text = f"Date range: {start} to {end}"
        if range_text:
            ttk.Label(container, text=range_text).pack(anchor="w")

        if not periods or not counts or all(c == 0 for c in counts):
            ttk.Label(
                container,
                text=f"No data to chart for this {label_kind.lower()}.",
            ).pack(anchor="w", pady=(8, 0))
            ttk.Button(container, text="Close", command=alert_popup.destroy).pack(anchor="e", pady=(6, 0))
            return

        fig = Figure(figsize=(7.4, 3.6), dpi=100)
        ax = fig.add_subplot(111)
        fig_canvas: FigureCanvasTkAgg | None = None

        def linear_regression(xs: list[float], ys: list[float]) -> tuple[float, float] | None:
            if len(xs) < 2:
                return None
            n = len(xs)
            sum_x = sum(xs)
            sum_y = sum(ys)
            sum_xx = sum(x * x for x in xs)
            sum_xy = sum(x * y for x, y in zip(xs, ys))
            denom = (n * sum_xx - sum_x * sum_x)
            if denom == 0:
                return None
            slope = (n * sum_xy - sum_x * sum_y) / denom
            intercept = (sum_y - slope * sum_x) / n
            return slope, intercept

        parsed_dates: list[datetime | None] = [parse_period_label(p) for p in periods]
        all_parsed = period_kind != "apm" and all(d is not None for d in parsed_dates)

        if all_parsed and parsed_dates:
            x_vals = [mdates.date2num(dt) for dt in parsed_dates]
            colors = ["#5b8def"] * len(counts)
            if period_kind == "apm":
                # Highlight top 3 counts
                top_idx = sorted(range(len(counts)), key=lambda i: counts[i], reverse=True)[:3]
                for i in top_idx:
                    colors[i] = "#d9534f"
            ax.bar(x_vals, counts, width=5, color=colors, edgecolor="#2f5fb3")
            ax.xaxis_date()
            formatter = mdates.DateFormatter("%Y-%m-%d")
            ax.xaxis.set_major_formatter(formatter)
            fig.autofmt_xdate()
            ax.set_xlabel("Date / Period")
        else:
            x_vals = list(range(len(periods)))
            colors = ["#5b8def"] * len(counts)
            if period_kind == "apm":
                top_idx = sorted(range(len(counts)), key=lambda i: counts[i], reverse=True)[:3]
                for i in top_idx:
                    colors[i] = "#d9534f"
            ax.bar(x_vals, counts, width=0.8, color=colors, edgecolor="#2f5fb3")
            ax.set_xticks(x_vals)
            ax.set_xticklabels(periods, rotation=45, ha="right")
            ax.set_xlabel("APM" if period_kind == "apm" else "Period")

        if period_kind != "apm":
            lr = linear_regression(x_vals, counts)
            if lr:
                slope, intercept = lr
                x_min, x_max = min(x_vals), max(x_vals)
                x_line = [x_min, x_max]
                y_line = [slope * xv + intercept for xv in x_line]
                ax.plot(x_line, y_line, color="#d9534f", linestyle="--", linewidth=1.5, label="Trend (linear)")

            if counts:
                avg = sum(counts) / len(counts)
                ax.axhline(avg, color="#f0ad4e", linestyle="-.", linewidth=1.5, label="Average")

            if lr or counts:
                ax.legend()

        ax.set_ylabel(f"{label_kind} count")
        ax.set_title(label)

        canvas = FigureCanvasTkAgg(fig, master=container)
        fig_canvas = canvas
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, pady=(8, 4))

        def sanitize_filename(text: str) -> str:
            fallback = label_kind.lower() if label_kind else "label"
            return re.sub(r"[^A-Za-z0-9._-]+", "_", text.strip()) or fallback

        def derive_range() -> tuple[str, str]:
            if all_parsed and parsed_dates:
                start = min(parsed_dates).strftime("%Y-%m-%d")
                end = max(parsed_dates).strftime("%Y-%m-%d")
            else:
                start = periods[0] if periods else "start"
                end = periods[-1] if periods else "end"
            return start, end

        def save_chart_and_close() -> None:
            start, end = derive_range()
            fname = f"{sanitize_filename(label)}-fleet-{start}_to_{end}-{period_kind or 'periods'}.png"
            path = BASE_DIR / fname
            fig.savefig(path)
            alert_popup.destroy()


        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(6, 0))
        ttk.Button(button_row, text="Save image and close (S)", command=save_chart_and_close).pack(side="left")
        ttk.Button(button_row, text="Close (Q)", command=alert_popup.destroy).pack(side="right")

        alert_popup.bind("s", lambda e: save_chart_and_close())
        bind_double_q_close(alert_popup, alert_popup.destroy)

    def show_duration_popup(
        label_names: list[str], label_kind: str, label_field: str
    ) -> None:
        nonlocal alert_popup
        if alert_popup is not None and alert_popup.winfo_exists():
            alert_popup.destroy()
        label_display = (
            label_names[0]
            if len(label_names) == 1
            else f"{len(label_names)} {label_kind.lower()}s"
        )
        alert_popup = tk.Toplevel(root)
        alert_popup.title(f"Selected {label_kind.lower()}: {label_display}")
        try:
            alert_popup.state("zoomed")
        except Exception:
            alert_popup.geometry("900x640")

        container = ttk.Frame(alert_popup, padding=10)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text=f"{label_kind}: {label_display}").pack(anchor="w")

        range_var = tk.StringVar(value="")
        range_label = ttk.Label(container, textvariable=range_var)
        range_label.pack(anchor="w")

        show_data_var = tk.BooleanVar(value=False)
        model_mode_var = tk.BooleanVar(value=False)
        pivot_toggle_var = tk.BooleanVar(value=False)
        current_fig: Figure | None = None
        current_period = {"value": "all"}

        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(6, 0))

        mode_row = ttk.Frame(container)
        mode_row.pack(fill="x", pady=(6, 0))
        table_mode_frame = ttk.LabelFrame(mode_row, text="Table mode")
        table_mode_frame.pack(side="left", fill="x", expand=True, padx=(0, 4))
        chart_mode_frame = ttk.LabelFrame(mode_row, text="Chart mode")
        chart_mode_frame.pack(side="left", fill="x", expand=True, padx=(4, 0))

        ttk.Checkbutton(
            table_mode_frame,
            text="Pivot show",
            variable=pivot_toggle_var,
            command=lambda: set_popup_table_mode("pivot"),
        ).pack(side="left", padx=(6, 6))
        ttk.Checkbutton(
            table_mode_frame,
            text="Raw data",
            variable=show_data_var,
            command=lambda: set_popup_table_mode("raw"),
        ).pack(side="left")
        ttk.Checkbutton(
            chart_mode_frame,
            text="Model",
            variable=model_mode_var,
            command=lambda: update_chart(current_period["value"]),
        ).pack(side="left", padx=(6, 6))

        chart_panes = ttk.Panedwindow(container, orient="vertical")
        chart_panes.pack(fill="both", expand=True, pady=(8, 4))
        chart_container = ttk.Frame(chart_panes)
        chart_panes.add(chart_container, weight=3)
        bottom_container = ttk.Frame(chart_panes)
        chart_panes.add(bottom_container, weight=1)
        data_frame = ttk.LabelFrame(bottom_container, text="Filtered data (all columns)")
        data_table_frame = ttk.Frame(data_frame)
        data_table_frame.pack(fill="both", expand=True)
        data_table = ttk.Treeview(data_table_frame, show="headings", height=6)
        data_y_scroll = ttk.Scrollbar(
            data_table_frame, orient="vertical", command=data_table.yview
        )
        data_x_scroll = ttk.Scrollbar(
            data_table_frame, orient="horizontal", command=data_table.xview
        )
        data_table.configure(
            yscrollcommand=data_y_scroll.set, xscrollcommand=data_x_scroll.set
        )
        data_table.pack(side="top", fill="both", expand=True)
        data_y_scroll.pack(side="right", fill="y")
        data_x_scroll.pack(side="bottom", fill="x")

        def clear_data_table() -> None:
            data_table.delete(*data_table.get_children())
            data_table.configure(columns=[])

        def populate_data_table(headers: list[str], rows: list[list[str]]) -> None:
            clear_data_table()
            data_table.configure(columns=headers)
            widths = compute_column_widths(headers, rows, min_width=60, max_width=220)
            for idx, name in enumerate(headers):
                data_table.heading(name, text=name)
                data_table.column(name, width=widths[idx], anchor="w")
            for row in rows:
                values = row + [""] * (len(headers) - len(row))
                data_table.insert("", "end", values=values[: len(headers)])

        def refresh_data_table() -> None:
            if not show_data_var.get():
                return
            if not current_headers:
                clear_data_table()
                return
            label_idx = find_header_index(label_field)
            if label_idx is None:
                clear_data_table()
                return
            rows_source = [
                row for row in get_filtered_rows()
                if label_idx < len(row) and row[label_idx] in label_names
            ]
            populate_data_table(current_headers, rows_source)

        def toggle_data_table() -> None:
            if show_data_var.get():
                data_frame.pack(fill="both", expand=True, pady=(6, 0))
                refresh_data_table()
            else:
                data_frame.pack_forget()

        def set_popup_table_mode(mode: str) -> None:
            if mode == "pivot":
                if pivot_toggle_var.get():
                    show_data_var.set(False)
            elif mode == "raw":
                if show_data_var.get():
                    pivot_toggle_var.set(False)
            if pivot_toggle_var.get():
                pivot_frame.pack(fill="both", expand=False, pady=(6, 0))
                update_pivot(current_period["value"])
            else:
                pivot_frame.pack_forget()
            if show_data_var.get():
                data_frame.pack(fill="both", expand=True, pady=(6, 0))
                refresh_data_table()
            else:
                data_frame.pack_forget()

        pivot_frame = ttk.LabelFrame(bottom_container, text=f"{label_kind} pivot")
        pivot_table_frame = ttk.Frame(pivot_frame)
        pivot_table_frame.pack(fill="both", expand=True)
        pivot_table = ttk.Treeview(pivot_table_frame, show="headings", height=6)
        pivot_y_scroll = ttk.Scrollbar(
            pivot_table_frame, orient="vertical", command=pivot_table.yview
        )
        pivot_x_scroll = ttk.Scrollbar(
            pivot_table_frame, orient="horizontal", command=pivot_table.xview
        )
        pivot_table.configure(
            yscrollcommand=pivot_y_scroll.set, xscrollcommand=pivot_x_scroll.set
        )
        pivot_table.pack(side="top", fill="both", expand=True)
        pivot_y_scroll.pack(side="right", fill="y")
        pivot_x_scroll.pack(side="bottom", fill="x")

        def build_model_pivot(period_kind: str) -> tuple[list[str], list[list[str]]]:
            label_idx = find_header_index(label_field)
            model_idx = find_header_index("Model")
            date_idx = find_header_index("Start time")
            if label_idx is None or model_idx is None:
                return [], []
            if period_kind != "all" and date_idx is None:
                return [], []

            period_order: list[str] = []
            counts: dict[str, dict[str, int]] = {}
            model_totals: dict[str, int] = {}

            def remember(value: str, collection: list[str]) -> None:
                if value not in collection:
                    collection.append(value)

            rows_source = get_filtered_rows()
            for row in rows_source:
                label_val = row[label_idx] if label_idx < len(row) else ""
                if label_val not in label_names:
                    continue
                model_val = row[model_idx] if model_idx < len(row) else ""
                if period_kind == "all":
                    period_key = "Sum"
                else:
                    date_val = row[date_idx] if date_idx < len(row) else ""
                    dt = parse_date(date_val)
                    if not dt:
                        continue
                    if period_kind == "day":
                        period_key = dt.strftime("%Y-%m-%d")
                    elif period_kind == "week":
                        iso_year, iso_week, _ = dt.isocalendar()
                        period_key = f"{iso_year}-W{iso_week:02d}"
                    elif period_kind == "month":
                        period_key = dt.strftime("%Y-%m")
                    elif period_kind == "quarter":
                        q = (dt.month - 1) // 3 + 1
                        period_key = f"{dt.year}-Q{q}"
                    else:
                        period_key = "Sum"
                remember(period_key, period_order)
                counts.setdefault(period_key, {})
                counts[period_key][model_val] = counts[period_key].get(model_val, 0) + 1
                model_totals[model_val] = model_totals.get(model_val, 0) + 1

            if period_kind != "all":
                period_order = sort_period_labels(period_order)
            if not period_order:
                period_order = ["Sum"]

            model_order = sorted(
                model_totals.keys(),
                key=lambda name: model_totals.get(name, 0),
                reverse=True,
            )
            headers = ["Model"] + period_order
            rows: list[list[str]] = []
            for model_name in model_order:
                row = [model_name]
                for period in period_order:
                    row.append(str(counts.get(period, {}).get(model_name, 0)))
                rows.append(row)
            return headers, rows

        def update_pivot(period_kind: str) -> None:
            if model_mode_var.get():
                headers, rows = build_model_pivot(period_kind)
            else:
                headers, rows = build_selected_pivot(
                    label_names,
                    label_field,
                    period_kind,
                    label_kind,
                )
            pivot_table.delete(*pivot_table.get_children())
            pivot_table.configure(columns=headers)
            if headers:
                widths = compute_column_widths(headers, rows, min_width=80, max_width=220)
                for idx, name in enumerate(headers):
                    pivot_table.heading(name, text=name)
                    pivot_table.column(name, width=widths[idx], anchor="w")
            for row in rows:
                values = row + [""] * (len(headers) - len(row))
                pivot_table.insert("", "end", values=values[: len(headers)])

        def toggle_pivot() -> None:
            if pivot_toggle_var.get():
                pivot_frame.pack(fill="both", expand=False, pady=(6, 0))
                update_pivot(current_period["value"])
            else:
                pivot_frame.pack_forget()

        def update_chart(period_kind: str) -> None:
            nonlocal current_fig
            current_period["value"] = period_kind
            for child in chart_container.winfo_children():
                child.destroy()
            computed_range = compute_date_range_for_labels(label_names, label_field)
            if computed_range:
                range_var.set(f"Date range: {computed_range[0]} to {computed_range[1]}")
            else:
                range_var.set("")
            if model_mode_var.get():
                model_idx = find_header_index("Model")
                if model_idx is None:
                    ttk.Label(
                        chart_container,
                        text="Model column missing.",
                    ).pack(anchor="w", pady=(8, 0))
                    return

                counts_by_period: dict[str, dict[str, int]] = {}
                totals_by_model: dict[str, int] = {}
                rows_source = get_filtered_rows()
                date_idx = find_header_index("Start time")
                for row in rows_source:
                    label_idx = find_header_index(label_field)
                    if label_idx is None:
                        continue
                    label_val = row[label_idx] if label_idx < len(row) else ""
                    if label_val not in label_names:
                        continue
                    model_val = row[model_idx] if model_idx < len(row) else ""
                    model_val = model_val.strip() if isinstance(model_val, str) else str(model_val)
                    if not model_val:
                        model_val = "Unclassified"
                    if period_kind == "all":
                        period_key = "Sum"
                    else:
                        if date_idx is None:
                            continue
                        date_val = row[date_idx] if date_idx < len(row) else ""
                        dt = parse_date(date_val)
                        if not dt:
                            continue
                        if period_kind == "day":
                            period_key = dt.strftime("%Y-%m-%d")
                        elif period_kind == "week":
                            iso_year, iso_week, _ = dt.isocalendar()
                            period_key = f"{iso_year}-W{iso_week:02d}"
                        elif period_kind == "month":
                            period_key = dt.strftime("%Y-%m")
                        elif period_kind == "quarter":
                            q = (dt.month - 1) // 3 + 1
                            period_key = f"{dt.year}-Q{q}"
                        else:
                            period_key = "Sum"
                    counts_by_period.setdefault(period_key, {})
                    counts_by_period[period_key][model_val] = (
                        counts_by_period[period_key].get(model_val, 0) + 1
                    )
                    totals_by_model[model_val] = totals_by_model.get(model_val, 0) + 1

                periods = list(counts_by_period.keys())
                if period_kind != "all":
                    periods = sort_period_labels(periods)
                if not periods:
                    ttk.Label(
                        chart_container,
                        text=f"No data to chart for this {label_kind.lower()}.",
                    ).pack(anchor="w", pady=(8, 0))
                    return

                shown_models = sorted(
                    totals_by_model.keys(),
                    key=lambda name: totals_by_model.get(name, 0),
                    reverse=True,
                )

                series: dict[str, list[int]] = {m: [] for m in shown_models}
                for period in periods:
                    model_counts = counts_by_period.get(period, {})
                    for model_name in shown_models:
                        series[model_name].append(model_counts.get(model_name, 0))

                if not any(sum(vals) for vals in series.values()):
                    ttk.Label(
                        chart_container,
                        text=f"No data to chart for this {label_kind.lower()}.",
                    ).pack(anchor="w", pady=(8, 0))
                    return

                fig = Figure(figsize=(7.4, 3.6), dpi=100)
                ax = fig.add_subplot(111)
                parsed_dates: list[datetime | None] = [
                    parse_period_label(p) for p in periods
                ]
                all_parsed = period_kind != "apm" and all(d is not None for d in parsed_dates)
                if all_parsed and parsed_dates:
                    x_vals = [mdates.date2num(dt) for dt in parsed_dates if dt is not None]
                    ax.xaxis_date()
                    formatter = mdates.DateFormatter("%Y-%m-%d")
                    ax.xaxis.set_major_formatter(formatter)
                    fig.autofmt_xdate()
                    ax.set_xlabel("Date / Period")
                else:
                    x_vals = list(range(len(periods)))
                    ax.set_xticks(x_vals)
                    ax.set_xticklabels(periods, rotation=45, ha="right")
                    ax.set_xlabel("Period")

                palette = ["#5b8def", "#f0ad4e", "#5cb85c", "#d9534f", "#9370db", "#9e9e9e", "#5bc0de"]
                bottom_vals = [0] * len(periods)
                for idx, model_name in enumerate(shown_models):
                    vals = series.get(model_name, [0] * len(periods))
                    ax.bar(
                        x_vals,
                        vals,
                        width=0.8 if not all_parsed else 5,
                        bottom=bottom_vals,
                        color=palette[idx % len(palette)],
                        edgecolor="#2f2f2f",
                        label=model_name,
                    )
                    bottom_vals = [b + v for b, v in zip(bottom_vals, vals)]
                ax.legend()
                ax.set_ylabel(f"{label_kind} count")
                ax.set_title(label_display)

                canvas = FigureCanvasTkAgg(fig, master=chart_container)
                canvas.draw()
                canvas.get_tk_widget().pack(fill="both", expand=True)
                current_fig = fig
            else:
                periods, counts = aggregate_selected_counts(
                    label_names, label_field, period_kind
                )
                if not periods or not counts or all(c == 0 for c in counts):
                    ttk.Label(
                        chart_container,
                        text=f"No data to chart for this {label_kind.lower()}.",
                    ).pack(anchor="w", pady=(8, 0))
                    return
                if period_kind != "apm":
                    periods, counts = sort_periods(periods, counts)

                fig = Figure(figsize=(7.4, 3.6), dpi=100)
                ax = fig.add_subplot(111)

                def linear_regression(xs: list[float], ys: list[float]) -> tuple[float, float] | None:
                    if len(xs) < 2:
                        return None
                    n = len(xs)
                    sum_x = sum(xs)
                    sum_y = sum(ys)
                    sum_xx = sum(x * x for x in xs)
                    sum_xy = sum(x * y for x, y in zip(xs, ys))
                    denom = (n * sum_xx - sum_x * sum_x)
                    if denom == 0:
                        return None
                    slope = (n * sum_xy - sum_x * sum_y) / denom
                    intercept = (sum_y - slope * sum_x) / n
                    return slope, intercept

                parsed_dates: list[datetime | None] = [
                    parse_period_label(p) for p in periods
                ]
                all_parsed = period_kind != "apm" and all(d is not None for d in parsed_dates)

                if all_parsed and parsed_dates:
                    x_vals = [mdates.date2num(dt) for dt in parsed_dates if dt is not None]
                    ax.bar(x_vals, counts, width=5, color="#5b8def", edgecolor="#2f5fb3")
                    ax.xaxis_date()
                    formatter = mdates.DateFormatter("%Y-%m-%d")
                    ax.xaxis.set_major_formatter(formatter)
                    fig.autofmt_xdate()
                    ax.set_xlabel("Date / Period")
                else:
                    x_vals = list(range(len(periods)))
                    ax.bar(x_vals, counts, width=0.8, color="#5b8def", edgecolor="#2f5fb3")
                    ax.set_xticks(x_vals)
                    ax.set_xticklabels(periods, rotation=45, ha="right")
                    ax.set_xlabel("Period")

                lr = linear_regression(x_vals, counts)
                if lr:
                    slope, intercept = lr
                    x_min, x_max = min(x_vals), max(x_vals)
                    x_line = [x_min, x_max]
                    y_line = [slope * xv + intercept for xv in x_line]
                    ax.plot(
                        x_line,
                        y_line,
                        color="#d9534f",
                        linestyle="--",
                        linewidth=1.5,
                        label="Trend (linear)",
                    )

                if counts:
                    avg = sum(counts) / len(counts)
                    ax.axhline(
                        avg,
                        color="#f0ad4e",
                        linestyle="-.",
                        linewidth=1.5,
                        label="Average",
                    )

                if lr or counts:
                    ax.legend()

                ax.set_ylabel(f"{label_kind} count")
                ax.set_title(label_display)

                canvas = FigureCanvasTkAgg(fig, master=chart_container)
                canvas.draw()
                canvas.get_tk_widget().pack(fill="both", expand=True)
                current_fig = fig
            if pivot_toggle_var.get():
                update_pivot(period_kind)
            if pivot_toggle_var.get():
                update_pivot(period_kind)
            if show_data_var.get():
                refresh_data_table()

        def save_chart() -> None:
            if current_fig is None:
                return
            fname_label = re.sub(r"[^A-Za-z0-9._-]+", "_", label_display.strip()) or label_kind.lower()
            fname = f"{fname_label}-duration.png"
            path = BASE_DIR / fname
            current_fig.savefig(path)

        ttk.Button(
            button_row, text="All day", command=lambda: update_chart("all")
        ).pack(side="left", padx=2)
        ttk.Button(
            button_row, text="Day", command=lambda: update_chart("day")
        ).pack(side="left", padx=2)
        ttk.Button(
            button_row, text="Week", command=lambda: update_chart("week")
        ).pack(side="left", padx=2)
        ttk.Button(
            button_row, text="Month", command=lambda: update_chart("month")
        ).pack(side="left", padx=2)
        ttk.Button(
            button_row, text="Quota", command=lambda: update_chart("quarter")
        ).pack(side="left", padx=2)
        ttk.Button(
            button_row, text="Save image (S)", command=save_chart
        ).pack(side="left", padx=(10, 0))
        ttk.Button(
            button_row, text="Close (Q)", command=alert_popup.destroy
        ).pack(side="right")

        alert_popup.bind("s", lambda e: save_chart())
        bind_double_q_close(alert_popup, alert_popup.destroy)
        update_chart("all")

    def show_selected_popup(
        table: ttk.Treeview, period_kind: str, label_kind: str, label_field: str
    ) -> None:
        sel = table.selection()
        if not sel:
            status_var.set(f"Select one or more {label_kind.lower()}s, then press 'A'")
            status_label.configure(foreground="red")
            return

        periods = list(table["columns"])[1:]
        if not periods and period_kind == "apm":
            status_var.set("No periods available for chart")
            status_label.configure(foreground="red")
            return

        labels: list[str] = []
        totals = [0 for _ in periods]
        for row_id in sel:
            item = table.item(row_id)
            values = item.get("values", [])
            if not values:
                continue
            label = str(values[0])
            if label in {"YEAR", "MONTH", "DAY", "WEEK", "QUARTER"}:
                continue
            labels.append(label)
            for idx, val in enumerate(values[1 : len(periods) + 1]):
                try:
                    totals[idx] += int(val)
                except (TypeError, ValueError):
                    continue

        if not labels:
            status_var.set(f"No {label_kind.lower()} rows selected")
            status_label.configure(foreground="red")
            return

        if period_kind == "apm":
            if len(labels) == 1:
                label = labels[0]
            else:
                label = f"{len(labels)} {label_kind.lower()}s"
            show_label_popup(label, periods, totals, period_kind, label_kind)
        else:
            show_duration_popup(labels, label_kind, label_field)

    def draw_bar_chart() -> None:
        headers = list(fleet_table["columns"])
        period = chart_period_var.get()
        for child in chart_canvas_frame.winfo_children():
            child.destroy()
        if not headers or not period:
            ttk.Label(
                chart_canvas_frame, text="Load data and choose a chart period."
            ).pack(anchor="nw")
            populate_id_table([])
            return
        try:
            period_idx = headers.index(period)
        except ValueError:
            ttk.Label(
                chart_canvas_frame, text="Selected period not found in table."
            ).pack(anchor="nw")
            populate_id_table([])
            return

        def normalize_hw_sw(value: str) -> str:
            val = (value or "").strip().lower()
            if "hard" in val:
                return "Hardware"
            if "soft" in val:
                return "Software"
            return "Other"

        def normalize_work_type(value: str) -> str:
            val = (value or "").strip().lower()
            if "cmt" in val:
                return "CMT"
            if "pm" in val:
                return "PM"
            if "upgrade" in val:
                return "Upgrade"
            if "internal" in val:
                return "Internal testing"
            if "inspect" in val:
                return "Inspection"
            return "Other"

        def normalize_icr(value: str) -> str:
            val = (value or "").strip().lower()
            if "inspect" in val:
                return "Inspection"
            if "change" in val:
                return "Change"
            if "rework" in val:
                return "Rework"
            return "Other"

        alerts: list[str] = []
        for row_id in fleet_table.get_children():
            row_vals = fleet_table.item(row_id)["values"]
            if len(row_vals) <= period_idx:
                continue
            label = str(row_vals[0])
            if label in {"YEAR", "MONTH", "DAY", "WEEK", "QUARTER"}:
                continue  # skip metadata rows
            alerts.append(label)

        if not alerts:
            ttk.Label(chart_canvas_frame, text="No data to chart.").pack(anchor="nw")
            populate_id_table([])
            return

        counts: list[int] = []
        hw_counts: list[int] = []
        sw_counts: list[int] = []
        other_counts: list[int] = []
        wt_counts: dict[str, list[int]] = {}
        icr_counts: dict[str, list[int]] = {}

        if not total_count_var.get() and (
            hw_sw_chart_var.get()
            or work_type_chart_var.get()
            or icr_chart_var.get()
            or accident_chart_var.get()
        ):
            comp_idx = find_header_index("Component")
            date_idx = find_header_index("Start time")
            hw_idx = find_header_index("Hardware/Software")
            wt_idx = find_header_index("Work type")
            icr_idx = find_header_index("Inspection/Change/Rework")
            acc_idx = find_header_index("Accident")
            if comp_idx is None or date_idx is None:
                ttk.Label(
                    chart_canvas_frame,
                    text="Missing Component/Start time columns.",
                ).pack(anchor="nw")
                populate_id_table([])
                return

            if hw_sw_chart_var.get() and hw_idx is None:
                ttk.Label(chart_canvas_frame, text="Hardware/Software column missing.").pack(anchor="nw")
                populate_id_table([])
                return
            if work_type_chart_var.get() and wt_idx is None:
                ttk.Label(chart_canvas_frame, text="Work type column missing.").pack(anchor="nw")
                populate_id_table([])
                return
            if icr_chart_var.get() and icr_idx is None:
                ttk.Label(chart_canvas_frame, text="Inspection/Change/Rework column missing.").pack(anchor="nw")
                populate_id_table([])
                return
            if accident_chart_var.get() and acc_idx is None:
                ttk.Label(chart_canvas_frame, text="Accident column missing.").pack(anchor="nw")
                populate_id_table([])
                return

            if hw_sw_chart_var.get():
                counts_by_comp: dict[str, dict[str, int]] = {
                    comp: {"Hardware": 0, "Software": 0, "Other": 0} for comp in alerts
                }
            elif work_type_chart_var.get():
                buckets = ["CMT", "PM", "Upgrade", "Internal testing", "Inspection", "Other"]
                counts_by_comp = {comp: {b: 0 for b in buckets} for comp in alerts}
            elif icr_chart_var.get():
                buckets = ["Inspection", "Change", "Rework", "Other"]
                counts_by_comp = {comp: {b: 0 for b in buckets} for comp in alerts}
            else:
                buckets = ["Yes", "No"]
                counts_by_comp = {comp: {b: 0 for b in buckets} for comp in alerts}
            rows_source = get_filtered_rows()

            for row in rows_source:
                comp_val = row[comp_idx] if comp_idx < len(row) else ""
                if comp_val not in counts_by_comp:
                    continue
                if fleet_period_var.get() == "all":
                    period_key = "Sum"
                else:
                    date_val = row[date_idx] if date_idx < len(row) else ""
                    dt = parse_date(date_val)
                    if not dt:
                        continue
                    if fleet_period_var.get() == "day":
                        period_key = dt.strftime("%Y-%m-%d")
                    elif fleet_period_var.get() == "week":
                        iso_year, iso_week, _ = dt.isocalendar()
                        period_key = f"{iso_year}-W{iso_week:02d}"
                    elif fleet_period_var.get() == "month":
                        period_key = dt.strftime("%Y-%m")
                    elif fleet_period_var.get() == "quarter":
                        q = (dt.month - 1) // 3 + 1
                        period_key = f"{dt.year}-Q{q}"
                    else:
                        period_key = "Sum"
                if period_key != period:
                    continue

                if hw_sw_chart_var.get():
                    hw_bucket = normalize_hw_sw(row[hw_idx] if hw_idx < len(row) else "")
                    counts_by_comp[comp_val][hw_bucket] += 1
                elif work_type_chart_var.get():
                    wt_bucket = normalize_work_type(row[wt_idx] if wt_idx < len(row) else "")
                    counts_by_comp[comp_val][wt_bucket] += 1
                elif icr_chart_var.get():
                    icr_bucket = normalize_icr(row[icr_idx] if icr_idx < len(row) else "")
                    counts_by_comp[comp_val][icr_bucket] += 1
                else:
                    acc_val = row[acc_idx] if acc_idx < len(row) else ""
                    acc_val = acc_val.strip().lower() if isinstance(acc_val, str) else str(acc_val).lower()
                    bucket = "Yes" if acc_val in {"yes", "y", "true", "1"} else "No"
                    counts_by_comp[comp_val][bucket] += 1

            if hw_sw_chart_var.get():
                for comp in alerts:
                    hw_counts.append(counts_by_comp[comp]["Hardware"])
                    sw_counts.append(counts_by_comp[comp]["Software"])
                    other_counts.append(counts_by_comp[comp]["Other"])
            elif work_type_chart_var.get():
                for bucket in ["CMT", "PM", "Upgrade", "Internal testing", "Inspection", "Other"]:
                    wt_counts[bucket] = [counts_by_comp[c][bucket] for c in alerts]
            elif icr_chart_var.get():
                for bucket in ["Inspection", "Change", "Rework", "Other"]:
                    icr_counts[bucket] = [counts_by_comp[c][bucket] for c in alerts]
            else:
                icr_counts["Yes"] = [counts_by_comp[c]["Yes"] for c in alerts]
                icr_counts["No"] = [counts_by_comp[c]["No"] for c in alerts]
        else:
            for row_id in fleet_table.get_children():
                row_vals = fleet_table.item(row_id)["values"]
                if len(row_vals) <= period_idx:
                    continue
                label = str(row_vals[0])
                if label in {"YEAR", "MONTH", "DAY", "WEEK", "QUARTER"}:
                    continue
                try:
                    counts.append(int(row_vals[period_idx]))
                except (ValueError, TypeError):
                    counts.append(0)

        if hw_sw_chart_var.get():
            if not any(hw_counts) and not any(sw_counts) and not any(other_counts):
                ttk.Label(chart_canvas_frame, text="No data to chart.").pack(anchor="nw")
                populate_id_table([])
                return
        elif work_type_chart_var.get():
            if not any(sum(vals) for vals in wt_counts.values()):
                ttk.Label(chart_canvas_frame, text="No data to chart.").pack(anchor="nw")
                populate_id_table([])
                return
        elif icr_chart_var.get():
            if not any(sum(vals) for vals in icr_counts.values()):
                ttk.Label(chart_canvas_frame, text="No data to chart.").pack(anchor="nw")
                populate_id_table([])
                return
        elif accident_chart_var.get():
            if not any(sum(vals) for vals in icr_counts.values()):
                ttk.Label(chart_canvas_frame, text="No data to chart.").pack(anchor="nw")
                populate_id_table([])
                return
        else:
            if not counts:
                ttk.Label(chart_canvas_frame, text="No data to chart.").pack(anchor="nw")
                populate_id_table([])
                return

        ids_to_alerts: list[tuple[str, str]] = []
        id_labels: list[str] = []
        for idx, alert_name in enumerate(alerts, start=1):
            id_label = f"A{idx}"
            ids_to_alerts.append((id_label, alert_name))
            id_labels.append(id_label)

        fig = Figure(figsize=(7.2, 3.2), dpi=100)
        ax = fig.add_subplot(111)
        x_vals = list(range(len(alerts)))
        if hw_sw_chart_var.get():
            ax.bar(x_vals, hw_counts, width=0.8, color="#5b8def", edgecolor="#2f5fb3", label="Hardware")
            ax.bar(x_vals, sw_counts, width=0.8, bottom=hw_counts, color="#f0ad4e", edgecolor="#b47a30", label="Software")
            if any(other_counts):
                bottom_vals = [h + s for h, s in zip(hw_counts, sw_counts)]
                ax.bar(x_vals, other_counts, width=0.8, bottom=bottom_vals, color="#9e9e9e", edgecolor="#6f6f6f", label="Other")
            ax.legend()
        elif work_type_chart_var.get():
            colors = {
                "CMT": "#5b8def",
                "PM": "#f0ad4e",
                "Upgrade": "#9370db",
                "Internal testing": "#5cb85c",
                "Inspection": "#d9534f",
                "Other": "#9e9e9e",
            }
            bottom_vals = [0] * len(alerts)
            for bucket in ["CMT", "PM", "Upgrade", "Internal testing", "Inspection", "Other"]:
                vals = wt_counts.get(bucket, [0] * len(alerts))
                ax.bar(
                    x_vals,
                    vals,
                    width=0.8,
                    bottom=bottom_vals,
                    color=colors[bucket],
                    edgecolor="#2f2f2f",
                    label=bucket,
                )
                bottom_vals = [b + v for b, v in zip(bottom_vals, vals)]
            ax.legend()
        elif icr_chart_var.get():
            colors = {
                "Inspection": "#5b8def",
                "Change": "#f0ad4e",
                "Rework": "#5cb85c",
                "Other": "#9e9e9e",
            }
            bottom_vals = [0] * len(alerts)
            for bucket in ["Inspection", "Change", "Rework", "Other"]:
                vals = icr_counts.get(bucket, [0] * len(alerts))
                ax.bar(
                    x_vals,
                    vals,
                    width=0.8,
                    bottom=bottom_vals,
                    color=colors[bucket],
                    edgecolor="#2f2f2f",
                    label=bucket,
                )
                bottom_vals = [b + v for b, v in zip(bottom_vals, vals)]
            ax.legend()
        elif accident_chart_var.get():
            colors = {
                "Yes": "#d9534f",
                "No": "#5cb85c",
            }
            bottom_vals = [0] * len(alerts)
            for bucket in ["Yes", "No"]:
                vals = icr_counts.get(bucket, [0] * len(alerts))
                ax.bar(
                    x_vals,
                    vals,
                    width=0.8,
                    bottom=bottom_vals,
                    color=colors[bucket],
                    edgecolor="#2f2f2f",
                    label=bucket,
                )
                bottom_vals = [b + v for b, v in zip(bottom_vals, vals)]
            ax.legend()
        else:
            ax.bar(x_vals, counts, width=0.8, color="#5b8def", edgecolor="#2f5fb3")
        ax.set_xticks(x_vals)
        ax.set_xticklabels(id_labels, rotation=45, ha="right")
        ax.set_ylabel("Count")
        ax.set_title(f"Period: {period}")

        canvas = FigureCanvasTkAgg(fig, master=chart_canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        id_table.heading("Component", text="Component")
        populate_id_table(ids_to_alerts)

    def draw_model_chart() -> None:
        headers = list(model_table["columns"])
        period = model_chart_period_var.get()
        for child in model_chart_canvas_frame.winfo_children():
            child.destroy()
        if not headers or not period:
            ttk.Label(model_chart_canvas_frame, text="Load data and choose a chart period.").pack(anchor="nw")
            populate_model_id_table([])
            return
        try:
            period_idx = headers.index(period)
        except ValueError:
            ttk.Label(model_chart_canvas_frame, text="Selected period not found in table.").pack(anchor="nw")
            populate_model_id_table([])
            return

        models: list[str] = []
        counts: list[int] = []
        for row_id in model_table.get_children():
            row_vals = model_table.item(row_id)["values"]
            if len(row_vals) <= period_idx:
                continue
            label = str(row_vals[0])
            if label in {"YEAR", "MONTH", "DAY", "WEEK", "QUARTER"}:
                continue
            models.append(label)
            try:
                counts.append(int(row_vals[period_idx]))
            except (ValueError, TypeError):
                counts.append(0)

        if not counts:
            ttk.Label(model_chart_canvas_frame, text="No data to chart.").pack(anchor="nw")
            populate_model_id_table([])
            return

        ids_to_models: list[tuple[str, str]] = []
        id_labels: list[str] = []
        for idx, model_name in enumerate(models, start=1):
            id_label = f"M{idx}"
            ids_to_models.append((id_label, model_name))
            id_labels.append(id_label)

        fig = Figure(figsize=(7.2, 3.2), dpi=100)
        ax = fig.add_subplot(111)
        x_vals = list(range(len(counts)))
        ax.bar(x_vals, counts, width=0.8, color="#5b8def", edgecolor="#2f5fb3")
        ax.set_xticks(x_vals)
        ax.set_xticklabels(id_labels, rotation=0, ha="center")
        ax.set_ylabel("Count")
        ax.set_title(f"Period: {period}")

        canvas = FigureCanvasTkAgg(fig, master=model_chart_canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        populate_model_id_table(ids_to_models)

    def parse_keywords(value: str) -> list[str]:
        text = value.strip()
        if not text:
            return []
        if "," in text:
            parts = [part.strip() for part in text.split(",")]
        else:
            parts = text.split()
        return [part for part in parts if part]

    def filter_rows_by_field_keywords(
        rows: list[list[str]], field_name: str, keywords_text: str
    ) -> tuple[list[list[str]], str | None]:
        tokens = parse_keywords(keywords_text)
        if not tokens:
            return rows, None
        field_idx = find_header_index(field_name)
        if field_idx is None:
            return [], f"Missing '{field_name}' column for keyword search"
        lowered = [token.lower() for token in tokens]
        filtered: list[list[str]] = []
        for row in rows:
            cell = row[field_idx] if field_idx < len(row) else ""
            text = str(cell).lower()
            if any(token in text for token in lowered):
                filtered.append(row)
        return filtered, None

    def refresh_keyword_view() -> None:
        for child in keyword_chart_body.winfo_children():
            child.destroy()
        if not current_headers:
            keyword_status_var.set("Load a CSV to view keyword data")
            clear_keyword_table()
            return

        rows_source = get_filtered_rows()
        rows_filtered = rows_source
        missing_messages: list[str] = []
        rows_filtered, missing = filter_rows_by_field_keywords(
            rows_filtered, "Work done", keyword_filter_var.get()
        )
        if missing:
            missing_messages.append(missing)
        rows_filtered, missing = filter_rows_by_field_keywords(
            rows_filtered, "Work done", keyword_filter_var_2.get()
        )
        if missing:
            missing_messages.append(missing)
        rows_filtered, missing = filter_rows_by_field_keywords(
            rows_filtered, "Model", keyword_model_filter_var.get()
        )
        if missing:
            missing_messages.append(missing)
        rows_filtered, missing = filter_rows_by_field_keywords(
            rows_filtered, "Component", keyword_component_filter_var.get()
        )
        if missing:
            missing_messages.append(missing)
        if missing_messages:
            keyword_status_var.set("; ".join(missing_messages))
        elif (
            keyword_filter_var.get().strip()
            or keyword_filter_var_2.get().strip()
            or keyword_model_filter_var.get().strip()
            or keyword_component_filter_var.get().strip()
        ) and not rows_filtered:
            keyword_status_var.set("No rows matched keyword search")
        preview_limit = 300
        preview_rows = rows_filtered[:preview_limit]
        populate_keyword_table(current_headers, preview_rows)
        if not missing_messages:
            keyword_status_var.set(
                f"Rows: {len(rows_filtered)} (showing {len(preview_rows)})"
            )
        draw_keyword_chart(rows_filtered)


    def draw_keyword_chart(rows_filtered: list[list[str]]) -> None:
        for child in keyword_chart_body.winfo_children():
            child.destroy()
        counts: dict[str, int] = {}
        label_kind = "Component"
        if keyword_vehicle_chart_var.get():
            apm_idx = find_header_index("APM")
            if apm_idx is None:
                ttk.Label(
                    keyword_chart_body, text="APM column missing."
                ).pack(anchor="nw")
                return
            label_kind = "Vehicle"
            all_vehicles: list[str] = []
            for row in current_rows:
                apm_val = row[apm_idx] if apm_idx < len(row) else ""
                apm_val = apm_val.strip() if isinstance(apm_val, str) else str(apm_val)
                if not apm_val or not re.search(r"\d", apm_val):
                    continue
                if apm_val not in all_vehicles:
                    all_vehicles.append(apm_val)
            all_vehicles = sort_filter_values(all_vehicles)
            counts = {vehicle: 0 for vehicle in all_vehicles}
            for row in rows_filtered:
                apm_val = row[apm_idx] if apm_idx < len(row) else ""
                apm_val = apm_val.strip() if isinstance(apm_val, str) else str(apm_val)
                if not apm_val or apm_val not in counts:
                    continue
                counts[apm_val] += 1
            if not counts:
                ttk.Label(
                    keyword_chart_body, text="No vehicles to chart."
                ).pack(anchor="nw")
                return
        else:
            if not rows_filtered:
                ttk.Label(
                    keyword_chart_body, text="No data for keyword filter."
                ).pack(anchor="nw")
                return
            comp_idx = find_header_index("Component")
            if comp_idx is None:
                ttk.Label(
                    keyword_chart_body, text="Component column missing."
                ).pack(anchor="nw")
                return
            for row in rows_filtered:
                comp_val = row[comp_idx] if comp_idx < len(row) else ""
                comp_val = comp_val.strip() if isinstance(comp_val, str) else str(comp_val)
                if not comp_val:
                    comp_val = "Unclassified"
                counts[comp_val] = counts.get(comp_val, 0) + 1

        if not counts:
            ttk.Label(keyword_chart_body, text="No data to chart.").pack(anchor="nw")
            return

        if keyword_vehicle_chart_var.get():
            items = [(vehicle, counts.get(vehicle, 0)) for vehicle in counts.keys()]
            shown = items
        else:
            items = sorted(counts.items(), key=lambda item: item[1], reverse=True)
            max_items = 20
            shown = items[:max_items]
        labels = [item[0] for item in shown]
        values = [item[1] for item in shown]

        fig = Figure(figsize=(7.2, 3.2), dpi=100)
        ax = fig.add_subplot(111)
        x_vals = list(range(len(labels)))
        ax.bar(x_vals, values, width=0.8, color="#5b8def", edgecolor="#2f5fb3")
        ax.set_xticks(x_vals)
        ax.set_xticklabels(labels, rotation=45, ha="right")
        ax.set_ylabel("Count")
        title = f"{label_kind} counts (keyword filter)"
        if not keyword_vehicle_chart_var.get() and len(items) > max_items:
            title += f" - top {max_items} of {len(items)}"
        ax.set_title(title)
        keyword_chart_frame.configure(text=f"{label_kind} bar chart (keyword filter)")

        canvas = FigureCanvasTkAgg(fig, master=keyword_chart_body)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def find_header_index(target: str) -> int | None:
        """Find header index ignoring case and surrounding spaces."""
        normalized = target.strip().lower()
        for idx, name in enumerate(current_headers):
            if name.strip().lower() == normalized:
                return idx
        return None

    def parse_date(value: str) -> datetime | None:
        """Parse common date formats; return None if invalid."""
        if not value:
            return None
        value = value.strip()
        formats = [
            "%m/%d/%Y",
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%m/%d/%Y %H:%M",
            "%m/%d/%Y %H:%M:%S",
            "%m/%d/%Y %I:%M %p",
            "%m/%d/%Y %I:%M:%S %p",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%dT%H:%M:%S",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                continue
        return None

    def build_unreadable_summary() -> str:
        start_idx = find_header_index("Start time")
        if start_idx is None:
            return ""
        total_rows = len(current_rows)
        invalid_rows = 0
        for row in current_rows:
            value = row[start_idx] if start_idx < len(row) else ""
            if not value or parse_date(value) is None:
                invalid_rows += 1
        return f" | Start time unreadable: {invalid_rows}/{total_rows}"

    def set_fleet_period(period: str) -> None:
        fleet_period_var.set(period)
        build_fleet_counts()

    def set_model_period(period: str) -> None:
        model_period_var.set(period)
        build_model_counts()

    def select_all_values(target_vars: dict[str, tk.BooleanVar], state: bool) -> None:
        for var in target_vars.values():
            var.set(state)
        build_fleet_counts()
        build_model_counts()

    def on_filter_change() -> None:
        build_fleet_counts()
        build_model_counts()
        refresh_keyword_view()

    def sort_filter_values(values: list[str]) -> list[str]:
        def sort_key(value: str) -> tuple[int, int, str]:
            digits = re.sub(r"\D", "", value)
            if digits:
                return (0, int(digits), value.lower())
            return (1, 0, value.lower())

        return sorted(values, key=sort_key)

    def current_accident_mode() -> str:
        if accident_yes_var.get():
            return "yes"
        if accident_no_var.get():
            return "no"
        return "all"

    def apply_saved_accident_filter() -> None:
        if saved_filters_applied:
            return
        saved = saved_filter_state.get("accident", set())
        if "yes" in saved:
            set_accident_mode("yes")
        elif "no" in saved:
            set_accident_mode("no")
        elif "all" in saved:
            set_accident_mode("all")

    def save_filters() -> None:
        state = {
            "apm": {val for val, var in selected_apm_vars.items() if var.get()},
            "trailer": {val for val, var in selected_trailer_vars.items() if var.get()},
            "work_type": {val for val, var in selected_work_type_vars.items() if var.get()},
            "hw_sw": {val for val, var in selected_hw_sw_vars.items() if var.get()},
            "icr": {val for val, var in selected_icr_vars.items() if var.get()},
            "accident": {current_accident_mode()},
        }
        size_text = f"{root.winfo_width()}x{root.winfo_height()}"
        try:
            component_sash = component_body.sashpos(0)
        except Exception:
            component_sash = None
        save_filter_state(state, size_text, component_sash)
        nonlocal saved_filter_state
        nonlocal saved_filters_applied
        saved_filter_state = state
        saved_filters_applied = True
        status_var.set("Filter settings saved")
        status_label.configure(foreground="green")

    def reset_filters() -> None:
        for var in selected_apm_vars.values():
            var.set(True)
        for var in selected_trailer_vars.values():
            var.set(True)
        for var in selected_work_type_vars.values():
            var.set(True)
        for var in selected_hw_sw_vars.values():
            var.set(True)
        for var in selected_icr_vars.values():
            var.set(True)
        set_accident_mode("all")
        try:
            FILTER_STATE_PATH.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass
        nonlocal saved_filter_state
        nonlocal saved_filters_applied
        saved_filter_state = {}
        saved_filters_applied = True
        on_filter_change()
        status_var.set("Filters reset")
        status_label.configure(foreground="green")

    def refresh_apm_filter_options() -> None:
        previous_selected = {vid for vid, var in selected_apm_vars.items() if var.get()}
        apm_targets: list[tuple[ttk.Frame, tk.Canvas]] = [
            (apm_list_inner, apm_list_canvas),
            (model_apm_list_inner, model_apm_list_canvas),
        ]
        for target_inner, _ in apm_targets:
            for child in target_inner.winfo_children():
                child.destroy()
        selected_apm_vars.clear()
        all_apm_ids.clear()
        apm_idx = find_header_index("APM")
        if apm_idx is None:
            for target_inner, _ in apm_targets:
                ttk.Label(target_inner, text="No APM column found").pack(anchor="w")
            return

        ids_seen: list[str] = []
        for row in current_rows:
            apm_val = row[apm_idx] if apm_idx < len(row) else ""
            if not apm_val:
                continue
            if apm_val not in ids_seen:
                ids_seen.append(apm_val)

        filtered_ids = ids_seen
        filtered_ids = sort_filter_values(filtered_ids)
        rows_per_col = 5
        use_saved = (not saved_filters_applied) and "apm" in saved_filter_state
        for idx, apm_id in enumerate(filtered_ids):
            all_apm_ids.append(apm_id)
            if use_saved:
                var = tk.BooleanVar(value=apm_id in saved_filter_state["apm"])
            else:
                var = tk.BooleanVar(value=apm_id in previous_selected or not previous_selected)
            selected_apm_vars[apm_id] = var
            for target_inner, _ in apm_targets:
                ttk.Checkbutton(
                    target_inner,
                    text=apm_id,
                    variable=var,
                    command=on_filter_change,
                ).grid(
                    row=idx % rows_per_col,
                    column=idx // rows_per_col,
                    sticky="w",
                    padx=4,
                    pady=2,
                )

        for _, target_canvas in apm_targets:
            target_canvas.yview_moveto(0)

    def refresh_trailer_filter_options() -> None:
        previous_selected = {vid for vid, var in selected_trailer_vars.items() if var.get()}
        trailer_targets: list[tuple[ttk.Frame, tk.Canvas]] = [
            (trailer_list_inner, trailer_list_canvas),
            (model_trailer_list_inner, model_trailer_list_canvas),
        ]
        for target_inner, _ in trailer_targets:
            for child in target_inner.winfo_children():
                child.destroy()
        selected_trailer_vars.clear()
        all_trailer_ids.clear()
        trailer_idx = find_header_index("Trailer")
        if trailer_idx is None:
            for target_inner, _ in trailer_targets:
                ttk.Label(target_inner, text="No Trailer column found").pack(anchor="w")
            return

        ids_seen: list[str] = []
        for row in current_rows:
            trailer_val = row[trailer_idx] if trailer_idx < len(row) else ""
            if not trailer_val:
                continue
            if trailer_val not in ids_seen:
                ids_seen.append(trailer_val)

        filtered_ids = ids_seen
        filtered_ids = sort_filter_values(filtered_ids)
        rows_per_col = 5
        use_saved = (not saved_filters_applied) and "trailer" in saved_filter_state
        for idx, trailer_id in enumerate(filtered_ids):
            all_trailer_ids.append(trailer_id)
            if use_saved:
                var = tk.BooleanVar(value=trailer_id in saved_filter_state["trailer"])
            else:
                var = tk.BooleanVar(value=trailer_id in previous_selected or not previous_selected)
            selected_trailer_vars[trailer_id] = var
            for target_inner, _ in trailer_targets:
                ttk.Checkbutton(
                    target_inner,
                    text=trailer_id,
                    variable=var,
                    command=on_filter_change,
                ).grid(
                    row=idx % rows_per_col,
                    column=idx // rows_per_col,
                    sticky="w",
                    padx=4,
                    pady=2,
                )

        for _, target_canvas in trailer_targets:
            target_canvas.yview_moveto(0)

    def select_all_apm(state: bool) -> None:
        for var in selected_apm_vars.values():
            var.set(state)
        build_fleet_counts()
        build_model_counts()

    def select_all_trailer(state: bool) -> None:
        for var in selected_trailer_vars.values():
            var.set(state)
        build_fleet_counts()
        build_model_counts()

    def refresh_work_type_filter_options() -> None:
        previous_selected = {
            val for val, var in selected_work_type_vars.items() if var.get()
        }
        work_type_targets: list[tuple[ttk.Frame, tk.Canvas]] = [
            (work_type_list_inner, work_type_list_canvas),
            (model_work_type_list_inner, model_work_type_list_canvas),
        ]
        for target_inner, _ in work_type_targets:
            for child in target_inner.winfo_children():
                child.destroy()
        selected_work_type_vars.clear()
        all_work_types.clear()
        field_idx = find_header_index("Work type")
        if field_idx is None:
            for target_inner, _ in work_type_targets:
                ttk.Label(target_inner, text="No Work type column found").pack(anchor="w")
            return

        values_seen: list[str] = []
        for row in current_rows:
            value = row[field_idx] if field_idx < len(row) else ""
            if not value:
                continue
            if value not in values_seen:
                values_seen.append(value)

        filtered_values = values_seen
        use_saved = (not saved_filters_applied) and "work_type" in saved_filter_state
        for idx, value in enumerate(filtered_values):
            all_work_types.append(value)
            if use_saved:
                var = tk.BooleanVar(value=value in saved_filter_state["work_type"])
            else:
                var = tk.BooleanVar(value=value in previous_selected or not previous_selected)
            selected_work_type_vars[value] = var
            for target_inner, _ in work_type_targets:
                ttk.Checkbutton(
                    target_inner,
                    text=value,
                    variable=var,
                    command=on_filter_change,
                ).grid(
                    row=idx // 4, column=idx % 4, sticky="w", padx=4, pady=2
                )
        for _, target_canvas in work_type_targets:
            target_canvas.yview_moveto(0)

    def refresh_hw_sw_filter_options() -> None:
        previous_selected = {
            val for val, var in selected_hw_sw_vars.items() if var.get()
        }
        hw_sw_targets: list[tuple[ttk.Frame, tk.Canvas]] = [
            (hw_sw_list_inner, hw_sw_list_canvas),
            (model_hw_sw_list_inner, model_hw_sw_list_canvas),
        ]
        for target_inner, _ in hw_sw_targets:
            for child in target_inner.winfo_children():
                child.destroy()
        selected_hw_sw_vars.clear()
        all_hw_sw_types.clear()
        field_idx = find_header_index("Hardware/Software")
        if field_idx is None:
            for target_inner, _ in hw_sw_targets:
                ttk.Label(target_inner, text="No Hardware/Software column found").pack(anchor="w")
            return

        values_seen: list[str] = []
        for row in current_rows:
            value = row[field_idx] if field_idx < len(row) else ""
            if not value:
                continue
            if value not in values_seen:
                values_seen.append(value)

        filtered_values = values_seen
        use_saved = (not saved_filters_applied) and "hw_sw" in saved_filter_state
        for idx, value in enumerate(filtered_values):
            all_hw_sw_types.append(value)
            if use_saved:
                var = tk.BooleanVar(value=value in saved_filter_state["hw_sw"])
            else:
                var = tk.BooleanVar(value=value in previous_selected or not previous_selected)
            selected_hw_sw_vars[value] = var
            for target_inner, _ in hw_sw_targets:
                ttk.Checkbutton(
                    target_inner,
                    text=value,
                    variable=var,
                    command=on_filter_change,
                ).grid(
                    row=idx // 4, column=idx % 4, sticky="w", padx=4, pady=2
                )
        for _, target_canvas in hw_sw_targets:
            target_canvas.yview_moveto(0)

    def refresh_icr_filter_options() -> None:
        previous_selected = {
            val for val, var in selected_icr_vars.items() if var.get()
        }
        icr_targets: list[tuple[ttk.Frame, tk.Canvas]] = [
            (icr_list_inner, icr_list_canvas),
            (model_icr_list_inner, model_icr_list_canvas),
        ]
        for target_inner, _ in icr_targets:
            for child in target_inner.winfo_children():
                child.destroy()
        selected_icr_vars.clear()
        all_icr_types.clear()
        field_idx = find_header_index("Inspection/Change/Rework")
        if field_idx is None:
            for target_inner, _ in icr_targets:
                ttk.Label(target_inner, text="No Inspection/Change/Rework column found").pack(anchor="w")
            return

        values_seen: list[str] = []
        for row in current_rows:
            value = row[field_idx] if field_idx < len(row) else ""
            if not value:
                continue
            if value not in values_seen:
                values_seen.append(value)

        filtered_values = values_seen
        use_saved = (not saved_filters_applied) and "icr" in saved_filter_state
        for idx, value in enumerate(filtered_values):
            all_icr_types.append(value)
            if use_saved:
                var = tk.BooleanVar(value=value in saved_filter_state["icr"])
            else:
                var = tk.BooleanVar(value=value in previous_selected or not previous_selected)
            selected_icr_vars[value] = var
            for target_inner, _ in icr_targets:
                ttk.Checkbutton(
                    target_inner,
                    text=value,
                    variable=var,
                    command=on_filter_change,
                ).grid(
                    row=idx // 4, column=idx % 4, sticky="w", padx=4, pady=2
                )
        for _, target_canvas in icr_targets:
            target_canvas.yview_moveto(0)


    def get_filtered_rows() -> list[list[str]]:
        apm_idx = find_header_index("APM")
        trailer_idx = find_header_index("Trailer")
        work_type_idx = find_header_index("Work type")
        hw_sw_idx = find_header_index("Hardware/Software")
        icr_idx = find_header_index("Inspection/Change/Rework")

        selected_apm_ids = [vid for vid, var in selected_apm_vars.items() if var.get()]
        selected_trailer_ids = [vid for vid, var in selected_trailer_vars.items() if var.get()]
        selected_work_types = [val for val, var in selected_work_type_vars.items() if var.get()]
        selected_hw_sw = [val for val, var in selected_hw_sw_vars.items() if var.get()]
        selected_icr = [val for val, var in selected_icr_vars.items() if var.get()]
        accident_mode = current_accident_mode()

        if apm_idx is not None and selected_apm_vars and not selected_apm_ids:
            return []
        if trailer_idx is not None and selected_trailer_vars and not selected_trailer_ids:
            return []
        if work_type_idx is not None and selected_work_type_vars and not selected_work_types:
            return []
        if hw_sw_idx is not None and selected_hw_sw_vars and not selected_hw_sw:
            return []
        if icr_idx is not None and selected_icr_vars and not selected_icr:
            return []

        filtered: list[list[str]] = []
        for row in current_rows:
            if accident_mode != "all":
                acc_idx = find_header_index("Accident")
                acc_val = row[acc_idx] if acc_idx is not None and acc_idx < len(row) else ""
                acc_val = acc_val.strip().lower() if isinstance(acc_val, str) else str(acc_val).lower()
                is_accident = acc_val in {"yes", "y", "true", "1"}
                if accident_mode == "yes" and not is_accident:
                    continue
                if accident_mode == "no" and is_accident:
                    continue

            if apm_idx is not None:
                apm_val = row[apm_idx] if apm_idx < len(row) else ""
                if selected_apm_ids and apm_val not in selected_apm_ids:
                    continue
            if trailer_idx is not None:
                trailer_val = row[trailer_idx] if trailer_idx < len(row) else ""
                if selected_trailer_ids and trailer_val not in selected_trailer_ids:
                    continue

            if work_type_idx is not None:
                work_type_val = row[work_type_idx] if work_type_idx < len(row) else ""
                if selected_work_types and work_type_val not in selected_work_types:
                    continue

            if hw_sw_idx is not None:
                hw_sw_val = row[hw_sw_idx] if hw_sw_idx < len(row) else ""
                if selected_hw_sw and hw_sw_val not in selected_hw_sw:
                    continue

            if icr_idx is not None:
                icr_val = row[icr_idx] if icr_idx < len(row) else ""
                if selected_icr and icr_val not in selected_icr:
                    continue

            filtered.append(row)
        return filtered

    def build_fleet_counts() -> None:
        """Builds a component count table grouped by period."""
        required_fields = ["Component"]
        if not current_headers:
            status_var.set("Load a CSV before building fleet counts")
            status_label.configure(foreground="red")
            clear_fleet_table()
            clear_apm_table()
            return
        missing = [field for field in required_fields if field not in current_headers]
        if missing:
            status_var.set(
                f"Missing columns for fleet counts: {', '.join(missing)}"
            )
            status_label.configure(foreground="red")
            clear_fleet_table()
            clear_apm_table()
            return
        alert_idx = find_header_index("Component")
        date_idx = find_header_index("Start time")
        if alert_idx is None:
            status_var.set("Could not find 'Component' column")
            status_label.configure(foreground="red")
            clear_fleet_table()
            clear_apm_table()
            return
        if date_idx is None and fleet_period_var.get() != "all":
            status_var.set("Start time column not found; cannot group by time")
            status_label.configure(foreground="red")
            clear_fleet_table()
            clear_apm_table()
            return

        alert_order: list[str] = []
        period_order: list[str] = []
        alert_totals: dict[str, dict[str, int]] = {}
        invalid_date_rows = 0

        def remember(value: str, collection: list[str]) -> None:
            if value not in collection:
                collection.append(value)

        rows_source = get_filtered_rows()
        if not rows_source:
            status_var.set("No data after filters")
            status_label.configure(foreground="red")
            clear_fleet_table()
            clear_apm_table()
            return

        for row in rows_source:
            alert_val = row[alert_idx] if alert_idx < len(row) else ""
            remember(alert_val, alert_order)

            if fleet_period_var.get() == "all":
                period_key = "Sum"
            else:
                date_val = row[date_idx] if date_idx < len(row) else ""
                dt = parse_date(date_val)
                if dt:
                    if fleet_period_var.get() == "day":
                        period_key = dt.strftime("%Y-%m-%d")
                    elif fleet_period_var.get() == "week":
                        iso_year, iso_week, _ = dt.isocalendar()
                        period_key = f"{iso_year}-W{iso_week:02d}"
                    elif fleet_period_var.get() == "month":
                        period_key = dt.strftime("%Y-%m")
                    elif fleet_period_var.get() == "quarter":
                        q = (dt.month - 1) // 3 + 1
                        period_key = f"{dt.year}-Q{q}"
                    else:
                        period_key = "Unknown"
                else:
                    invalid_date_rows += 1
                    continue
            remember(period_key, period_order)

            alert_totals.setdefault(alert_val, {})
            alert_totals[alert_val][period_key] = (
                alert_totals[alert_val].get(period_key, 0) + 1
            )

        if not period_order:
            period_order = ["Sum"]
        else:
            period_order = sort_period_labels(period_order)

        headers = ["Component"] + period_order
        fleet_rows: list[list[str]] = []

        def add_meta_rows_from_dates(date_list: list[datetime | None]) -> None:
            years: list[str] = []
            months: list[str] = []
            days: list[str] = []
            weeks: list[str] = []
            quarters: list[str] = []
            for dt in date_list:
                if dt is None:
                    years.append("")
                    months.append("")
                    days.append("")
                    weeks.append("")
                    quarters.append("")
                    continue
                years.append(str(dt.year))
                months.append(f"{dt.month:02d}")
                days.append(f"{dt.day:02d}")
                iso_year, iso_week, _ = dt.isocalendar()
                weeks.append(f"{iso_week:02d}")
                q = (dt.month - 1) // 3 + 1
                quarters.append(str(q))
            fleet_rows.extend(
                [
                    ["YEAR"] + years,
                    ["MONTH"] + months,
                    ["DAY"] + days,
                    ["WEEK"] + weeks,
                    ["QUARTER"] + quarters,
                ]
            )

        if fleet_period_var.get() in {"day", "week", "month", "quarter"}:
            parsed_dates: list[datetime | None] = []
            for period in period_order:
                dt: datetime | None = None
                try:
                    if fleet_period_var.get() == "day":
                        dt = datetime.strptime(period, "%Y-%m-%d")
                    elif fleet_period_var.get() == "week":
                        m = re.match(r"^(\d{4})-W(\d{2})$", period)
                        if m:
                            year, week = int(m.group(1)), int(m.group(2))
                            dt = datetime.strptime(f"{year}-W{week}-1", "%G-W%V-%u")
                    elif fleet_period_var.get() == "month":
                        dt = datetime.strptime(period, "%Y-%m")
                    elif fleet_period_var.get() == "quarter":
                        m = re.match(r"^(\d{4})-Q(\d)$", period)
                        if m:
                            year, quarter = int(m.group(1)), int(m.group(2))
                            month = (quarter - 1) * 3 + 1
                            dt = datetime(year, month, 1)
                except ValueError:
                    dt = None
                parsed_dates.append(dt)
            add_meta_rows_from_dates(parsed_dates)

        def total_for_alert(alert_val: str) -> int:
            counts_for_alert = alert_totals.get(alert_val, {})
            return sum(counts_for_alert.get(period, 0) for period in period_order)

        alert_order.sort(key=total_for_alert, reverse=True)

        for alert_val in alert_order:
            counts_for_alert = alert_totals.get(alert_val, {})
            fleet_rows.append(
                [alert_val] + [str(counts_for_alert.get(period, 0)) for period in period_order]
            )

        populate_fleet_table(headers, fleet_rows)

        # Build APM table
        apm_idx = find_header_index("APM")
        trailer_idx = find_header_index("Trailer")
        if apm_idx is not None or trailer_idx is not None:
            apm_headers = ["Component"]
            apm_rows: list[list[str]] = []
            # Keep consistent vehicle order based on data appearance
            vehicle_order: list[str] = []
            apm_counts: dict[str, dict[str, int]] = {}
            for row in rows_source:
                alert_val = row[alert_idx] if alert_idx < len(row) else ""
                apm_val = row[apm_idx] if apm_idx is not None and apm_idx < len(row) else ""
                trailer_val = row[trailer_idx] if trailer_idx is not None and trailer_idx < len(row) else ""
                vehicle_val = " / ".join([val for val in [apm_val, trailer_val] if val])
                if not vehicle_val:
                    continue
                if vehicle_val not in vehicle_order:
                    vehicle_order.append(vehicle_val)
                apm_counts.setdefault(alert_val, {})
                apm_counts[alert_val][vehicle_val] = apm_counts[alert_val].get(vehicle_val, 0) + 1
            apm_headers += vehicle_order
            for alert_val in alert_order:
                row_counts = apm_counts.get(alert_val, {})
                apm_rows.append(
                    [alert_val] + [str(row_counts.get(v, 0)) for v in vehicle_order]
                )
            populate_apm_table(apm_headers, apm_rows)
        else:
            clear_apm_table()

        refresh_chart_period_menu(period_order)
        draw_bar_chart()
        status_msg = f"Component totals built ({fleet_period_var.get()} grouping)"
        if invalid_date_rows and fleet_period_var.get() != "all":
            status_msg += f"; skipped {invalid_date_rows} invalid Start time row(s)"
        status_var.set(status_msg)
        status_label.configure(foreground="green")

    def build_model_counts() -> None:
        """Builds a model count table grouped by period."""
        required_fields = ["Model"]
        if not current_headers:
            status_var.set("Load a CSV before building model counts")
            status_label.configure(foreground="red")
            clear_model_table()
            return
        missing = [field for field in required_fields if field not in current_headers]
        if missing:
            status_var.set(
                f"Missing columns for model counts: {', '.join(missing)}"
            )
            status_label.configure(foreground="red")
            clear_model_table()
            return
        model_idx = find_header_index("Model")
        date_idx = find_header_index("Start time")
        if model_idx is None:
            status_var.set("Could not find 'Model' column")
            status_label.configure(foreground="red")
            clear_model_table()
            return
        if date_idx is None and model_period_var.get() != "all":
            status_var.set("Start time column not found; cannot group by time")
            status_label.configure(foreground="red")
            clear_model_table()
            return

        model_order: list[str] = []
        period_order: list[str] = []
        model_totals: dict[str, dict[str, int]] = {}
        invalid_date_rows = 0

        def remember(value: str, collection: list[str]) -> None:
            if value not in collection:
                collection.append(value)

        rows_source = get_filtered_rows()
        if not rows_source:
            status_var.set("No data after filters")
            status_label.configure(foreground="red")
            clear_model_table()
            return

        for row in rows_source:
            model_val = row[model_idx] if model_idx < len(row) else ""
            remember(model_val, model_order)

            if model_period_var.get() == "all":
                period_key = "Sum"
            else:
                date_val = row[date_idx] if date_idx < len(row) else ""
                dt = parse_date(date_val)
                if dt:
                    if model_period_var.get() == "day":
                        period_key = dt.strftime("%Y-%m-%d")
                    elif model_period_var.get() == "week":
                        iso_year, iso_week, _ = dt.isocalendar()
                        period_key = f"{iso_year}-W{iso_week:02d}"
                    elif model_period_var.get() == "month":
                        period_key = dt.strftime("%Y-%m")
                    elif model_period_var.get() == "quarter":
                        q = (dt.month - 1) // 3 + 1
                        period_key = f"{dt.year}-Q{q}"
                    else:
                        period_key = "Unknown"
                else:
                    invalid_date_rows += 1
                    continue
            remember(period_key, period_order)

            model_totals.setdefault(model_val, {})
            model_totals[model_val][period_key] = (
                model_totals[model_val].get(period_key, 0) + 1
            )

        if not period_order:
            period_order = ["Sum"]
        else:
            period_order = sort_period_labels(period_order)

        headers = ["Model"] + period_order
        model_rows: list[list[str]] = []

        def add_meta_rows_from_dates(date_list: list[datetime | None]) -> None:
            years: list[str] = []
            months: list[str] = []
            days: list[str] = []
            weeks: list[str] = []
            quarters: list[str] = []
            for dt in date_list:
                if dt is None:
                    years.append("")
                    months.append("")
                    days.append("")
                    weeks.append("")
                    quarters.append("")
                    continue
                years.append(str(dt.year))
                months.append(f"{dt.month:02d}")
                days.append(f"{dt.day:02d}")
                iso_year, iso_week, _ = dt.isocalendar()
                weeks.append(f"{iso_week:02d}")
                q = (dt.month - 1) // 3 + 1
                quarters.append(str(q))
            model_rows.extend(
                [
                    ["YEAR"] + years,
                    ["MONTH"] + months,
                    ["DAY"] + days,
                    ["WEEK"] + weeks,
                    ["QUARTER"] + quarters,
                ]
            )

        if model_period_var.get() in {"day", "week", "month", "quarter"}:
            parsed_dates: list[datetime | None] = []
            for period in period_order:
                dt: datetime | None = None
                try:
                    if model_period_var.get() == "day":
                        dt = datetime.strptime(period, "%Y-%m-%d")
                    elif model_period_var.get() == "week":
                        m = re.match(r"^(\d{4})-W(\d{2})$", period)
                        if m:
                            year, week = int(m.group(1)), int(m.group(2))
                            dt = datetime.strptime(f"{year}-W{week}-1", "%G-W%V-%u")
                    elif model_period_var.get() == "month":
                        dt = datetime.strptime(period, "%Y-%m")
                    elif model_period_var.get() == "quarter":
                        m = re.match(r"^(\d{4})-Q(\d)$", period)
                        if m:
                            year, quarter = int(m.group(1)), int(m.group(2))
                            month = (quarter - 1) * 3 + 1
                            dt = datetime(year, month, 1)
                except ValueError:
                    dt = None
                parsed_dates.append(dt)
            add_meta_rows_from_dates(parsed_dates)

        for model_val in model_order:
            counts_for_model = model_totals.get(model_val, {})
            model_rows.append(
                [model_val] + [str(counts_for_model.get(period, 0)) for period in period_order]
            )

        populate_model_table(headers, model_rows)

        refresh_model_chart_period_menu(period_order)
        draw_model_chart()
        status_msg = f"Model totals built ({model_period_var.get()} grouping)"
        if invalid_date_rows and model_period_var.get() != "all":
            status_msg += f"; skipped {invalid_date_rows} invalid Start time row(s)"
        status_var.set(status_msg)
        status_label.configure(foreground="green")

    def build_pivot() -> None:
        if not current_headers:
            status_var.set("Load a CSV before building a pivot")
            return
        row_field = pivot_row_field.get()
        col_field = pivot_col_field.get()
        if not row_field or not col_field:
            status_var.set("Choose row and column fields for the pivot")
            status_label.configure(foreground="red")
            return
        try:
            row_idx = current_headers.index(row_field)
            col_idx = current_headers.index(col_field)
        except ValueError:
            status_var.set("Selected fields not found; reload the CSV")
            status_label.configure(foreground="red")
            return

        row_order: list[str] = []
        col_order: list[str] = []
        counts: dict[str, dict[str, int]] = {}

        def remember(value: str, collection: list[str]) -> None:
            if value not in collection:
                collection.append(value)

        for row in current_rows:
            row_val = row[row_idx] if row_idx < len(row) else ""
            col_val = row[col_idx] if col_idx < len(row) else ""
            remember(row_val, row_order)
            remember(col_val, col_order)
            counts.setdefault(row_val, {})
            counts[row_val][col_val] = counts[row_val].get(col_val, 0) + 1

        pivot_headers = [row_field] + col_order
        pivot_rows: list[list[str]] = []
        for r_val in row_order:
            row_counts = counts.get(r_val, {})
            pivot_rows.append(
                [r_val] + [str(row_counts.get(c_val, 0)) for c_val in col_order]
            )

        populate_pivot_table(pivot_headers, pivot_rows)
        status_var.set(
            f"Pivot built using '{row_field}' as rows and '{col_field}' as columns"
        )
        status_label.configure(foreground="green")

    refresh_dropdown()
    auto_load_mwo()
    def update_window_size(_event: tk.Event) -> None:
        window_size_var.set(f"{root.winfo_width()}x{root.winfo_height()}")

    def apply_saved_panel_size() -> None:
        if saved_component_sash is None:
            return
        try:
            component_body.sashpos(0, saved_component_sash)
        except Exception:
            pass

    root.bind("<Configure>", update_window_size)
    update_window_size(None)
    apply_saved_panel_size()
    bind_double_q_close(root, root.destroy, "Press Q again to exit")
    fleet_table.bind(
        "a",
        lambda e: show_selected_popup(
            fleet_table, fleet_period_var.get(), "Component", "Component"
        ),
    )
    apm_table.bind(
        "a",
        lambda e: show_selected_popup(apm_table, "apm", "Component", "Component"),
    )
    model_table.bind(
        "a",
        lambda e: show_selected_popup(
            model_table, model_period_var.get(), "Model", "Model"
        ),
    )
    root.mainloop()


if __name__ == "__main__":
    build_ui()
