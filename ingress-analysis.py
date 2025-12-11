#!/usr/bin/env python3

import csv
from datetime import datetime
import tkinter as tk
from pathlib import Path
from tkinter import ttk
import re
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates


BASE_DIR = Path(__file__).resolve().parent


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
        headers: list[str] = []
        preview_rows: list[list[str]] = []
        all_rows: list[list[str]] = []
        data_rows = 0
        corrupted_rows = 0
        with csv_path.open(newline="") as handle:
            reader = csv.reader(handle)
            try:
                headers = next(reader)
            except StopIteration:
                return False, f"'{filename}' is empty", [], [], []

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

        total_rows = data_rows + corrupted_rows
        message = f"Total {total_rows}, Success {data_rows}, Fail {corrupted_rows}"
        return True, message, headers, preview_rows, all_rows
    except Exception as exc:  # noqa: BLE001
        return False, f"Failed to load '{filename}': {exc}", [], [], []


def build_ui() -> None:
    root = tk.Tk()
    root.title("Ingress CSV Loader")
    root.geometry("1100x780")

    status_var = tk.StringVar(value="Select a CSV and click Load")
    selected_file = tk.StringVar()
    current_headers: list[str] = []
    current_rows: list[list[str]] = []
    pivot_row_field = tk.StringVar()
    pivot_col_field = tk.StringVar()
    fleet_period_var = tk.StringVar(value="all")
    chart_period_var = tk.StringVar(value="")
    alert_col_width_var = tk.StringVar(value="300")
    day_col_width_var = tk.StringVar(value="25")
    last_prompted_label = tk.StringVar(value="")
    alert_popup: tk.Toplevel | None = None

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
            load_info_var.set(message)
            populate_table(headers, preview_rows)
            refresh_pivot_options(headers)
            build_fleet_counts()
        else:
            clear_table()
            clear_pivot_table()
            clear_fleet_table()
            load_info_var.set("")

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
    notebook.add(pivot_tab, text="Pivot")
    notebook.add(fleet_tab, text="Fleet")
    notebook.add(apm_tab, text="BY APM")

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

    fleet_info = ttk.Label(
        fleet_tab,
        text=(
            "Fleet alert totals (rows = alert names, columns = selected time bucket; "
            "uses 'RFDS Alert' and Date columns)"
        ),
        anchor="w",
    )
    fleet_info.pack(fill="x", pady=(4, 4), padx=4)

    fleet_buttons = ttk.Frame(fleet_tab)
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

    chart_period_row = ttk.Frame(fleet_tab)
    chart_period_row.pack(fill="x", padx=4, pady=(0, 4))
    ttk.Label(chart_period_row, text="Chart period:").pack(side="left")
    chart_period_dropdown = tk.OptionMenu(chart_period_row, chart_period_var, "")
    chart_period_dropdown.configure(width=20)
    chart_period_dropdown.pack(side="left", padx=(4, 12))

    width_controls = ttk.Frame(fleet_tab)
    width_controls.pack(fill="x", padx=4, pady=(0, 6))
    ttk.Label(width_controls, text="Alert col width:").pack(side="left")
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
        fleet_tab, text="Fleet alert totals (rows: alerts, columns: dates/periods)"
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
    fleet_table.bind("<<TreeviewSelect>>", lambda e: on_fleet_select())

    chart_frame = ttk.LabelFrame(fleet_tab, text="Fleet alert bar chart")
    chart_frame.pack(fill="both", expand=True, padx=2, pady=(0, 6))
    chart_body = ttk.Frame(chart_frame)
    chart_body.pack(fill="both", expand=True, padx=4, pady=4)

    bar_canvas = tk.Canvas(chart_body, height=260, background="white")
    bar_canvas.pack(side="left", fill="both", expand=True)

    id_frame = ttk.Frame(chart_body, width=360)
    id_frame.pack(side="right", fill="both", expand=True)
    ttk.Label(id_frame, text="ID / Alert name").pack(anchor="nw")
    id_table = ttk.Treeview(id_frame, columns=["ID", "Alert"], show="headings", height=10)
    id_table.heading("ID", text="ID")
    id_table.heading("Alert", text="Alert")
    id_table.column("ID", width=50, anchor="w")
    id_table.column("Alert", width=180, anchor="w")
    id_table.pack(fill="both", expand=True, pady=(4, 0))

    # BY APM tab
    apm_info = ttk.Label(
        apm_tab,
        text="Alert counts by APM (Vehicle ID): rows = alerts, columns = Vehicle IDs",
        anchor="w",
    )
    apm_info.pack(fill="x", pady=(4, 4), padx=4)

    apm_frame = ttk.LabelFrame(apm_tab, text="Alerts x APM")
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

    def populate_apm_table(headers: list[str], rows: list[list[str]]) -> None:
        clear_apm_table()
        apm_table.configure(columns=headers)
        widths = compute_column_widths(headers, rows, min_width=60, max_width=240)
        for idx, name in enumerate(headers):
            apm_table.heading(name, text=name)
            apm_table.column(name, width=widths[idx], anchor="w", stretch=False)
        for row in rows:
            values = row + [""] * (len(headers) - len(row))
            apm_table.insert("", "end", values=values[: len(headers)])

    def populate_id_table(ids_to_alerts: list[tuple[str, str]]) -> None:
        id_table.delete(*id_table.get_children())
        for id_label, alert_name in ids_to_alerts:
            id_table.insert("", "end", values=[id_label, alert_name])

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

    def set_chart_period(period: str) -> None:
        chart_period_var.set(period)
        draw_bar_chart()

    def apply_width_settings() -> None:
        """Re-render fleet table with current width settings."""
        populate_fleet_table(list(fleet_table["columns"]), [
            fleet_table.item(row_id)["values"] for row_id in fleet_table.get_children()
        ])
        draw_bar_chart()

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

    def show_alert_popup(label: str, periods: list[str], counts: list[int], period_kind: str) -> None:
        nonlocal alert_popup
        if alert_popup is not None and alert_popup.winfo_exists():
            alert_popup.destroy()
        periods, counts = sort_periods(periods, counts)
        alert_popup = tk.Toplevel(root)
        alert_popup.title(f"Selected alert: {label}")
        try:
            alert_popup.state("zoomed")
        except Exception:
            alert_popup.geometry("900x640")
        container = ttk.Frame(alert_popup, padding=10)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text=f"Alert: {label}").pack(anchor="w")

        if not periods or not counts or all(c == 0 for c in counts):
            ttk.Label(container, text="No data to chart for this alert.").pack(anchor="w", pady=(8, 0))
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
        all_parsed = all(d is not None for d in parsed_dates)

        if all_parsed and parsed_dates:
            x_vals = [mdates.date2num(dt) for dt in parsed_dates]
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
            ax.plot(x_line, y_line, color="#d9534f", linestyle="--", linewidth=1.5, label="Trend (linear)")

        if counts:
            avg = sum(counts) / len(counts)
            ax.axhline(avg, color="#f0ad4e", linestyle="-.", linewidth=1.5, label="Average")

        if lr or counts:
            ax.legend()

        ax.set_ylabel("Alert count")
        ax.set_title(label)

        canvas = FigureCanvasTkAgg(fig, master=container)
        fig_canvas = canvas
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, pady=(8, 4))

        def sanitize_filename(text: str) -> str:
            return re.sub(r"[^A-Za-z0-9._-]+", "_", text.strip()) or "alert"

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
        alert_popup.bind("q", lambda e: alert_popup.destroy())

    def on_fleet_select() -> None:
        sel = fleet_table.selection()
        if not sel:
            return
        item = fleet_table.item(sel[0])
        values = item.get("values", [])
        if not values:
            return
        label = str(values[0])
        if label in {"YEAR", "MONTH", "DAY", "WEEK", "QUARTER"}:
            return
        if label == last_prompted_label.get():
            return
        last_prompted_label.set(label)
        periods = list(fleet_table["columns"])[1:]
        counts: list[int] = []
        for val in values[1 : len(periods) + 1]:
            try:
                counts.append(int(val))
            except (TypeError, ValueError):
                counts.append(0)
        show_alert_popup(label, periods, counts, fleet_period_var.get())

    def draw_bar_chart() -> None:
        headers = list(fleet_table["columns"])
        period = chart_period_var.get()
        bar_canvas.delete("all")
        if not headers or not period:
            bar_canvas.create_text(
                10, 10, anchor="nw", text="Load data and choose a chart period."
            )
            populate_id_table([])
            return
        try:
            period_idx = headers.index(period)
        except ValueError:
            bar_canvas.create_text(
                10, 10, anchor="nw", text="Selected period not found in table."
            )
            populate_id_table([])
            return
        alerts = []
        counts = []
        for row_id in fleet_table.get_children():
            row_vals = fleet_table.item(row_id)["values"]
            if len(row_vals) <= period_idx:
                continue
            label = str(row_vals[0])
            if label in {"YEAR", "MONTH", "DAY", "WEEK", "QUARTER"}:
                continue  # skip metadata rows
            alerts.append(label)
            try:
                counts.append(int(row_vals[period_idx]))
            except (ValueError, TypeError):
                counts.append(0)

        if not counts:
            bar_canvas.create_text(10, 10, anchor="nw", text="No data to chart.")
            populate_id_table([])
            return

        max_val = max(counts)
        width = int(bar_canvas.winfo_width() or 800)
        height = int(bar_canvas.winfo_height() or 260)
        margin = 40
        bar_space = max(1, width - margin * 2)
        n = len(counts)
        bar_width = max(10, min(120, bar_space // max(1, n)))
        gap = max(6, min(20, (bar_space - bar_width * n) // max(1, n)))

        ids_to_alerts: list[tuple[str, str]] = []
        x = margin
        for idx, (label, value) in enumerate(zip(alerts, counts), start=1):
            id_label = f"A{idx}"
            ids_to_alerts.append((id_label, label))
            bar_height = 0 if max_val == 0 else int((value / max_val) * (height - 80))
            y0 = height - margin
            y1 = y0 - bar_height
            bar_canvas.create_rectangle(
                x, y1, x + bar_width, y0, fill="#5b8def", outline="#2f5fb3"
            )
            bar_canvas.create_text(
                x + bar_width / 2, y0 + 12, text=id_label, anchor="n", angle=0
            )
            bar_canvas.create_text(
                x + bar_width / 2, y1 - 10, text=str(value), anchor="s"
            )
            x += bar_width + gap

        populate_id_table(ids_to_alerts)

        # Axes
        bar_canvas.create_line(margin, height - margin, width - margin, height - margin)
        bar_canvas.create_line(margin, height - margin, margin, margin)
        bar_canvas.create_text(
            margin, margin - 10, anchor="w", text=f"Period: {period}"
        )

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
        formats = ["%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"]
        for fmt in formats:
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                continue
        return None

    def set_fleet_period(period: str) -> None:
        fleet_period_var.set(period)
        build_fleet_counts()

    def build_fleet_counts() -> None:
        """Builds a fleet-wide alert count table grouped by period."""
        required_fields = ["RFDS Alert"]
        if not current_headers:
            status_var.set("Load a CSV before building fleet counts")
            status_label.configure(foreground="red")
            clear_fleet_table()
            return
        missing = [field for field in required_fields if field not in current_headers]
        if missing:
            status_var.set(
                f"Missing columns for fleet counts: {', '.join(missing)}"
            )
            status_label.configure(foreground="red")
            clear_fleet_table()
            return
        alert_idx = find_header_index("RFDS Alert")
        date_idx = find_header_index("Date")
        if alert_idx is None:
            status_var.set("Could not find 'RFDS Alert' column")
            status_label.configure(foreground="red")
            clear_fleet_table()
            return
        if date_idx is None and fleet_period_var.get() != "all":
            status_var.set("Date column not found; cannot group by time")
            status_label.configure(foreground="red")
            clear_fleet_table()
            return

        alert_order: list[str] = []
        period_order: list[str] = []
        alert_totals: dict[str, dict[str, int]] = {}

        def remember(value: str, collection: list[str]) -> None:
            if value not in collection:
                collection.append(value)

        for row in current_rows:
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
                    period_key = "Unknown date"
            remember(period_key, period_order)

            alert_totals.setdefault(alert_val, {})
            alert_totals[alert_val][period_key] = (
                alert_totals[alert_val].get(period_key, 0) + 1
            )

        if not period_order:
            period_order = ["Sum"]
        else:
            period_order = sort_period_labels(period_order)

        headers = ["Alert"] + period_order
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

        for alert_val in alert_order:
            counts_for_alert = alert_totals.get(alert_val, {})
            fleet_rows.append(
                [alert_val] + [str(counts_for_alert.get(period, 0)) for period in period_order]
            )

        populate_fleet_table(headers, fleet_rows)
        refresh_chart_period_menu(period_order)
        draw_bar_chart()
        status_var.set(
            f"Fleet alert totals built ({fleet_period_var.get()} grouping)"
        )
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
    root.mainloop()


if __name__ == "__main__":
    build_ui()
