#!/usr/bin/env python3
"""
GUI helper to load all downloaded arbitrator.csv files and pivot alerts.
Looks for ./logs/*/arbitrator.csv and combines them into one table.
"""

from __future__ import annotations

import re
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import messagebox, ttk

import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "logs"
ARBITRATOR_NAME = "arbitrator.csv"
OP_HOUR_FILENAMES = ("op_hours.csv", "op_hour.csv")
DISPLAY_COLS = [
    "alert_name",
    "vehicle_name",
    "vehicle_number",
    "duration",
    "alert_category",
    "alert_severity",
    "start_timestamp",
    "end_timestamp",
    "auto_mode",
    "source_folder",
]

_LAST_FOUR_DIGITS_RE = re.compile(r"(\d{4})(?!.*\d)")


def _extract_vehicle_number(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    match = _LAST_FOUR_DIGITS_RE.search(text)
    if match:
        return match.group(1)
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) >= 4:
        return digits[-4:]
    return digits or text


def add_vehicle_number_column(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "vehicle_name" not in df.columns:
        return df
    out = df.copy()
    out["vehicle_number"] = out["vehicle_name"].map(_extract_vehicle_number)
    return out


def add_parsed_timestamps(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    if "start_timestamp" in out.columns and "_start_dt" not in out.columns:
        out["_start_dt"] = pd.to_datetime(out["start_timestamp"], errors="coerce", utc=True)
    if "end_timestamp" in out.columns and "_end_dt" not in out.columns:
        out["_end_dt"] = pd.to_datetime(out["end_timestamp"], errors="coerce", utc=True)
    return out


def normalize_vehicle_number(series: pd.Series) -> pd.Series:
    out = series.astype(str).fillna("")
    return out.mask(out.str.strip() == "", "UNKNOWN")


def filter_last_days(df: pd.DataFrame, days: int | None) -> tuple[pd.DataFrame, str | None]:
    """
    Filter rows to last N days based on _start_dt/max timestamp in data.
    Returns (filtered_df, warning_message).
    """
    if df.empty or days is None:
        return df, None
    if "_start_dt" not in df.columns:
        return df, "No start_timestamp parsed; cannot filter by days."
    start_dt = df["_start_dt"]
    if start_dt.isna().all():
        return df, "start_timestamp values could not be parsed; cannot filter by days."

    max_dt = start_dt.max()
    min_dt = start_dt.min()
    available_days = max(0, int((max_dt - min_dt).days))
    window_start = max_dt - pd.Timedelta(days=days)
    filtered = df[start_dt >= window_start].copy()

    if available_days < days:
        return filtered, f"Data is only available for the last {available_days} days."
    if filtered.empty:
        return filtered, (
            f"No data in the last {days} days. Data is only available for the last {available_days} days."
        )
    return filtered, None


def describe_time_window_days(df: pd.DataFrame) -> int | None:
    if df.empty or "_start_dt" not in df.columns:
        return None
    start_dt = df["_start_dt"]
    if start_dt.isna().all():
        return None
    min_dt = start_dt.min()
    max_dt = start_dt.max()
    if pd.isna(min_dt) or pd.isna(max_dt):
        return None
    return max(0, int((max_dt - min_dt).days))


def find_arbitrator_files() -> list[Path]:
    """Return paths to arbitrator.csv under logs/*/."""
    if not LOGS_DIR.exists():
        return []
    return sorted(LOGS_DIR.glob(f"*/{ARBITRATOR_NAME}"))


def find_op_hours_files() -> list[Path]:
    """Return all op_hours/op_hour.csv paths under logs/ (any depth), sorted."""
    matches: list[Path] = []
    for name in OP_HOUR_FILENAMES:
        direct = LOGS_DIR / name
        if direct.exists():
            matches.append(direct)
    for name in OP_HOUR_FILENAMES:
        matches.extend(LOGS_DIR.rglob(name))
    # de-dupe
    uniq = sorted(set(matches))
    return uniq


def load_op_hours() -> tuple[pd.DataFrame, str]:
    """Load and combine all op_hours/op_hour.csv files."""
    files = find_op_hours_files()
    if not files:
        return pd.DataFrame(), "op_hours/op_hour.csv not found under logs/"
    frames: list[pd.DataFrame] = []
    for path in files:
        try:
            df = pd.read_csv(path)
            df["source_folder"] = path.parent.name
            df = add_vehicle_number_column(df)
            frames.append(df)
        except Exception as exc:  # noqa: BLE001
            return pd.DataFrame(), f"Failed to load {path}: {exc}"
    combined = pd.concat(frames, ignore_index=True)
    return combined, f"Loaded {len(files)} op_hours file(s), {len(combined)} rows"


def load_all() -> tuple[pd.DataFrame, str]:
    """Load and combine all arbitrator.csv files."""
    frames: list[pd.DataFrame] = []
    files = find_arbitrator_files()
    if not files:
        return pd.DataFrame(), "No arbitrator.csv files found in logs/"
    for path in files:
        try:
            df = pd.read_csv(path)
            df["source_folder"] = path.parent.name
            frames.append(df)
        except Exception as exc:  # noqa: BLE001
            return pd.DataFrame(), f"Failed to load {path}: {exc}"
    combined = pd.concat(frames, ignore_index=True)
    combined["alert_severity"] = combined.get("alert_severity", "").astype(str).str.upper()
    combined = add_vehicle_number_column(combined)
    combined = add_parsed_timestamps(combined)
    return combined, f"Loaded {len(files)} file(s), {len(combined)} rows"


def pivot_dataframe(
    df: pd.DataFrame, row_field: str, col_field: str, value_field: str | None, agg: str
) -> pd.DataFrame:
    """Build a pivot table with minimal guard rails."""
    if df.empty:
        return pd.DataFrame()
    if row_field not in df.columns or col_field not in df.columns:
        return pd.DataFrame()

    if agg == "count":
        pivot = pd.pivot_table(
            df,
            index=row_field,
            columns=col_field,
            aggfunc="size",
            fill_value=0,
        )
    else:
        if value_field is None or value_field not in df.columns:
            return pd.DataFrame()
        pivot = pd.pivot_table(
            df,
            index=row_field,
            columns=col_field,
            values=value_field,
            aggfunc=agg,
            fill_value=0,
        )
    # put row_field back as a column for display
    pivot = pivot.reset_index()
    pivot.columns = [str(c) for c in pivot.columns]
    return pivot


def category_severity_table(df: pd.DataFrame) -> pd.DataFrame:
    """Return counts by alert_category x alert_severity with fixed severity columns."""
    if df.empty:
        return pd.DataFrame()
    if "alert_category" not in df.columns or "alert_severity" not in df.columns:
        return pd.DataFrame()
    ordered = ["INFO", "WARNING", "ERROR", "FATAL"]
    temp = df.copy()
    temp["alert_severity"] = (
        temp["alert_severity"].astype(str).str.upper().astype("category")
    )
    temp["alert_severity"] = temp["alert_severity"].cat.set_categories(
        ordered, ordered=True
    )
    pivot = pd.pivot_table(
        temp,
        index="alert_category",
        columns="alert_severity",
        aggfunc="size",
        fill_value=0,
    )
    pivot = pivot.reindex(columns=ordered, fill_value=0)
    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.reset_index()
    # Add grand total row
    totals = ["Total"] + [pivot[col].sum() for col in ordered] + [pivot["Total"].sum()]
    pivot.loc[len(pivot)] = totals
    pivot.columns = [str(c) for c in pivot.columns]
    return pivot


def build_ui() -> None:
    root = tk.Tk()
    root.title("Arbitrator Alerts Pivot")
    root.geometry("1180x780")
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.bind("<Key-q>", lambda e: root.destroy())
    root.bind("<Key-Q>", lambda e: root.destroy())
    root.bind("<Key-a>", lambda e: on_graphs_clicked())
    root.bind("<Key-A>", lambda e: on_graphs_clicked())

    status_var = tk.StringVar(value="Load data to begin")
    row_field = tk.StringVar(value="alert_name")
    col_field = tk.StringVar(value="vehicle_number")
    value_field = tk.StringVar(value="duration")
    agg_func = tk.StringVar(value="count")
    time_filter = tk.StringVar(value="all")
    severity_filter = tk.StringVar(value="FATAL")
    alert_search_text = tk.StringVar(value="")
    vehicle_id_width_px = tk.StringVar(value="40")
    vehicle_filter_text = tk.StringVar(value="")
    data_df: pd.DataFrame | None = None
    vehicle_universe_df: pd.DataFrame | None = None
    full_df: pd.DataFrame | None = None
    op_hours_df: pd.DataFrame | None = None
    op_hours_status_text = ""
    base_status_msg = ""
    selected_vehicle_vars: dict[str, tk.BooleanVar] = {}

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)

    top_row = ttk.Frame(main)
    top_row.pack(fill="x")
    top_row2 = ttk.Frame(main)
    top_row2.pack(fill="x", pady=(4, 0))

    def refresh_data() -> None:
        nonlocal data_df, full_df, base_status_msg, op_hours_df, op_hours_status_text
        df, msg = load_all()
        full_df = df
        op_hours_df, op_hours_status_text = load_op_hours()
        base_status_msg = msg
        apply_filters()
        rebuild_pivot()
        rebuild_catsev()
        populate_preview()
        populate_op_preview()
        refresh_vehicle_filter_options()
        rebuild_op_hours_tables()
        build_fleet_counts()

    ttk.Button(top_row, text="Refresh data", command=refresh_data).pack(
        side="left", padx=(0, 8)
    )

    def apply_filters() -> None:
        nonlocal data_df, vehicle_universe_df
        if full_df is None:
            data_df = None
            vehicle_universe_df = None
            status_var.set(base_status_msg or "Load data to begin")
            return

        time_value = time_filter.get()
        if time_value == "all":
            filtered = full_df
            days: int | None = None
        else:
            try:
                days = int(time_value)
            except ValueError:
                days = None
            filtered, warn = filter_last_days(full_df, days)
            if warn:
                messagebox.showwarning("Date filter", warn)
            if days is not None and filtered.empty:
                time_filter.set("all")
                filtered = full_df
                days = None

        severity_value = severity_filter.get()
        if severity_value != "all":
            if "alert_severity" in filtered.columns:
                filtered = filtered[filtered["alert_severity"].astype(str).str.upper() == severity_value]
            else:
                filtered = pd.DataFrame()

        # Capture the vehicle universe after range+severity filters, before alert-name search.
        vehicle_universe_df = filtered

        search_raw = alert_search_text.get().strip().lower()
        if search_raw:
            if "alert_name" in filtered.columns:
                tokens = [t for t in re.split(r"\\s+", search_raw) if t]
                series = filtered["alert_name"].astype(str).fillna("").str.lower()
                for token in tokens:
                    token_mask = series.str.contains(re.escape(token), na=False)
                    filtered = filtered[token_mask]
                    series = series[token_mask]
            else:
                filtered = pd.DataFrame()

        data_df = filtered
        suffix_parts: list[str] = []
        if days is not None:
            suffix_parts.append(f"last {days} days")
        if severity_value != "all":
            suffix_parts.append(f"severity={severity_value}")
        if search_raw:
            suffix_parts.append(f"search='{search_raw}'")
        suffix = f" | filter: {', '.join(suffix_parts)}" if suffix_parts else ""
        status_var.set(f"{base_status_msg}{suffix}")

    filter_frame = ttk.Frame(top_row)
    filter_frame.pack(side="left", padx=(8, 8))
    ttk.Label(filter_frame, text="Range:").pack(side="left", padx=(0, 6))

    def add_filter_button(label: str, value: str) -> None:
        tk.Radiobutton(
            filter_frame,
            text=label,
            value=value,
            variable=time_filter,
            command=lambda: (
                apply_filters(),
                rebuild_pivot(),
                rebuild_catsev(),
                populate_preview(),
                refresh_vehicle_filter_options(),
                build_fleet_counts(),
            ),
            indicatoron=0,
            width=max(5, len(label)),
        ).pack(side="left", padx=2)

    add_filter_button("All", "all")
    add_filter_button("7days", "7")
    add_filter_button("14days", "14")
    add_filter_button("30days", "30")
    add_filter_button("60days", "60")
    add_filter_button("90days", "90")
    add_filter_button("180", "180")
    add_filter_button("1year", "365")

    sev_frame = ttk.Frame(top_row)
    sev_frame.pack(side="left", padx=(8, 8))
    ttk.Label(sev_frame, text="Severity:").pack(side="left", padx=(0, 6))

    def add_severity_button(label: str, value: str) -> None:
        tk.Radiobutton(
            sev_frame,
            text=label,
            value=value,
            variable=severity_filter,
            command=lambda: (
                apply_filters(),
                rebuild_pivot(),
                rebuild_catsev(),
                populate_preview(),
                refresh_vehicle_filter_options(),
                build_fleet_counts(),
            ),
            indicatoron=0,
            width=max(6, len(label)),
        ).pack(side="left", padx=2)

    add_severity_button("All", "all")
    add_severity_button("Info", "INFO")
    add_severity_button("Warning", "WARNING")
    add_severity_button("Error", "ERROR")
    add_severity_button("Fatal", "FATAL")

    sev_controls_row2 = ttk.Frame(top_row2)
    sev_controls_row2.pack(side="left", padx=(8, 8))

    def refresh_after_vehicle_width_change() -> None:
        rebuild_pivot()
        populate_preview()

    ttk.Label(sev_controls_row2, text="Vehicle width:").pack(side="left", padx=(0, 6))
    vehicle_width_entry = ttk.Entry(
        sev_controls_row2, textvariable=vehicle_id_width_px, width=6
    )
    vehicle_width_entry.pack(side="left")
    ttk.Label(sev_controls_row2, text="px").pack(side="left", padx=(6, 12))
    vehicle_width_entry.bind("<Return>", lambda e: refresh_after_vehicle_width_change())
    vehicle_width_entry.bind("<FocusOut>", lambda e: refresh_after_vehicle_width_change())

    ttk.Label(sev_controls_row2, text="Search alerts:").pack(side="left", padx=(0, 6))
    search_entry = ttk.Entry(sev_controls_row2, textvariable=alert_search_text, width=22)
    search_entry.pack(side="left")

    search_after_id: str | None = None

    def refresh_after_search_change() -> None:
        apply_filters()
        rebuild_pivot()
        rebuild_catsev()
        populate_preview()
        refresh_vehicle_filter_options()
        build_fleet_counts()

    def on_search_key(event=None) -> None:
        nonlocal search_after_id
        if search_after_id is not None:
            try:
                root.after_cancel(search_after_id)
            except Exception:  # noqa: BLE001
                pass
        search_after_id = root.after(250, refresh_after_search_change)

    search_entry.bind("<Return>", lambda e: refresh_after_search_change())
    search_entry.bind("<FocusOut>", lambda e: refresh_after_search_change())
    search_entry.bind("<KeyRelease>", on_search_key)

    ttk.Label(main, textvariable=status_var).pack(anchor="w", pady=(6, 8))

    controls = ttk.LabelFrame(main, text="Pivot controls")
    controls.pack(fill="x", padx=2, pady=4)

    def add_field_picker(label: str, var: tk.StringVar) -> None:
        frame = ttk.Frame(controls)
        frame.pack(side="left", padx=6, pady=4)
        ttk.Label(frame, text=label).pack(anchor="w")
        ttk.OptionMenu(frame, var, var.get(), *DISPLAY_COLS).pack(fill="x")

    add_field_picker("Rows", row_field)
    add_field_picker("Columns", col_field)
    add_field_picker("Values (for sum/mean)", value_field)

    agg_frame = ttk.Frame(controls)
    agg_frame.pack(side="left", padx=6, pady=4)
    ttk.Label(agg_frame, text="Aggregation").pack(anchor="w")
    ttk.OptionMenu(agg_frame, agg_func, agg_func.get(), "count", "sum", "mean").pack(
        fill="x"
    )

    ttk.Button(controls, text="Build pivot", command=lambda: rebuild_pivot()).pack(
        side="left", padx=10, pady=8
    )

    def get_selected_alert_names(tree: ttk.Treeview) -> list[str]:
        selected = tree.selection()
        if not selected:
            return []
        alerts: list[str] = []
        for item_id in selected:
            values = tree.item(item_id).get("values", [])
            if not values:
                continue
            alert = str(values[0])
            if alert == "Total":
                continue
            alerts.append(alert)
        # Preserve order but de-dupe
        seen: set[str] = set()
        ordered: list[str] = []
        for alert in alerts:
            if alert not in seen:
                seen.add(alert)
                ordered.append(alert)
        return ordered

    def open_graphs_window(
        selected_alerts: list[str],
        source_df: pd.DataFrame | None = None,
        vehicles_df: pd.DataFrame | None = None,
    ) -> None:
        active_df = data_df if source_df is None else source_df
        if active_df is None or active_df.empty:
            messagebox.showinfo("Graphs", "No data loaded.")
            return
        if not selected_alerts:
            messagebox.showinfo("Graphs", "Select one or more alert rows first.")
            return

        if "alert_name" not in active_df.columns:
            messagebox.showerror("Graphs", "Data does not contain alert_name.")
            return
        if "vehicle_number" not in active_df.columns:
            messagebox.showerror("Graphs", "Data does not contain vehicle_number.")
            return

        subset = active_df[active_df["alert_name"].isin(selected_alerts)].copy()
        if subset.empty:
            messagebox.showinfo("Graphs", "No rows match the selected alert(s).")
            return

        subset["vehicle_number"] = subset["vehicle_number"].astype(str).fillna("")
        subset.loc[subset["vehicle_number"].str.strip() == "", "vehicle_number"] = "UNKNOWN"

        counts = (
            subset.groupby(["vehicle_number", "alert_name"], dropna=False)
            .size()
            .unstack(fill_value=0)
        )

        def sort_key(value: str) -> tuple[int, str]:
            txt = str(value)
            if txt.isdigit():
                return (0, f"{int(txt):08d}")
            return (1, txt)

        vehicles_basis = vehicles_df if vehicles_df is not None else active_df
        all_vehicles = vehicles_basis["vehicle_number"].astype(str).fillna("")
        all_vehicles = all_vehicles.mask(all_vehicles.str.strip() == "", "UNKNOWN")
        vehicles = sorted([str(v) for v in all_vehicles.unique().tolist()], key=sort_key)
        counts = counts.reindex(index=vehicles, columns=selected_alerts, fill_value=0)

        window_days = describe_time_window_days(active_df)
        days_text = "N/A" if window_days is None else str(window_days)
        title_text = (
            "Alert counts per vehicle (APM) for the selected time range "
            f"(last {days_text} days)"
        )

        win = tk.Toplevel(root)
        win.title(title_text)
        win.geometry("1100x650")
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.bind("<Key-q>", lambda e: win.destroy())
        win.bind("<Key-Q>", lambda e: win.destroy())

        header = ttk.Frame(win, padding=10)
        header.pack(fill="x")
        ttk.Label(
            header,
            text=f"Selected alerts ({len(selected_alerts)}): " + ", ".join(selected_alerts),
            wraplength=1000,
            justify="left",
        ).pack(anchor="w")

        chart_frame = ttk.Frame(win, padding=(10, 0, 10, 10))
        chart_frame.pack(fill="both", expand=True)

        toggle_frame = ttk.Frame(win, padding=(10, 0, 10, 6))
        toggle_frame.pack(fill="x")
        show_op_hours_var = tk.BooleanVar(value=False)
        normalize_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            toggle_frame,
            text="Show OP_hours",
            variable=show_op_hours_var,
            command=lambda: redraw_chart(),
        ).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(
            toggle_frame,
            text="Normalized with OP_hours",
            variable=normalize_var,
            command=lambda: redraw_chart(),
        ).pack(side="left")

        fig = Figure(figsize=(10, 5), dpi=100)
        ax = fig.add_subplot(111)
        ax2 = ax.twinx()

        series_alerts = selected_alerts

        x_vals = list(range(len(vehicles)))

        palette = [
            "#5b8def",
            "#f0ad4e",
            "#5cb85c",
            "#d9534f",
            "#9370db",
            "#20b2aa",
            "#ff7f50",
            "#708090",
        ]

        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Prepare op_hours map for normalization/overlay.
        op_hours_map: dict[str, float] = {}
        try:
            veh_hours_df = op_hours_by_vehicle()
            if not veh_hours_df.empty:
                for _, row in veh_hours_df.iterrows():
                    try:
                        op_hours_map[str(row["vehicle_number"])] = float(row["live_hours_total"])
                    except Exception:  # noqa: BLE001
                        continue
        except Exception:  # noqa: BLE001
            op_hours_map = {}

        def redraw_chart() -> None:
            ax.clear()

            use_normalized = normalize_var.get()
            include_op_hours = show_op_hours_var.get() and not use_normalized

            series_to_plot: list[tuple[str, list[float], str]] = []
            if use_normalized:
                for idx, alert in enumerate(series_alerts):
                    y_vals: list[float] = []
                    for v in vehicles:
                        base = float(counts.at[v, alert]) if (v in counts.index) else 0.0
                        denom = op_hours_map.get(v, 0.0)
                        if denom and denom != 0:
                            y_vals.append(round(base / denom, 3))
                        else:
                            y_vals.append(0.0)
                    series_to_plot.append((alert, y_vals, palette[idx % len(palette)]))
            else:
                for idx, alert in enumerate(series_alerts):
                    y = [int(v) for v in counts[alert].tolist()]
                    series_to_plot.append((alert, y, palette[idx % len(palette)]))

            # Plot alerts on ax; OP_hours (if requested) on ax2.
            ax_artists: list = []
            ax2_artists: list = []

            # Alerts layer
            n_series_alert = max(1, len(series_to_plot))
            total_width = 0.8
            bar_width = total_width / n_series_alert
            for idx, (label, y_vals, color) in enumerate(series_to_plot):
                offsets = [x + (idx - (n_series_alert - 1) / 2) * bar_width for x in x_vals]
                bars = ax.bar(
                    offsets,
                    y_vals,
                    width=bar_width * 0.95,
                    label=label,
                    color=color,
                    edgecolor="#2f5fb3",
                    picker=True,
                )
                ax_artists.append(bars)
                for bar_rect, vehicle in zip(bars, vehicles):
                    setattr(bar_rect, "_vehicle", vehicle)
                    setattr(bar_rect, "_alert", label)

            # OP hours layer on secondary y-axis
            ax2.cla()
            if include_op_hours:
                oh_vals = [float(op_hours_map.get(v, 0.0)) for v in vehicles]
                oh_offsets = x_vals  # align to centers
                oh_bars = ax2.bar(
                    oh_offsets,
                    oh_vals,
                    width=0.35,
                    label="OP_hours",
                    color="#666666",
                    edgecolor="#444444",
                    alpha=0.6,
                )
                ax2_artists.append(oh_bars)
                ax2.set_ylabel("OP_hours")
            else:
                ax2.set_ylabel("")
                ax2.set_yticks([])

            ax.set_xticks(x_vals)
            ax.set_xticklabels(vehicles, rotation=45, ha="right")
            ax.set_xlabel("APM (vehicle number)")
            ax.set_ylabel("Alert count" if not use_normalized else "Alert count per OP_hour")
            title_extra = " (normalized)" if use_normalized else ""
            ax.set_title(title_text + title_extra)

            # Build combined legend
            handles, labels = ax.get_legend_handles_labels()
            if include_op_hours and ax2_artists:
                h2, l2 = ax2.get_legend_handles_labels()
                handles += h2
                labels += l2
            ax.legend(handles, labels)

            fig.tight_layout()
            canvas.draw()

        redraw_chart()

        def show_vehicle_alert_timeseries(vehicle: str, alert_name: str) -> None:
            base_df = subset.copy()
            if "_start_dt" not in base_df.columns:
                messagebox.showinfo("Vehicle click", "No timestamps to build daily counts.")
                return
            base_df["_start_dt"] = pd.to_datetime(base_df["_start_dt"], errors="coerce")
            base_df = base_df.dropna(subset=["_start_dt"])
            if base_df.empty:
                messagebox.showinfo("Vehicle click", "No timestamps to build daily counts.")
                return
            min_dt = base_df["_start_dt"].min().normalize()
            max_dt = base_df["_start_dt"].max().normalize()
            if pd.isna(min_dt) or pd.isna(max_dt):
                messagebox.showinfo("Vehicle click", "No timestamps to build daily counts.")
                return

            normalized_vehicle = normalize_vehicle_number(base_df["vehicle_number"])
            mask = (normalized_vehicle == str(vehicle)) & (base_df["alert_name"] == alert_name)
            filtered = base_df[mask]
            if filtered.empty:
                messagebox.showinfo(
                    "Vehicle click", f"No rows for vehicle {vehicle} and alert {alert_name}."
                )
                return

            daily = (
                filtered.set_index("_start_dt")
                .groupby(pd.Grouper(freq="D"))
                .size()
                .reset_index(name="count")
            )
            # Ensure the entire date range (per current filtered data) is covered, filling zeros.
            full_range = pd.date_range(start=min_dt, end=max_dt, freq="D")
            daily = daily.set_index("_start_dt").reindex(full_range, fill_value=0).reset_index()
            daily.rename(columns={"index": "_start_dt"}, inplace=True)
            daily["_start_dt"] = daily["_start_dt"].dt.date
            daily = daily.sort_values("_start_dt")
            x_labels = [str(d) for d in daily["_start_dt"].tolist()]
            y_vals = daily["count"].astype(int).tolist()
            x_pos = list(range(len(x_labels)))

            win_ts = tk.Toplevel(root)
            win_ts.title(f"{alert_name} daily counts for vehicle {vehicle}")
            win_ts.geometry("900x520")
            win_ts.protocol("WM_DELETE_WINDOW", win_ts.destroy)
            win_ts.bind("<Key-q>", lambda e: win_ts.destroy())
            win_ts.bind("<Key-Q>", lambda e: win_ts.destroy())

            header_ts = ttk.Frame(win_ts, padding=10)
            header_ts.pack(fill="x")
            ttk.Label(
                header_ts,
                text=f"Vehicle: {vehicle} | Alert: {alert_name}",
                justify="left",
            ).pack(anchor="w")

            chart_ts = ttk.Frame(win_ts, padding=(10, 0, 10, 10))
            chart_ts.pack(fill="both", expand=True)

            fig_ts = Figure(figsize=(8.5, 3.5), dpi=100)
            ax_ts = fig_ts.add_subplot(111)
            ax_ts.bar(x_pos, y_vals, width=0.8, color="#5b8def", edgecolor="#2f5fb3")
            ax_ts.set_xticks(x_pos)
            ax_ts.set_xticklabels(x_labels, rotation=45, ha="right")
            ax_ts.set_xlabel("Day")
            ax_ts.set_ylabel("Count")
            ax_ts.set_title(f"Daily counts for {alert_name} on {vehicle}")
            fig_ts.tight_layout()

            canvas_ts = FigureCanvasTkAgg(fig_ts, master=chart_ts)
            canvas_ts.draw()
            canvas_ts.get_tk_widget().pack(fill="both", expand=True)

            btn_ts = ttk.Frame(win_ts, padding=10)
            btn_ts.pack(fill="x")
            ttk.Button(btn_ts, text="Close (q)", command=win_ts.destroy).pack(side="right")

        def on_pick(event) -> None:
            artist = getattr(event, "artist", None)
            if artist is None:
                return
            vehicle = getattr(artist, "_vehicle", None)
            alert_name = getattr(artist, "_alert", None)
            if vehicle is None:
                return
            if alert_name is None:
                messagebox.showinfo("Vehicle click", f"Vehicle {vehicle} clicked.\nOK?")
                return
            show_vehicle_alert_timeseries(vehicle, alert_name)

        canvas.mpl_connect("pick_event", on_pick)

        btn_row = ttk.Frame(win, padding=10)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Close (q)", command=win.destroy).pack(side="right")

    def open_fleet_totals_graph(selected_alerts: list[str]) -> None:
        subset = filtered_by_vehicle(data_df)
        if subset.empty:
            messagebox.showinfo("Graphs", "No data after vehicle filter.")
            return
        if "alert_name" not in subset.columns:
            messagebox.showerror("Graphs", "Data does not contain alert_name.")
            return

        window_days = describe_time_window_days(subset)
        days_text = "N/A" if window_days is None else str(window_days)

        win = tk.Toplevel(root)
        win.title("Total alert counts for selected vehicles")
        win.geometry("1100x700")
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.bind("<Key-q>", lambda e: win.destroy())
        win.bind("<Key-Q>", lambda e: win.destroy())

        header = ttk.Frame(win, padding=10)
        header.pack(fill="x")
        ttk.Label(
            header,
            text=f"Selected alerts ({len(selected_alerts)}): " + ", ".join(selected_alerts),
            wraplength=1000,
            justify="left",
        ).pack(anchor="w")

        controls_row = ttk.Frame(win, padding=(10, 0, 10, 6))
        controls_row.pack(fill="x")
        ttk.Label(controls_row, text="Duration:").pack(side="left", padx=(0, 6))
        duration_var = tk.StringVar(value="all")
        duration_labels = [
            ("All", "all"),
            ("Daly", "daly"),
            ("Weekly", "weekly"),
            ("Monthly", "monthly"),
            ("Quotally", "quotally"),
        ]
        for label, value in duration_labels:
            tk.Radiobutton(
                controls_row,
                text=label,
                value=value,
                variable=duration_var,
                indicatoron=0,
                width=max(6, len(label)),
                command=lambda: redraw_chart(),
            ).pack(side="left", padx=3)

        chart_frame = ttk.Frame(win, padding=(10, 0, 10, 10))
        chart_frame.pack(fill="both", expand=True)

        fig = Figure(figsize=(10, 5.6), dpi=100)
        ax = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.get_tk_widget().pack(fill="both", expand=True)

        palette = [
            "#5b8def",
            "#f0ad4e",
            "#5cb85c",
            "#d9534f",
            "#9370db",
            "#20b2aa",
            "#ff7f50",
            "#708090",
        ]

        def _format_bucket_label(ts: pd.Timestamp, duration: str) -> str:
            if duration == "daly":
                return ts.strftime("%Y-%m-%d")
            if duration == "weekly":
                return f"Week of {ts.strftime('%Y-%m-%d')}"
            if duration == "monthly":
                return ts.strftime("%Y-%m")
            if duration == "quotally":
                quarter = ((ts.month - 1) // 3) + 1
                return f"{ts.year}-Q{quarter}"
            return "All time"

        def aggregate_counts(duration: str) -> pd.DataFrame:
            working = subset[subset["alert_name"].isin(selected_alerts)].copy()
            working = add_parsed_timestamps(working)
            if working.empty:
                return pd.DataFrame()

            if duration == "all":
                totals = (
                    working.groupby("alert_name", dropna=False)
                    .size()
                    .reindex(selected_alerts, fill_value=0)
                )
                df = pd.DataFrame([totals.tolist()], columns=selected_alerts)
                df.insert(0, "bucket", ["All time"])
                return df

            if "_start_dt" not in working.columns:
                return pd.DataFrame()

            working = working.dropna(subset=["_start_dt"])
            if working.empty:
                return pd.DataFrame()

            freq_map = {
                "daly": "D",
                "weekly": "W",
                "monthly": "M",
                "quotally": "Q",
            }
            freq = freq_map.get(duration)
            if freq is None:
                return pd.DataFrame()

            try:
                working_indexed = working.set_index("_start_dt")
            except KeyError:
                return pd.DataFrame()

            try:
                grouped = (
                    working_indexed.groupby([pd.Grouper(freq=freq), "alert_name"])
                    .size()
                    .unstack(fill_value=0)
                )
            except KeyError:
                return pd.DataFrame()
            grouped = grouped.reindex(columns=selected_alerts, fill_value=0)
            if grouped.empty:
                return pd.DataFrame()

            grouped = grouped.reset_index(names=["bucket"])
            grouped["bucket"] = grouped["bucket"].apply(
                lambda ts: _format_bucket_label(ts, duration)
            )
            return grouped

        def redraw_chart() -> None:
            duration = duration_var.get()
            data = aggregate_counts(duration)
            ax.clear()

            if data.empty:
                ax.text(
                    0.5,
                    0.5,
                    "No data to plot for this duration.",
                    ha="center",
                    va="center",
                    fontsize=12,
                    color="#555",
                )
                ax.set_xticks([])
                ax.set_yticks([])
                canvas.draw()
                return

            buckets = data["bucket"].tolist()
            x_vals = list(range(len(buckets)))
            n_series = max(1, len(selected_alerts))
            total_width = 0.8
            bar_width = total_width / n_series

            for idx, alert in enumerate(selected_alerts):
                y_vals = data[alert].astype(int).tolist()
                offsets = [x + (idx - (n_series - 1) / 2) * bar_width for x in x_vals]
                ax.bar(
                    offsets,
                    y_vals,
                    width=bar_width * 0.95,
                    label=alert,
                    color=palette[idx % len(palette)],
                    edgecolor="#2f5fb3",
                )

            ax.set_xticks(x_vals)
            ax.set_xticklabels(buckets, rotation=45, ha="right")
            ax.set_xlabel("Time")
            ax.set_ylabel("Alert count")
            ax.set_title(
                f"Total alert counts for selected vehicles (last {days_text} days) — {duration.capitalize()}"
            )
            ax.legend()
            fig.tight_layout()
            canvas.draw()

        redraw_chart()

        btn_row = ttk.Frame(win, padding=10)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Close (q)", command=win.destroy).pack(side="right")

    def on_graphs_clicked() -> None:
        try:
            tab_text = notebook.tab(notebook.select(), "text")
        except Exception:  # noqa: BLE001
            tab_text = "BY VEHICLE"

        tab_map: dict[str, tuple[ttk.Treeview, bool]] = {
            "BY VEHICLE": (pivot_tree, True),
            "BY FLEET": (fleet_tree, False),
        }

        mapping = tab_map.get(tab_text)
        if mapping is None:
            messagebox.showinfo(
                "Graphs",
                f"Graphs is not available for the '{tab_text}' tab.\n"
                "Use BY VEHICLE or BY FLEET.",
            )
            return

        tree, requires_alert_rows = mapping
        if requires_alert_rows and row_field.get() != "alert_name":
            messagebox.showinfo(
                "Graphs",
                f"Graphs expects Rows = alert_name.\nCurrent Rows = {row_field.get()}.\n"
                "Set Rows to alert_name and rebuild the pivot, then select alerts.",
            )
            return

        selected_alerts = get_selected_alert_names(tree)
        if not selected_alerts:
            messagebox.showinfo("Graphs", "Select one or more alert rows first.")
            return

        if tab_text == "BY FLEET":
            open_fleet_totals_graph(selected_alerts)
            return

        open_graphs_window(selected_alerts, vehicles_df=vehicle_universe_df)

    ttk.Button(controls, text="Graphs (a)", command=on_graphs_clicked).pack(
        side="left", padx=6, pady=8
    )

    bottom_bar = ttk.Frame(main, padding=(0, 8, 0, 0))
    bottom_bar.pack(side="bottom", fill="x")
    ttk.Button(bottom_bar, text="Close (q)", command=root.destroy).pack(side="right")

    notebook = ttk.Notebook(main)
    notebook.pack(fill="both", expand=True, pady=(6, 0))

    pivot_tab = ttk.Frame(notebook)
    fleet_tab = ttk.Frame(notebook)
    catsev_tab = ttk.Frame(notebook)
    preview_tab = ttk.Frame(notebook)
    op_preview_tab = ttk.Frame(notebook)
    notebook.add(pivot_tab, text="BY VEHICLE")
    notebook.add(fleet_tab, text="BY FLEET")
    notebook.add(catsev_tab, text="Category x Severity")
    notebook.add(preview_tab, text="Data preview")
    notebook.add(op_preview_tab, text="OP HOURS preview")

    # Pivot table view
    pivot_tree = ttk.Treeview(pivot_tab, show="headings", selectmode="extended")
    pscroll_y = ttk.Scrollbar(pivot_tab, orient="vertical", command=pivot_tree.yview)
    pscroll_x = ttk.Scrollbar(pivot_tab, orient="horizontal", command=pivot_tree.xview)
    pivot_tree.configure(yscrollcommand=pscroll_y.set, xscrollcommand=pscroll_x.set)
    pivot_tree.pack(side="top", fill="both", expand=True)
    pscroll_y.pack(side="right", fill="y")
    pscroll_x.pack(side="bottom", fill="x")

    op_hours_vehicle_frame = ttk.LabelFrame(pivot_tab, text="OP HOURS by vehicle (sum live_hours)")
    op_hours_vehicle_frame.pack(side="top", fill="x", padx=4, pady=(6, 4))
    op_hours_vehicle_tree = ttk.Treeview(op_hours_vehicle_frame, show="headings", height=1)
    ohv_scroll_y = ttk.Scrollbar(op_hours_vehicle_frame, orient="vertical", command=op_hours_vehicle_tree.yview)
    ohv_scroll_x = ttk.Scrollbar(op_hours_vehicle_frame, orient="horizontal", command=op_hours_vehicle_tree.xview)
    op_hours_vehicle_tree.configure(yscrollcommand=ohv_scroll_y.set, xscrollcommand=ohv_scroll_x.set)
    op_hours_vehicle_tree.pack(side="top", fill="x", expand=False)
    ohv_scroll_y.pack(side="right", fill="y")
    ohv_scroll_x.pack(side="bottom", fill="x")
    op_hours_vehicle_status = tk.StringVar(value="op_hours.csv not loaded")
    ttk.Label(op_hours_vehicle_frame, textvariable=op_hours_vehicle_status, anchor="w").pack(
        side="top", fill="x", padx=4, pady=(2, 2)
    )

    # Category x Severity view
    catsev_tree = ttk.Treeview(catsev_tab, show="headings")
    cscroll_y = ttk.Scrollbar(catsev_tab, orient="vertical", command=catsev_tree.yview)
    cscroll_x = ttk.Scrollbar(catsev_tab, orient="horizontal", command=catsev_tree.xview)
    catsev_tree.configure(yscrollcommand=cscroll_y.set, xscrollcommand=cscroll_x.set)
    catsev_tree.pack(side="top", fill="both", expand=True)
    cscroll_y.pack(side="right", fill="y")
    cscroll_x.pack(side="bottom", fill="x")

    # Preview view
    preview_tree = ttk.Treeview(preview_tab, show="headings")
    prscroll_y = ttk.Scrollbar(preview_tab, orient="vertical", command=preview_tree.yview)
    prscroll_x = ttk.Scrollbar(preview_tab, orient="horizontal", command=preview_tree.xview)
    preview_tree.configure(yscrollcommand=prscroll_y.set, xscrollcommand=prscroll_x.set)
    preview_tree.pack(side="top", fill="both", expand=True)
    prscroll_y.pack(side="right", fill="y")
    prscroll_x.pack(side="bottom", fill="x")

    op_preview_tree = ttk.Treeview(op_preview_tab, show="headings")
    oprscroll_y = ttk.Scrollbar(op_preview_tab, orient="vertical", command=op_preview_tree.yview)
    oprscroll_x = ttk.Scrollbar(op_preview_tab, orient="horizontal", command=op_preview_tree.xview)
    op_preview_tree.configure(yscrollcommand=oprscroll_y.set, xscrollcommand=oprscroll_x.set)
    op_preview_tree.pack(side="top", fill="both", expand=True)
    oprscroll_y.pack(side="right", fill="y")
    oprscroll_x.pack(side="bottom", fill="x")

    # BY FLEET view
    fleet_info = ttk.Label(
        fleet_tab,
        text="Fleet alert totals (rows = alert_name, column = count; uses current range + vehicle filter)",
        anchor="w",
    )
    fleet_info.pack(fill="x", pady=(4, 4), padx=4)

    vehicle_filter_frame = ttk.LabelFrame(fleet_tab, text="Filter by APM (vehicle number)")
    vehicle_filter_frame.pack(fill="both", padx=4, pady=(0, 6))
    filter_controls = ttk.Frame(vehicle_filter_frame)
    filter_controls.pack(fill="x", pady=(4, 2))
    ttk.Label(filter_controls, text="Filter:").pack(side="left")
    filter_entry = ttk.Entry(filter_controls, textvariable=vehicle_filter_text, width=24)
    filter_entry.pack(side="left", padx=(4, 10))

    vehicle_list_inner = ttk.Frame(vehicle_filter_frame)
    vehicle_list_inner.pack(fill="x", padx=4, pady=(0, 4))

    fleet_tree = ttk.Treeview(fleet_tab, show="headings")
    fscroll_y = ttk.Scrollbar(fleet_tab, orient="vertical", command=fleet_tree.yview)
    fscroll_x = ttk.Scrollbar(fleet_tab, orient="horizontal", command=fleet_tree.xview)
    fleet_tree.configure(yscrollcommand=fscroll_y.set, xscrollcommand=fscroll_x.set)
    fleet_tree.pack(side="top", fill="both", expand=True)
    fscroll_y.pack(side="right", fill="y")
    fscroll_x.pack(side="bottom", fill="x")

    op_hours_fleet_frame = ttk.LabelFrame(fleet_tab, text="OP HOURS by fleet (total live_hours)")
    op_hours_fleet_frame.pack(side="top", fill="x", padx=4, pady=(6, 4))
    op_hours_fleet_tree = ttk.Treeview(op_hours_fleet_frame, show="headings", height=4)
    ohf_scroll_y = ttk.Scrollbar(op_hours_fleet_frame, orient="vertical", command=op_hours_fleet_tree.yview)
    ohf_scroll_x = ttk.Scrollbar(op_hours_fleet_frame, orient="horizontal", command=op_hours_fleet_tree.xview)
    op_hours_fleet_tree.configure(yscrollcommand=ohf_scroll_y.set, xscrollcommand=ohf_scroll_x.set)
    op_hours_fleet_tree.pack(side="top", fill="x", expand=False)
    ohf_scroll_y.pack(side="right", fill="y")
    ohf_scroll_x.pack(side="bottom", fill="x")
    op_hours_fleet_status = tk.StringVar(value="op_hours.csv not loaded")
    ttk.Label(op_hours_fleet_frame, textvariable=op_hours_fleet_status, anchor="w").pack(
        side="top", fill="x", padx=4, pady=(2, 2)
    )

    def populate_tree(tree: ttk.Treeview, df: pd.DataFrame, max_rows: int = 300) -> None:
        tree.delete(*tree.get_children())
        if df.empty:
            tree["columns"] = []
            return
        columns = list(df.columns)
        tree["columns"] = columns
        try:
            font = tkfont.nametofont(str(tree.cget("font")))
        except Exception:  # noqa: BLE001
            font = tkfont.nametofont("TkDefaultFont")

        def _get_vehicle_width_px() -> int:
            raw = vehicle_id_width_px.get().strip()
            try:
                value = int(raw)
            except ValueError:
                value = 90
            value = min(2000, value)
            if value < 0:
                value = 0
            if raw != str(value):
                vehicle_id_width_px.set(str(value))
            return value

        def width_for_column(col: str) -> int:
            if col == "vehicle_number" or col.isdigit() or col == "UNKNOWN":
                return _get_vehicle_width_px()
            if col == "alert_name":
                return max(140, font.measure("0" * 30) + 24)
            return 140

        for col in columns:
            tree.heading(col, text=col)
            is_vehicle_col = col == "vehicle_number" or col.isdigit() or col == "UNKNOWN"
            tree.column(
                col,
                width=width_for_column(col),
                anchor="w",
                stretch=not is_vehicle_col,
            )
        for _, row in df.head(max_rows).iterrows():
            tree.insert("", "end", values=[row.get(col, "") for col in columns])

    def _vehicle_sort_key(value: str) -> tuple[int, str]:
        txt = str(value)
        if txt.isdigit():
            return (0, f"{int(txt):08d}")
        return (1, txt)

    def get_vehicle_list(df: pd.DataFrame | None) -> list[str]:
        if df is None or df.empty or "vehicle_number" not in df.columns:
            return []
        series = normalize_vehicle_number(df["vehicle_number"])
        return sorted([str(v) for v in series.unique().tolist()], key=_vehicle_sort_key)

    def refresh_vehicle_filter_options() -> None:
        visible_filter = vehicle_filter_text.get().strip().lower()
        basis_df = vehicle_universe_df if vehicle_universe_df is not None else data_df
        current_vehicles = get_vehicle_list(basis_df)

        any_selected = any(var.get() for var in selected_vehicle_vars.values())
        default_new = True if not selected_vehicle_vars else any_selected

        for vehicle in current_vehicles:
            if vehicle not in selected_vehicle_vars:
                selected_vehicle_vars[vehicle] = tk.BooleanVar(value=default_new)

        for child in vehicle_list_inner.winfo_children():
            child.destroy()

        shown = 0
        max_rows = 6
        current_col = 0
        current_row = 0
        for vehicle in current_vehicles:
            if visible_filter and visible_filter not in vehicle.lower():
                continue
            ttk.Checkbutton(
                vehicle_list_inner,
                text=vehicle,
                variable=selected_vehicle_vars[vehicle],
                command=build_fleet_counts,
            ).grid(row=current_row, column=current_col, sticky="w", padx=6, pady=2)
            current_row += 1
            if current_row >= max_rows:
                current_row = 0
                current_col += 1
            shown += 1

        if shown == 0:
            ttk.Label(vehicle_list_inner, text="No vehicles match filter").grid(
                row=0, column=0, sticky="w", padx=6, pady=2
            )

    def select_all_vehicles(value: bool) -> None:
        for var in selected_vehicle_vars.values():
            var.set(value)
        build_fleet_counts()

    def filtered_by_vehicle(df: pd.DataFrame | None) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame()
        selected = [vid for vid, var in selected_vehicle_vars.items() if var.get()]
        all_ids = set(get_vehicle_list(df))
        if not selected:
            return pd.DataFrame()
        if set(selected) >= all_ids:
            return df.copy()
        normalized = normalize_vehicle_number(df["vehicle_number"])
        return df[normalized.isin(selected)].copy()

    def build_fleet_counts() -> None:
        if data_df is None or data_df.empty:
            populate_tree(fleet_tree, pd.DataFrame())
            return
        if "alert_name" not in data_df.columns:
            populate_tree(fleet_tree, pd.DataFrame())
            status_var.set("Missing alert_name; cannot build BY FLEET.")
            return

        subset = filtered_by_vehicle(data_df)
        if subset.empty:
            populate_tree(fleet_tree, pd.DataFrame())
            status_var.set("No data after vehicle filter.")
            return

        counts_df = (
            subset.groupby("alert_name", dropna=False)
            .size()
            .reset_index(name="Count")
            .sort_values("Count", ascending=False)
        )
        counts_df["alert_name"] = counts_df["alert_name"].astype(str)
        populate_tree(fleet_tree, counts_df, max_rows=2000)

    ttk.Button(filter_controls, text="Select all", command=lambda: select_all_vehicles(True)).pack(
        side="left", padx=(0, 4)
    )
    ttk.Button(filter_controls, text="Clear all", command=lambda: select_all_vehicles(False)).pack(
        side="left"
    )
    ttk.Button(filter_controls, text="Refresh", command=lambda: (refresh_vehicle_filter_options(), build_fleet_counts())).pack(
        side="left", padx=(6, 0)
    )
    filter_entry.bind("<Return>", lambda e: (refresh_vehicle_filter_options(), build_fleet_counts()))

    def on_tab_change(event) -> None:
        tab_text = event.widget.tab(event.widget.select(), "text")
        if tab_text == "BY FLEET":
            refresh_vehicle_filter_options()
            build_fleet_counts()

    def rebuild_pivot() -> None:
        if data_df is None or data_df.empty:
            populate_tree(pivot_tree, pd.DataFrame())
            return
        agg = agg_func.get()
        rows = row_field.get()
        cols = col_field.get()
        val = value_field.get()
        if agg == "count":
            val_field = None
        else:
            val_field = val
        pivot_df = pivot_dataframe(data_df, rows, cols, val_field, agg)
        if (
            not pivot_df.empty
            and cols == "vehicle_number"
            and vehicle_universe_df is not None
            and not vehicle_universe_df.empty
            and "vehicle_number" in vehicle_universe_df.columns
        ):
            universe = normalize_vehicle_number(vehicle_universe_df["vehicle_number"])
            all_vehicle_cols = sorted([str(v) for v in universe.unique().tolist()], key=_vehicle_sort_key)
            if rows in pivot_df.columns:
                ordered_cols = [rows] + all_vehicle_cols
                for col in pivot_df.columns:
                    if col not in ordered_cols:
                        ordered_cols.append(col)
                pivot_df = pivot_df.reindex(columns=ordered_cols, fill_value=0)

        if pivot_df.empty:
            status_var.set("Pivot empty; check fields/aggregation.")
        populate_tree(pivot_tree, pivot_df)

    def rebuild_catsev() -> None:
        if data_df is None or data_df.empty:
            populate_tree(catsev_tree, pd.DataFrame())
            return
        catsev_df = category_severity_table(data_df)
        populate_tree(catsev_tree, catsev_df)

    def op_hours_by_vehicle() -> pd.DataFrame:
        if op_hours_df is None or op_hours_df.empty:
            return pd.DataFrame()
        # Pick the first matching hours column.
        hour_col: str | None = None
        for candidate in ("live_hours", "live_hour", "op_hour"):
            if candidate in op_hours_df.columns:
                hour_col = candidate
                break
        if hour_col is None:
            return pd.DataFrame()
        if "vehicle_number" in op_hours_df.columns:
            vehicles = normalize_vehicle_number(op_hours_df["vehicle_number"])
        elif "vehicle_name" in op_hours_df.columns:
            vehicles = normalize_vehicle_number(op_hours_df["vehicle_name"].map(_extract_vehicle_number))
        else:
            return pd.DataFrame()
        try:
            hours = pd.to_numeric(op_hours_df[hour_col], errors="coerce")
        except Exception:  # noqa: BLE001
            return pd.DataFrame()
        names_col = op_hours_df["vehicle_name"] if "vehicle_name" in op_hours_df.columns else None
        df = pd.DataFrame(
            {
                "vehicle_number": vehicles,
                "vehicle_name": names_col if names_col is not None else "",
                "live_hours": hours,
            }
        )
        df = df.dropna(subset=["live_hours"])
        if df.empty:
            return pd.DataFrame()
        def pick_name(series: pd.Series) -> str:
            non_empty = [str(x) for x in series.dropna() if str(x).strip() != ""]
            return non_empty[0] if non_empty else ""

        grouped = (
            df.groupby("vehicle_number", dropna=False)
            .agg(live_hours_total=("live_hours", "sum"), vehicle_name=("vehicle_name", pick_name))
            .reset_index()
        )
        grouped["live_hours_total"] = grouped["live_hours_total"].round(0).astype(int)
        grouped = grouped.sort_values("vehicle_number", key=lambda s: s.map(_vehicle_sort_key))
        return grouped

    def op_hours_by_fleet() -> pd.DataFrame:
        vehicle_df = op_hours_by_vehicle()
        if vehicle_df.empty:
            return pd.DataFrame()
        total = vehicle_df["live_hours_total"].sum()
        return pd.DataFrame([{"fleet": "ALL", "live_hours_total": int(round(total))}])

    def rebuild_op_hours_tables() -> None:
        status_fallback = op_hours_status_text or "op_hours.csv not loaded"
        veh_df = op_hours_by_vehicle()
        if veh_df.empty:
            op_hours_vehicle_status.set(status_fallback)
            populate_tree(op_hours_vehicle_tree, pd.DataFrame())
        else:
            op_hours_vehicle_status.set(op_hours_status_text or "live_hours loaded")
            # Build a pivot-style single-row table: row = OP hours, cols = vehicle
            vehicle_cols = [str(v) for v in veh_df["vehicle_number"].tolist()]
            hours_vals = [int(v) for v in veh_df["live_hours_total"].tolist()]
            pivot_df = pd.DataFrame(
                [["OP hours"] + hours_vals],
                columns=["metric"] + vehicle_cols,
            )
            populate_tree(op_hours_vehicle_tree, pivot_df, max_rows=10)
        fleet_df = op_hours_by_fleet()
        if fleet_df.empty:
            op_hours_fleet_status.set(status_fallback)
        else:
            op_hours_fleet_status.set(op_hours_status_text or "live_hours loaded")
        populate_tree(op_hours_fleet_tree, fleet_df, max_rows=10)

    def populate_preview() -> None:
        if data_df is None or data_df.empty:
            populate_tree(preview_tree, pd.DataFrame())
            return
        subset_cols = [c for c in DISPLAY_COLS if c in data_df.columns]
        preview_df = data_df[subset_cols]
        populate_tree(preview_tree, preview_df)

    def populate_op_preview() -> None:
        if op_hours_df is None or op_hours_df.empty:
            populate_tree(op_preview_tree, pd.DataFrame())
            return
        populate_tree(op_preview_tree, op_hours_df)

    notebook.bind("<<NotebookTabChanged>>", on_tab_change)

    refresh_data()
    root.mainloop()


if __name__ == "__main__":
    build_ui()
