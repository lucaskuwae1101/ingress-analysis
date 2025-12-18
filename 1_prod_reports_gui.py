#!/usr/bin/env python3
"""
Small Tkinter GUI to list prod report folders, compare with local downloads, and fetch CSVs.

Remote base: https://rosbag.data.ventitechnologies.net/analytics/prod_reports/
Local downloads land in ./logs/<folder>/arbitrator.csv and op_hour.csv.
"""

from __future__ import annotations

import re
import shutil
import threading
import time
import tkinter as tk
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import messagebox, ttk
from urllib.error import URLError
from urllib.request import urlopen

BASE_URL = "https://rosbag.data.ventitechnologies.net/analytics/prod_reports/"
LOGS_DIR = Path(__file__).resolve().parent / "logs"
FILES_TO_FETCH = ("arbitrator.csv", "op_hour.csv")
FOLDER_PATTERN = re.compile(r"(\d{8}-\d{4}_\d{8}-\d{4})/?")


def parse_folder_name(name: str) -> tuple[datetime | None, datetime | None]:
    """Return start/end datetimes parsed from a folder name."""
    match = re.match(r"(\d{8})-(\d{4})_(\d{8})-(\d{4})", name)
    if not match:
        return None, None
    start_date, start_time, end_date, end_time = match.groups()
    try:
        start_dt = datetime.strptime(start_date + start_time, "%Y%m%d%H%M")
        end_dt = datetime.strptime(end_date + end_time, "%Y%m%d%H%M")
        return start_dt, end_dt
    except ValueError:
        return None, None


def folder_sort_key(name: str) -> tuple[int, str]:
    """Sort most recent first; fall back to name."""
    start_dt, _ = parse_folder_name(name)
    timestamp = int(start_dt.timestamp()) if start_dt else 0
    return (-timestamp, name)


def fetch_remote_folders() -> tuple[bool, list[str], str]:
    """Fetch folder names from the prod_reports directory listing."""
    try:
        with urlopen(BASE_URL, timeout=20) as response:
            html = response.read().decode("utf-8", errors="replace")
        names = sorted(set(FOLDER_PATTERN.findall(html)), key=folder_sort_key)
        if not names:
            return False, [], "No folders found in remote listing"
        return True, names, f"Found {len(names)} folders on server"
    except URLError as exc:
        return False, [], f"Network error: {exc}"
    except Exception as exc:  # noqa: BLE001
        return False, [], f"Failed to fetch remote listing: {exc}"


def local_folder_status() -> tuple[set[str], set[str]]:
    """
    Return (complete, partial) folder sets.

    Complete means both CSVs exist; partial means folder exists but one is missing.
    """
    complete: set[str] = set()
    partial: set[str] = set()
    if not LOGS_DIR.exists():
        return complete, partial

    for p in LOGS_DIR.iterdir():
        if not p.is_dir() or not FOLDER_PATTERN.match(p.name):
            continue
        files_present = {f.name for f in p.glob("*.csv")}
        if all(name in files_present for name in FILES_TO_FETCH):
            complete.add(p.name)
        else:
            partial.add(p.name)
    return complete, partial


def stream_download(
    folder: str,
    filename: str,
    dest_path: Path,
    progress_hook,
    cancel_flag,
) -> tuple[bool, str, int]:
    """
    Stream a single file with progress updates.

    Returns success flag, message, and bytes downloaded.
    """
    url = f"{BASE_URL}{folder}/{filename}"
    temp_path = dest_path.with_suffix(dest_path.suffix + ".part")
    downloaded = 0
    total = 0
    start_time = time.time()
    try:
        with urlopen(url, timeout=30) as resp:
            total_header = resp.getheader("Content-Length")
            total = int(total_header) if total_header and total_header.isdigit() else 0
            with temp_path.open("wb") as handle:
                while True:
                    chunk = resp.read(8192)
                    if not chunk:
                        break
                    if cancel_flag():
                        temp_path.unlink(missing_ok=True)
                        return False, "Cancelled", downloaded
                    handle.write(chunk)
                    downloaded += len(chunk)
                    if progress_hook:
                        progress_hook(filename, downloaded, total, time.time() - start_time)
        temp_path.rename(dest_path)
        return True, "ok", downloaded
    except URLError as exc:
        temp_path.unlink(missing_ok=True)
        return False, f"Network error while downloading {filename}: {exc}", downloaded
    except Exception as exc:  # noqa: BLE001
        temp_path.unlink(missing_ok=True)
        return False, f"Failed to save {filename}: {exc}", downloaded


def download_folder(
    folder: str,
    progress_hook,
    cancel_flag,
) -> tuple[bool, str]:
    """Download arbitrator.csv and op_hour.csv for a specific folder."""
    dest_dir = LOGS_DIR / folder
    existed_before = dest_dir.exists()
    dest_dir.mkdir(parents=True, exist_ok=True)
    try:
        for filename in FILES_TO_FETCH:
            dest_path = dest_dir / filename

            def per_file_progress(name, downloaded, total, elapsed) -> None:
                progress_hook(folder, name, downloaded, total, elapsed)

            ok, msg, _ = stream_download(folder, filename, dest_path, per_file_progress, cancel_flag)
            if not ok:
                raise RuntimeError(msg)
        return True, f"Downloaded {', '.join(FILES_TO_FETCH)} to {dest_dir}"
    except Exception as exc:  # noqa: BLE001
        if not existed_before and dest_dir.exists():
            shutil.rmtree(dest_dir, ignore_errors=True)
        if str(exc) == "Cancelled":
            return False, "Cancelled"
        return False, str(exc)


def build_ui() -> None:
    LOGS_DIR.mkdir(exist_ok=True)

    root = tk.Tk()
    root.title("Prod Reports Downloader")
    root.geometry("960x640")

    status_var = tk.StringVar(value="Click 'Refresh remote list' to load folders")
    progress_var = tk.StringVar(value="")
    remote_names: list[str] = []
    local_names: set[str] = set()
    partial_names: set[str] = set()
    filter_days: int | None = None
    download_thread: threading.Thread | None = None
    cancel_requested = threading.Event()
    is_downloading = tk.BooleanVar(value=False)
    overall_start = 0.0
    overall_expected: dict[tuple[str, str], int] = {}
    overall_downloaded: dict[tuple[str, str], int] = {}

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)

    top_row = ttk.Frame(main)
    top_row.pack(fill="x")

    ttk.Button(top_row, text="Refresh remote list", command=lambda: refresh()).pack(
        side="left", padx=(0, 6)
    )
    ttk.Button(
        top_row,
        text="Download selected",
        command=lambda: download_selected(),
        state="normal",
    ).pack(side="left")
    cancel_btn = ttk.Button(
        top_row, text="Cancel download", command=lambda: request_cancel(), state="disabled"
    )
    cancel_btn.pack(side="left", padx=(6, 6))
    ttk.Button(top_row, text="Close", command=root.destroy).pack(side="right")

    ttk.Label(main, textvariable=status_var).pack(anchor="w", pady=(6, 8))

    filter_row = ttk.Frame(main)
    filter_row.pack(fill="x", pady=(0, 8))
    ttk.Label(filter_row, text="Filter by age:").pack(side="left", padx=(0, 6))

    def set_filter(days: int | None) -> None:
        nonlocal filter_days
        filter_days = days
        apply_filter(select_after=True)

    for label, days in [
        ("7 days", 7),
        ("14 days", 14),
        ("30 days", 30),
        ("90 days", 90),
        ("180 days", 180),
        ("365 days", 365),
        ("All", None),
    ]:
        ttk.Button(filter_row, text=label, command=lambda d=days: set_filter(d)).pack(
            side="left", padx=(0, 4)
        )

    columns = ("folder", "range", "downloaded")
    tree = ttk.Treeview(main, columns=columns, show="headings", selectmode="extended")
    tree.heading("folder", text="Folder")
    tree.column("folder", width=280, anchor="w")
    tree.heading("range", text="Range")
    tree.column("range", width=260, anchor="w")
    tree.heading("downloaded", text="Downloaded?")
    tree.column("downloaded", width=120, anchor="center")

    y_scroll = ttk.Scrollbar(main, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=y_scroll.set)
    tree.pack(side="left", fill="both", expand=True)
    y_scroll.pack(side="right", fill="y")

    progress_frame = ttk.Frame(main)
    progress_frame.pack(fill="x", pady=(8, 0))
    progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
    progress_bar.pack(fill="x", expand=True)
    ttk.Label(progress_frame, textvariable=progress_var).pack(anchor="w")

    def format_range(name: str) -> str:
        start_dt, end_dt = parse_folder_name(name)
        if not start_dt or not end_dt:
            return "Unknown"
        return f"{start_dt:%Y-%m-%d %H:%M} → {end_dt:%Y-%m-%d %H:%M}"

    def apply_filter(select_after: bool = False) -> None:
        tree.delete(*tree.get_children())
        now = datetime.utcnow()
        filtered = []
        for name in remote_names:
            start_dt, _ = parse_folder_name(name)
            if filter_days is not None and start_dt:
                if now - start_dt > timedelta(days=filter_days):
                    continue
            filtered.append(name)

        if not filtered and not remote_names:
            for name in sorted(local_names | partial_names, key=folder_sort_key):
                status = "yes" if name in local_names else "no"
                tag = "downloaded" if status == "yes" else "local"
                tree.insert(
                    "",
                    "end",
                    values=(name, format_range(name), status),
                    tags=(tag,),
                )
            tree.tag_configure("local", foreground="gray")
            return

        for name in filtered:
            if name in local_names:
                downloaded = "yes"
                tag = "downloaded"
            else:
                downloaded = "no"
                tag = "remote"
            item_id = tree.insert(
                "", "end", values=(name, format_range(name), downloaded), tags=(tag,)
            )
            if select_after:
                tree.selection_add(item_id)
        tree.tag_configure("downloaded", foreground="green")
        tree.tag_configure("remote", foreground="black")

    def refresh(local_only: bool = False) -> None:
        status_var.set("Refreshing remote list...")
        root.update_idletasks()
        ok, names, message = fetch_remote_folders()
        complete, partial = local_folder_status()
        remote_names[:] = names if ok else []
        local_names.clear()
        local_names.update(complete)
        partial_names.clear()
        partial_names.update(partial)
        if not ok:
            status_var.set(message)
        else:
            status_var.set(message)
        apply_filter(select_after=False)

    def download_selected() -> None:
        nonlocal download_thread, overall_start
        items = tree.selection()
        if not items:
            messagebox.showinfo("No selection", "Select one or more folders first.")
            return
        if is_downloading.get():
            messagebox.showinfo("Busy", "A download is already in progress.")
            return

        selected_folders = []
        skipped = 0
        for item in items:
            folder = tree.item(item, "values")[0]
            if folder in local_names and folder not in partial_names:
                skipped += 1
                continue
            selected_folders.append(folder)

        if not selected_folders:
            status_var.set("All selected folders already downloaded; nothing to do.")
            return
        cancel_requested.clear()
        is_downloading.set(True)
        progress_bar["value"] = 0
        progress_bar["maximum"] = 100
        progress_var.set("Starting download...")
        overall_start = time.time()
        overall_expected.clear()
        overall_downloaded.clear()
        cancel_btn.configure(state="normal")

        def run_download() -> None:
            failures: list[str] = []
            successes = 0
            for folder in selected_folders:
                if cancel_requested.is_set():
                    failures.append(f"{folder}: Cancelled")
                    break

                def progress_hook(
                    fld: str,
                    filename: str,
                    downloaded: int,
                    total: int,
                    _elapsed: float,
                ) -> None:
                    key = (fld, filename)
                    prev = overall_downloaded.get(key, 0)
                    delta = max(0, downloaded - prev)
                    overall_downloaded[key] = downloaded
                    if total and key not in overall_expected:
                        overall_expected[key] = total
                    elapsed = max(time.time() - overall_start, 0.001)
                    total_expected = sum(overall_expected.values())
                    total_downloaded = sum(overall_downloaded.values())
                    percent = (
                        (total_downloaded / total_expected) * 100
                        if total_expected > 0
                        else 0
                    )
                    speed_mbps = (total_downloaded / 1_000_000) / elapsed

                    def update_ui() -> None:
                        progress_bar["value"] = percent
                        progress_var.set(
                            f"{percent:5.1f}% • {speed_mbps:0.2f} MB/s • {fld}/{filename}"
                        )

                    root.after(0, update_ui)

                ok, msg = download_folder(folder, progress_hook, cancel_requested.is_set)
                if ok:
                    successes += 1
                else:
                    failures.append(f"{folder}: {msg}")
                if cancel_requested.is_set():
                    break

            def finalize() -> None:
                nonlocal download_thread
                is_downloading.set(False)
                cancel_btn.configure(state="disabled")
                if successes:
                    msg = f"Downloaded {successes} folder(s)"
                    if skipped:
                        msg += f"; skipped {skipped} already complete"
                    status_var.set(msg)
                if failures:
                    messagebox.showerror("Download issues", "\n".join(failures))
                refresh()
                progress_var.set("Idle")
                progress_bar["value"] = 0
                download_thread = None

            root.after(0, finalize)

        download_thread = threading.Thread(target=run_download, daemon=True)
        download_thread.start()

    def request_cancel() -> None:
        if not is_downloading.get():
            return
        cancel_requested.set()
        progress_var.set("Cancelling...")
        status_var.set("Cancelling download...")

    def on_close() -> None:
        if is_downloading.get():
            if not messagebox.askyesno(
                "Cancel download", "A download is in progress. Cancel and close?"
            ):
                return
            request_cancel()
            # give the thread a moment to exit and clean up
            root.after(500, root.destroy)
            return
        root.destroy()

    refresh()
    progress_var.set("Idle")
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    build_ui()
