from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import ttk


BG = "#080c10"
CARD = "#0d1219"
PANEL = "#111a24"
GOLD = "#c09040"
GOLD_H = "#d4a84e"
FG = "#e8edf2"
MUTED = "#6e8096"
SUCCESS = "#4ade80"
ERROR = "#f87171"


def append_trace(trace_path: Path, message: str) -> None:
    trace_path.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with trace_path.open("a", encoding="utf-8") as handle:
        handle.write(f"[{timestamp}] {message}\n")


def apply_dark_titlebar(window: tk.Tk | tk.Toplevel) -> None:
    if sys.platform != "win32":
        return
    try:
        window.update_idletasks()
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        value = ctypes.c_int(1)
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value),
            ctypes.sizeof(value),
        )
    except Exception:
        pass


def wait_for_process_exit(pid: int, timeout_seconds: int) -> bool:
    if sys.platform != "win32":
        deadline = time.time() + timeout_seconds
        while time.time() < deadline:
            try:
                os.kill(pid, 0)
            except OSError:
                return True
            time.sleep(1)
        return False

    SYNCHRONIZE = 0x00100000
    WAIT_OBJECT_0 = 0x00000000
    WAIT_TIMEOUT = 0x00000102

    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(SYNCHRONIZE, False, pid)
    if not handle:
        return True
    try:
        result = kernel32.WaitForSingleObject(handle, timeout_seconds * 1000)
        return result == WAIT_OBJECT_0
    finally:
        kernel32.CloseHandle(handle)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


@dataclass
class ProgressEvent:
    kind: str
    message: str
    progress: int | None = None


class UpdateProgressWindow:
    def __init__(self, event_queue: "queue.Queue[ProgressEvent]") -> None:
        self.event_queue = event_queue
        self.root = tk.Tk()
        self.root.title("Updating Citizen StarString Helper")
        self.root.configure(bg=BG)
        self.root.resizable(False, False)
        self.root.geometry("520x228")
        self.root.minsize(520, 228)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Root.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Title.TLabel", background=CARD, foreground=FG, font=("Segoe UI Semibold", 13))
        style.configure("Body.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=PANEL, foreground=FG, font=("Segoe UI Semibold", 10))
        style.configure(
            "Update.Horizontal.TProgressbar",
            troughcolor=PANEL,
            bordercolor=PANEL,
            background=GOLD,
            lightcolor=GOLD_H,
            darkcolor=GOLD,
        )

        accent = tk.Frame(self.root, bg=GOLD, height=3)
        accent.pack(fill="x", side="top")

        outer = ttk.Frame(self.root, style="Root.TFrame", padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        card = ttk.Frame(outer, style="Card.TFrame", padding=18)
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)

        ttk.Label(card, text="Applying Update", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            card,
            text="Please keep this window open while the helper swaps files and relaunches the updated app.",
            style="Body.TLabel",
            wraplength=450,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 12))

        status_panel = tk.Frame(card, bg=PANEL, padx=12, pady=12)
        status_panel.grid(row=2, column=0, sticky="ew")
        status_panel.grid_columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="Preparing updater helper...")
        tk.Label(status_panel, textvariable=self.status_var, fg=FG, bg=PANEL, font=("Segoe UI Semibold", 10)).grid(
            row=0, column=0, sticky="w"
        )

        self.progress = ttk.Progressbar(
            card,
            style="Update.Horizontal.TProgressbar",
            orient="horizontal",
            mode="determinate",
            maximum=100,
            value=8,
            length=460,
        )
        self.progress.grid(row=3, column=0, sticky="ew", pady=(14, 10))

        self.detail_var = tk.StringVar(value="Initializing update workflow")
        ttk.Label(card, textvariable=self.detail_var, style="Body.TLabel").grid(row=4, column=0, sticky="w")

        apply_dark_titlebar(self.root)
        self._center()
        self.root.after(120, self._pump_queue)

    def _center(self) -> None:
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _pump_queue(self) -> None:
        close_delay = None
        try:
            while True:
                event = self.event_queue.get_nowait()
                if event.progress is not None:
                    self.progress["value"] = max(0, min(100, event.progress))
                if event.message:
                    self.status_var.set(event.message)
                    self.detail_var.set(event.message)
                if event.kind == "success":
                    self.status_var.set("Update complete. Launching app...")
                    self.detail_var.set(event.message)
                    self.progress["value"] = 100
                    close_delay = 1200
                elif event.kind == "error":
                    self.status_var.set("Update could not be completed")
                    self.detail_var.set(event.message)
                    close_delay = 3200
        except queue.Empty:
            pass

        if close_delay is not None:
            self.root.after(close_delay, self.root.destroy)
            return
        self.root.after(120, self._pump_queue)

    def run(self) -> None:
        self.root.mainloop()


def swap_files(current_exe: Path, new_exe: Path, backup_exe: Path, trace_path: Path, emit) -> bool:
    expected_size = new_exe.stat().st_size
    expected_hash = sha256(new_exe)

    for attempt in range(1, 41):
        emit("progress", f"Swapping files (attempt {attempt}/40)...", 56)
        append_trace(trace_path, f"Swap attempt {attempt}")
        try:
            if backup_exe.exists():
                backup_exe.unlink()
            if current_exe.exists():
                shutil.move(str(current_exe), str(backup_exe))
            shutil.move(str(new_exe), str(current_exe))

            actual_size = current_exe.stat().st_size
            actual_hash = sha256(current_exe)
            append_trace(trace_path, f"Swap verification size={actual_size} hash={actual_hash}")
            if actual_size == expected_size and actual_hash == expected_hash:
                emit("progress", "File swap verified.", 76)
                return True
        except Exception as exc:
            append_trace(trace_path, f"Swap attempt error: {exc}")
            emit("progress", f"Retrying file swap: {exc}", 56)

        time.sleep(1.5)

    return False


def launch_updated_app(current_exe: Path, trace_path: Path, emit) -> bool:
    try:
        emit("progress", "Preparing updated app...", 84)
        append_trace(trace_path, "Running warmup pass...")
        warmup = subprocess.run(
            [str(current_exe), "--warmup"],
            capture_output=True,
            text=True,
            timeout=60,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
        append_trace(trace_path, f"Warmup finished with code {warmup.returncode}.")
    except Exception as exc:
        append_trace(trace_path, f"Warmup failed: {exc}")
        emit("error", f"Warmup failed: {exc}", 88)
        return False

    time.sleep(2)

    try:
        emit("progress", "Launching updated app...", 94)
        append_trace(trace_path, "Launching updated app...")
        subprocess.Popen(
            [str(current_exe)],
            creationflags=(subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP) if sys.platform == "win32" else 0,
            close_fds=True,
        )
        append_trace(trace_path, "Updated app launch requested.")
        return True
    except Exception as exc:
        append_trace(trace_path, f"Launch failed: {exc}")
        emit("error", f"Launch failed: {exc}", 94)
        return False


def run_update(args, event_queue: "queue.Queue[ProgressEvent]") -> int:
    current_exe = Path(args.current_exe)
    new_exe = Path(args.new_exe)
    backup_exe = Path(args.backup_exe)
    result_file = Path(args.result_file)
    trace_path = Path(args.trace_file)

    def emit(kind: str, message: str, progress: int | None = None) -> None:
        event_queue.put(ProgressEvent(kind=kind, message=message, progress=progress))
        append_trace(trace_path, message)

    append_trace(trace_path, f"Updater helper started for v{args.version}. pid={args.pid}")
    emit("progress", f"Preparing update to v{args.version}...", 10)
    exited = wait_for_process_exit(args.pid, 120)
    append_trace(trace_path, f"Process wait finished. exited={exited}")
    if not exited:
        emit("error", "Timed out waiting for the old app to close.", 18)
        return 1

    emit("progress", "Old app closed. Validating staged files...", 28)
    time.sleep(1.0)

    try:
        if not new_exe.exists():
            append_trace(trace_path, f"New exe missing: {new_exe}")
            emit("error", "The downloaded update could not be found.", 30)
            return 1

        emit("progress", "Staged update located. Swapping files...", 44)
        if swap_files(current_exe, new_exe, backup_exe, trace_path, emit):
            payload = {
                "version": args.version,
                "applied_at": datetime.now().isoformat(timespec="seconds"),
                "target": str(current_exe),
            }
            result_file.write_text(json.dumps(payload), encoding="utf-8")
            append_trace(trace_path, "Wrote last_update.json.")
            launched = launch_updated_app(current_exe, trace_path, emit)
            append_trace(trace_path, f"Launch status after update: launched={launched}")
            if backup_exe.exists():
                backup_exe.unlink()
            if new_exe.exists():
                new_exe.unlink()
            if launched:
                emit("success", "Updated app launched successfully.", 100)
                append_trace(trace_path, "Update apply completed successfully.")
                return 0
            emit("error", "The updated app could not be launched.", 100)
            return 1

        append_trace(trace_path, "Swap failed - restoring backup if needed.")
        emit("progress", "Restore path engaged after unsuccessful swap...", 58)
        if backup_exe.exists() and not current_exe.exists():
            shutil.move(str(backup_exe), str(current_exe))
            append_trace(trace_path, "Restored backup executable.")
        emit("error", "The update could not be applied. Previous version restored.", 66)
        return 1
    except Exception as exc:
        append_trace(trace_path, f"Updater helper fatal error: {exc}")
        try:
            if backup_exe.exists() and not current_exe.exists():
                shutil.move(str(backup_exe), str(current_exe))
                append_trace(trace_path, "Restored backup after fatal error.")
        except Exception as restore_exc:
            append_trace(trace_path, f"Backup restore also failed: {restore_exc}")
        emit("error", f"Updater helper fatal error: {exc}", 70)
        return 1


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--current-exe", required=True)
    parser.add_argument("--new-exe", required=True)
    parser.add_argument("--backup-exe", required=True)
    parser.add_argument("--result-file", required=True)
    parser.add_argument("--trace-file", required=True)
    parser.add_argument("--version", required=True)
    parser.add_argument("--pid", required=True, type=int)
    args = parser.parse_args()

    event_queue: "queue.Queue[ProgressEvent]" = queue.Queue()
    result: dict[str, int] = {"code": 1}

    def worker() -> None:
        result["code"] = run_update(args, event_queue)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    window = UpdateProgressWindow(event_queue)
    window.run()
    thread.join(timeout=1)
    return result["code"]


if __name__ == "__main__":
    raise SystemExit(main())
