import os
import sys
import time
import subprocess
import tkinter as tk
from typing import Optional, Tuple


def load_env_file(path: str) -> None:
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip("\"").strip("'")
            os.environ.setdefault(key, value)


def read_status(path: str) -> Tuple[str, Optional[int]]:
    if not os.path.exists(path):
        return ("unknown", None)
    try:
        with open(path, "r", encoding="utf-8") as f:
            line = f.readline().strip()
        if "|" not in line:
            return (line or "unknown", None)
        state, ts = line.split("|", 1)
        return (state.strip(), int(ts.strip()))
    except Exception:
        return ("unknown", None)


class ControllerApp:
    def __init__(self, root: tk.Tk, base_dir: str) -> None:
        self.root = root
        self.base_dir = base_dir
        self.proc: Optional[subprocess.Popen] = None

        env_path = os.getenv("ENV_FILE", os.path.join(self.base_dir, ".env"))
        load_env_file(env_path)

        self.ocr_script = os.path.join(self.base_dir, "ocrStart.py")
        self.stop_file = os.getenv("STOP_FILE", "STOP")
        self.pid_file = os.getenv("PID_FILE", "ocr.pid")
        self.status_file = os.getenv("STATUS_FILE", "ocr.status")
        self.heartbeat_sec = float(os.getenv("HEARTBEAT_SEC", "5"))

        self.status_var = tk.StringVar(value="Status: unknown")
        self.log_var = tk.StringVar(value="")

        root.title("OCR Controller")
        root.geometry("420x180")

        status_label = tk.Label(root, textvariable=self.status_var, font=("Segoe UI", 12))
        status_label.pack(pady=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)
        self.start_btn = tk.Button(btn_frame, text="Start", width=10, command=self.start)
        self.stop_btn = tk.Button(btn_frame, text="Stop", width=10, command=self.stop)
        self.start_btn.grid(row=0, column=0, padx=8)
        self.stop_btn.grid(row=0, column=1, padx=8)

        log_label = tk.Label(root, textvariable=self.log_var, anchor="w", justify="left")
        log_label.pack(fill="x", padx=10, pady=10)

        self.root.after(500, self.refresh_status)

    def start(self) -> None:
        if self.is_running():
            self.log_var.set("Already running.")
            return
        if os.path.exists(self.stop_file):
            try:
                os.remove(self.stop_file)
            except Exception:
                pass
        self.proc = subprocess.Popen([sys.executable, self.ocr_script], cwd=self.base_dir)
        self.log_var.set("Started OCR process.")

    def stop(self) -> None:
        try:
            with open(self.stop_file, "w", encoding="utf-8") as f:
                f.write("stop\n")
        except Exception:
            pass
        if self.proc and self.proc.poll() is None:
            self.log_var.set("Stop requested. Waiting...")
        else:
            self.log_var.set("Stop requested.")

    def is_running(self) -> bool:
        state, ts = read_status(self.status_file)
        if state != "running" or ts is None:
            return False
        if time.time() - ts > self.heartbeat_sec * 3:
            return False
        return True

    def refresh_status(self) -> None:
        running = self.is_running()
        self.status_var.set("Status: running" if running else "Status: stopped")
        self.root.after(1000, self.refresh_status)


def main() -> None:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    root = tk.Tk()
    ControllerApp(root, base_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
