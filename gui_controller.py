import os
import sys
import time
import subprocess
import tkinter as tk
from typing import Optional, Tuple

try:
    import win32gui
    import win32con
except Exception:
    win32gui = None
    win32con = None


def load_env_file(path: str) -> None:
    # .env を常に上書き反映（GUIで変更した値を即座に使う）
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
            os.environ[key] = value


def read_env_map(path: str) -> dict:
    # 子プロセス起動時に渡すため、最新の .env を辞書化する
    data = {}
    if not os.path.exists(path):
        return data
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
            data[key] = value
    return data


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


def read_log(path: str, max_chars: int) -> str:
    if not os.path.exists(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = f.read()
        if len(data) > max_chars:
            return data[-max_chars:]
        return data
    except Exception:
        return ""


def read_env_value(path: str, key: str) -> str:
    if not os.path.exists(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                k, v = line.split("=", 1)
                if k.strip() == key:
                    return v.strip()
    except Exception:
        return ""
    return ""

def update_env_value(path: str, key: str, value: str) -> None:
    lines = []
    found = False
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.rstrip("\n")
                stripped = line.strip()
                if stripped.startswith(f"{key}=") or stripped.startswith(f"#{key}="):
                    lines.append(f"{key}={value}")
                    found = True
                else:
                    lines.append(line)
    if not found:
        lines.append(f"{key}={value}")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def _find_best_window(title: str) -> int:
    if win32gui is None:
        raise RuntimeError("pywin32 is required")

    candidates = []

    def enum_handler(h, result):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            if title in t:
                result.append(h)

    results = []
    win32gui.EnumWindows(enum_handler, results)
    if not results:
        # fallback to exact match
        hwnd = win32gui.FindWindow(None, title)
        return hwnd if hwnd else 0

    # 最も大きいウィンドウを選択（小さい子ウィンドウ誤検出を防ぐ）
    best_hwnd = 0
    best_area = -1
    for h in results:
        l, t, r, b = win32gui.GetWindowRect(h)
        area = max(0, r - l) * max(0, b - t)
        if area > best_area:
            best_area = area
            best_hwnd = h
    return best_hwnd


def get_client_rect(title: str) -> Tuple[int, int, int, int]:
    if win32gui is None:
        raise RuntimeError("pywin32 is required")
    hwnd = _find_best_window(title)
    if not hwnd:
        raise RuntimeError(f"Window not found: {title}")
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    pt_left_top = win32gui.ClientToScreen(hwnd, (left, top))
    pt_right_bottom = win32gui.ClientToScreen(hwnd, (right, bottom))
    return (pt_left_top[0], pt_left_top[1], pt_right_bottom[0], pt_right_bottom[1])


def get_client_rect_by_hwnd(hwnd: int) -> Tuple[int, int, int, int]:
    if win32gui is None:
        raise RuntimeError("pywin32 is required")
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    pt_left_top = win32gui.ClientToScreen(hwnd, (left, top))
    pt_right_bottom = win32gui.ClientToScreen(hwnd, (right, bottom))
    return (pt_left_top[0], pt_left_top[1], pt_right_bottom[0], pt_right_bottom[1])


def get_window_rect_by_hwnd(hwnd: int) -> Tuple[int, int, int, int]:
    if win32gui is None:
        raise RuntimeError("pywin32 is required")
    return win32gui.GetWindowRect(hwnd)


def enable_dpi_awareness() -> None:
    try:
        import ctypes
        shcore = ctypes.WinDLL("shcore")
        shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            import ctypes
            user32 = ctypes.WinDLL("user32")
            user32.SetProcessDPIAware()
        except Exception:
            pass


class ControllerApp:
    def __init__(self, root: tk.Tk, base_dir: str) -> None:
        self.root = root
        self.base_dir = base_dir
        self.proc: Optional[subprocess.Popen] = None

        self.env_path = os.getenv("ENV_FILE", os.path.join(self.base_dir, ".env"))
        load_env_file(self.env_path)

        self.ocr_script = os.path.join(self.base_dir, "ocrStart.py")
        self.stop_file = os.getenv("STOP_FILE", "STOP")
        self.pid_file = os.getenv("PID_FILE", "ocr.pid")
        self.status_file = os.getenv("STATUS_FILE", "ocr.status")
        self.heartbeat_sec = float(os.getenv("HEARTBEAT_SEC", "5"))
        self.log_file = os.getenv("LOG_FILE", "ocr.latest.txt")
        self.log_max_chars = int(os.getenv("LOG_MAX_CHARS", "800"))
        self.recent_log_file = os.getenv("RECENT_LOG_FILE", "ocr.recent.txt")
        self.recent_log_lines = int(os.getenv("RECENT_LOG_LINES", "200"))
        self.webhook_log_file = os.getenv("WEBHOOK_LOG_FILE", "ocr.webhook.txt")
        self.webhook_log_lines = int(os.getenv("WEBHOOK_LOG_LINES", "200"))
        self.window_title = os.getenv("LINE_WINDOW_TITLE", "LINE")

        def resolve_path(p: str) -> str:
            if not p:
                return p
            return p if os.path.isabs(p) else os.path.join(self.base_dir, p)

        self.stop_file = resolve_path(self.stop_file)
        self.pid_file = resolve_path(self.pid_file)
        self.status_file = resolve_path(self.status_file)
        self.log_file = resolve_path(self.log_file)
        self.recent_log_file = resolve_path(self.recent_log_file)
        self.webhook_log_file = resolve_path(self.webhook_log_file)

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
        self.measure_btn = tk.Button(btn_frame, text="Frame", width=10, command=self.measure)
        self.test_btn = tk.Button(btn_frame, text="Test", width=10, command=self.test_send)
        self.settings_btn = tk.Button(btn_frame, text="Settings", width=10, command=self.open_settings)
        self.start_btn.grid(row=0, column=0, padx=8)
        self.stop_btn.grid(row=0, column=1, padx=8)
        self.measure_btn.grid(row=0, column=2, padx=8)
        self.test_btn.grid(row=0, column=3, padx=8)
        self.settings_btn.grid(row=0, column=4, padx=8)

        self.crop_var = tk.StringVar(value="CROP_RECT: (not set)")
        crop_label = tk.Label(root, textvariable=self.crop_var, anchor="w")
        crop_label.pack(fill="x", padx=10)

        log_title = tk.Label(root, text="Latest OCR", anchor="w")
        log_title.pack(fill="x", padx=10)
        self.log_text = tk.Text(root, height=4, wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=4)
        self.log_text.configure(state="disabled")

        recent_title = tk.Label(root, text="Recent OCR (last 200 lines)", anchor="w")
        recent_title.pack(fill="x", padx=10)
        self.recent_text = tk.Text(root, height=6, wrap="word")
        self.recent_text.pack(fill="both", expand=True, padx=10, pady=6)
        self.recent_text.configure(state="disabled")

        webhook_title = tk.Label(root, text="Discord Webhook Log (last 200 lines)", anchor="w")
        webhook_title.pack(fill="x", padx=10)
        self.webhook_text = tk.Text(root, height=4, wrap="word")
        self.webhook_text.pack(fill="both", expand=True, padx=10, pady=6)
        self.webhook_text.configure(state="disabled")

        self.root.after(500, self.refresh_status)

    def start(self) -> None:
        # 最新の .env を子プロセスへ渡す（CROP更新を即反映）
        if self.is_running():
            self.log_var.set("Already running.")
            return
        if os.path.exists(self.stop_file):
            try:
                os.remove(self.stop_file)
            except Exception:
                pass
        env = os.environ.copy()
        env.update(read_env_map(self.env_path))
        self.proc = subprocess.Popen([sys.executable, self.ocr_script], cwd=self.base_dir, env=env)
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

    def measure(self) -> None:
        # 画面座標ベースのフレーム指定（LINEに依存しない）
        if win32gui is None:
            self.log_var.set("pywin32 is required for measure.")
            return

        self._open_frame_selector()

    def _open_frame_selector(self) -> None:
        # 透明フレームの位置・サイズを画面座標で保存する
        frame_win = tk.Toplevel(self.root)
        frame_win.attributes("-topmost", False)
        frame_win.attributes("-alpha", 0.25)
        frame_win.title("Resize this frame over capture area")
        frame_win.geometry("400x300+100+100")
        frame_win.configure(bg="black")
        frame_win.minsize(200, 150)

        canvas = tk.Canvas(frame_win, bg="black", highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        rect_id = canvas.create_rectangle(2, 2, 398, 298, outline="red", width=3)

        # Opaque control window (header/coords/apply)
        control_win = tk.Toplevel(self.root)
        control_win.attributes("-topmost", True)
        control_win.attributes("-alpha", 1.0)
        control_win.overrideredirect(False)
        control_win.title("Frame Controls")

        coord_var = tk.StringVar(value="CROP_RECT=-, -,-,-")
        coord_label = tk.Label(control_win, textvariable=coord_var, bg="white", fg="black")
        coord_label.pack(fill="x", padx=8, pady=6)
        status_var = tk.StringVar(value="")
        status_label = tk.Label(control_win, textvariable=status_var, bg="white", fg="red")
        status_label.pack(fill="x", padx=8, pady=(0, 6))
        apply_btn = tk.Button(control_win, text="Apply")
        apply_btn.pack(padx=8, pady=6)

        def on_resize(_event):
            w = frame_win.winfo_width()
            h = frame_win.winfo_height()
            canvas.coords(rect_id, 2, 2, max(3, w - 2), max(3, h - 2))
            update_coords()

        def update_coords():
            try:
                hwnd = frame_win.winfo_id()
                f_left, f_top, f_right, f_bottom = get_client_rect_by_hwnd(hwnd)
                coord_var.set(f"CROP_RECT={f_left},{f_top},{f_right},{f_bottom}")
                control_win.update_idletasks()
                ch = control_win.winfo_height()
                margin = 8
                y_above = f_top - ch - margin
                y_below = f_bottom + margin
                y = y_above if (y_above + ch) <= f_top else y_below
                control_win.geometry(f"+{f_left}+{y}")
                control_win.lift()
                control_win.focus_force()
            except Exception:
                coord_var.set("CROP_RECT=-, -,-,-")

        def on_move(_event):
            update_coords()

        def apply_rect():
            try:
                hwnd = frame_win.winfo_id()
                f_left, f_top, f_right, f_bottom = get_client_rect_by_hwnd(hwnd)
            except Exception:
                self.log_var.set("フレーム座標の取得に失敗しました。")
                status_var.set("フレーム座標の取得に失敗しました。")
                return

            if (f_right - f_left) < 10 or (f_bottom - f_top) < 10:
                self.log_var.set("範囲が小さすぎます。")
                status_var.set("範囲が小さすぎます。")
                return

            update_env_value(self.env_path, "CROP_RECT", f"{f_left},{f_top},{f_right},{f_bottom}")
            # 画面絶対座標として扱う
            update_env_value(self.env_path, "CROP_MODE", "screen")
            self.log_var.set("更新されました！")
            status_var.set("更新されました！")
            frame_win.destroy()
            control_win.destroy()

        apply_btn.configure(command=apply_rect)

        def tick():
            update_coords()
            frame_win.after(200, tick)

        frame_win.bind("<Configure>", on_resize)
        canvas.bind("<Motion>", on_move)
        frame_win.bind("<Escape>", lambda _e: (frame_win.destroy(), control_win.destroy()))
        control_win.bind("<Escape>", lambda _e: (frame_win.destroy(), control_win.destroy()))
        update_coords()
        tick()

    def test_send(self) -> None:
        webhook_url = read_env_value(self.env_path, "WEBHOOK_URL")
        if not webhook_url:
            self.log_var.set("WEBHOOK_URL not set.")
            return
        try:
            import requests
            payload = {"content": "test: webhook ok"}
            resp = requests.post(webhook_url, json=payload, timeout=10)
            if resp.status_code >= 400:
                self.log_var.set(f"Test failed: {resp.status_code}")
            else:
                self.log_var.set("Test sent.")
        except Exception:
            self.log_var.set("Test send error.")

    def open_settings(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.geometry("420x520")
        win.attributes("-topmost", True)

        def get(key: str, default: str = "") -> str:
            v = read_env_value(self.env_path, key)
            return v if v != "" else default

        # Variables
        var_poll = tk.StringVar(value=get("POLL_SEC", "1.0"))
        var_lang = tk.StringVar(value=get("OCR_LANG", "jpn+eng"))
        var_scale = tk.StringVar(value=get("OCR_SCALE", "1.0"))
        var_thresh = tk.StringVar(value=get("THRESHOLD", "160"))
        var_median = tk.StringVar(value=get("MEDIAN_SIZE", "0"))
        var_keepnl = tk.BooleanVar(value=get("KEEP_NEWLINES", "false").lower() in {"1","true","yes","on"})
        var_pre = tk.BooleanVar(value=get("PREPROCESS", "false").lower() in {"1","true","yes","on"})
        var_sharp = tk.BooleanVar(value=get("SHARPEN", "false").lower() in {"1","true","yes","on"})
        var_ts = tk.BooleanVar(value=get("ADD_TIMESTAMP", "true").lower() in {"1","true","yes","on"})
        var_once = tk.BooleanVar(value=get("SAVE_SCREENSHOT_ONCE", "false").lower() in {"1","true","yes","on"})
        var_layout = tk.BooleanVar(value=get("SAVE_LAYOUT_TSV", "false").lower() in {"1","true","yes","on"})
        var_save = tk.BooleanVar(value=get("SAVE_SCREENSHOT", "false").lower() in {"1","true","yes","on"})
        var_only_change = tk.BooleanVar(value=get("ONLY_ON_CHANGE", "true").lower() in {"1","true","yes","on"})
        var_webhook = tk.StringVar(value=get("WEBHOOK_URL", ""))
        var_invert = tk.BooleanVar(value=get("INVERT", "false").lower() in {"1","true","yes","on"})
        var_psm = tk.StringVar(value=get("PSM", "6"))

        row = 0
        def add_label(text: str):
            nonlocal row
            tk.Label(win, text=text, anchor="w").grid(row=row, column=0, sticky="w", padx=10, pady=6)
            row += 1

        def add_entry(label: str, var: tk.StringVar):
            nonlocal row
            tk.Label(win, text=label, anchor="w").grid(row=row, column=0, sticky="w", padx=10)
            tk.Entry(win, textvariable=var, width=24).grid(row=row, column=1, padx=10)
            row += 1

        def add_check(label: str, var: tk.BooleanVar):
            nonlocal row
            tk.Checkbutton(win, text=label, variable=var).grid(row=row, column=0, columnspan=2, sticky="w", padx=10)
            row += 1

        add_label("OCR")
        add_entry("間隔(POLL_SEC) [推奨:10]", var_poll)
        add_entry("言語(OCR_LANG) [推奨:jpn+eng]", var_lang)
        add_entry("拡大率(OCR_SCALE) [推奨:2.5]", var_scale)
        add_entry("閾値(THRESHOLD) [推奨:180]", var_thresh)
        add_entry("ノイズ除去(MEDIAN_SIZE) [推奨:0]", var_median)
        add_check("前処理(PREPROCESS) [推奨:true]", var_pre)
        add_check("反転(INVERT) [推奨:true]", var_invert)
        add_check("シャープ(SHARPEN) [推奨:true]", var_sharp)
        add_check("改行維持(KEEP_NEWLINES) [推奨:true]", var_keepnl)
        add_entry("PSM(PSM) [推奨:6]", var_psm)
        add_check("時刻付与(ADD_TIMESTAMP)", var_ts)

        add_label("Debug/Tools")
        add_check("一回スクショ(SAVE_SCREENSHOT_ONCE)", var_once)
        add_check("常時スクショ(SAVE_SCREENSHOT)", var_save)
        add_check("TSV保存(SAVE_LAYOUT_TSV)", var_layout)

        add_label("送信")
        add_check("変化時のみ(ONLY_ON_CHANGE) [推奨:true]", var_only_change)
        add_entry("Webhook(URL)(WEBHOOK_URL)", var_webhook)

        def save_settings():
            update_env_value(self.env_path, "POLL_SEC", var_poll.get())
            update_env_value(self.env_path, "OCR_LANG", var_lang.get())
            update_env_value(self.env_path, "OCR_SCALE", var_scale.get())
            update_env_value(self.env_path, "THRESHOLD", var_thresh.get())
            update_env_value(self.env_path, "MEDIAN_SIZE", var_median.get())
            update_env_value(self.env_path, "PREPROCESS", str(var_pre.get()).lower())
            update_env_value(self.env_path, "INVERT", str(var_invert.get()).lower())
            update_env_value(self.env_path, "SHARPEN", str(var_sharp.get()).lower())
            update_env_value(self.env_path, "KEEP_NEWLINES", str(var_keepnl.get()).lower())
            update_env_value(self.env_path, "ADD_TIMESTAMP", str(var_ts.get()).lower())
            update_env_value(self.env_path, "SAVE_SCREENSHOT_ONCE", str(var_once.get()).lower())
            update_env_value(self.env_path, "SAVE_LAYOUT_TSV", str(var_layout.get()).lower())
            update_env_value(self.env_path, "SAVE_SCREENSHOT", str(var_save.get()).lower())
            update_env_value(self.env_path, "ONLY_ON_CHANGE", str(var_only_change.get()).lower())
            update_env_value(self.env_path, "WEBHOOK_URL", var_webhook.get())
            update_env_value(self.env_path, "PSM", var_psm.get())
            self.log_var.set("更新されました！")
            win.destroy()

        tk.Button(win, text="Save", command=save_settings).grid(row=row, column=0, columnspan=2, pady=12)

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
        crop_val = read_env_value(self.env_path, "CROP_RECT")
        if hasattr(self, "_last_crop_val"):
            if crop_val and crop_val != self._last_crop_val:
                self.log_var.set("CROP_RECT updated.")
        self._last_crop_val = crop_val
        if crop_val:
            self.crop_var.set(f"CROP_RECT: {crop_val}")
        else:
            self.crop_var.set("CROP_RECT: (not set)")
        latest = read_log(self.log_file, self.log_max_chars)
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.insert("1.0", latest)
        self.log_text.configure(state="disabled")

        recent = read_log(self.recent_log_file, self.recent_log_lines * 200)
        if recent:
            recent_lines = recent.splitlines()[-self.recent_log_lines:]
            recent = "\n".join(recent_lines)
        self.recent_text.configure(state="normal")
        self.recent_text.delete("1.0", "end")
        self.recent_text.insert("1.0", recent)
        self.recent_text.configure(state="disabled")

        webhook = read_log(self.webhook_log_file, self.webhook_log_lines * 200)
        if webhook:
            webhook_lines = webhook.splitlines()[-self.webhook_log_lines:]
            webhook = "\n".join(webhook_lines)
        self.webhook_text.configure(state="normal")
        self.webhook_text.delete("1.0", "end")
        self.webhook_text.insert("1.0", webhook)
        self.webhook_text.configure(state="disabled")
        self.root.after(1000, self.refresh_status)


def main() -> None:
    enable_dpi_awareness()
    base_dir = os.path.dirname(os.path.abspath(__file__))
    root = tk.Tk()
    ControllerApp(root, base_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
