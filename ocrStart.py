# OCR -> Discord Webhook (LINE PC OpenChat)
# Minimal, production-ready-ish scaffold with pluggable sender.
# ASCII-only comments, Japanese text can be OCR'ed.

import os
import time
import re
from dataclasses import dataclass
from typing import Optional, Tuple
from datetime import datetime
import signal

import requests
from PIL import Image, ImageOps, ImageFilter
import pytesseract
from pytesseract import Output
import mss

try:
    import win32gui
    import win32con
except Exception:
    win32gui = None
    win32con = None


@dataclass
class Config:
    window_title: str
    webhook_url: str
    poll_sec: float = 1.0
    ocr_lang: str = "jpn+eng"
    # Optional fixed crop (left, top, right, bottom) relative to window.
    crop_rect: Optional[Tuple[int, int, int, int]] = None
    # Crop mode: "line" (relative to LINE client) or "screen" (absolute screen coords)
    crop_mode: str = "line"
    # If True, only send when new text appears.
    only_on_change: bool = True
    # Reduce noise by collapsing whitespace
    normalize_whitespace: bool = True
    # Keep newlines if True
    keep_newlines: bool = False
    # Preprocess image before OCR
    preprocess: bool = False
    # Threshold for binarization when preprocess is enabled
    threshold: int = 160
    # Add timestamp to message
    add_timestamp: bool = True
    # Save screenshots for crop tuning
    save_screenshot: bool = False
    # Save only once and exit
    save_screenshot_once: bool = False
    # Output directory for screenshots
    screenshot_dir: str = "screenshots"
    # Stop if this file exists
    stop_file: str = "STOP"
    # PID file path
    pid_file: str = "ocr.pid"
    # Status file path
    status_file: str = "ocr.status"
    # Heartbeat interval seconds
    heartbeat_sec: float = 5.0
    # Latest OCR output file
    log_file: str = "ocr.latest.txt"
    # Max chars to keep in latest log
    log_max_chars: int = 800
    # Keep a rolling log of recent OCR outputs
    recent_log_file: str = "ocr.recent.txt"
    # Max lines to keep in recent log
    recent_log_lines: int = 200
    # Webhook log file
    webhook_log_file: str = "ocr.webhook.txt"
    # Max lines to keep in webhook log
    webhook_log_lines: int = 200
    # Preprocess scale factor
    ocr_scale: float = 1.0
    # Median filter size (0 to disable)
    median_size: int = 0
    # Apply sharpen filter
    sharpen: bool = False
    # Invert colors before preprocessing
    invert: bool = False
    # Tesseract PSM
    psm: int = 6
    # Save layout OCR TSV
    save_layout_tsv: bool = False
    # Directory for TSV output
    layout_tsv_dir: str = "layout_tsv"


class Sender:
    def send(self, text: str) -> None:
        raise NotImplementedError


class WebhookSender(Sender):
    def __init__(self, webhook_url: str, log_path: Optional[str] = None, log_lines: int = 200):
        self.webhook_url = webhook_url
        self.log_path = log_path
        self.log_lines = log_lines

    def send(self, text: str) -> None:
        # Discord webhook message max ~2000 chars
        if not text:
            return
        if len(text) > 1900:
            text = text[:1900] + "..."
        resp = requests.post(self.webhook_url, json={"content": text})
        if self.log_path:
            msg = f"status={resp.status_code}"
            if resp.status_code >= 400:
                body = resp.text.replace("\n", " ")
                if len(body) > 200:
                    body = body[:200] + "..."
                msg += f" error={body}"
            append_recent_log(self.log_path, msg, self.log_lines)
        if resp.status_code >= 400:
            raise RuntimeError(f"Webhook failed: {resp.status_code} {resp.text}")


def load_env_file(path: str) -> None:
    # GUIで更新した .env を常に優先して反映する
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


def parse_crop_rect(value: str) -> Optional[Tuple[int, int, int, int]]:
    if not value:
        return None
    parts = [p.strip() for p in value.split(",")]
    if len(parts) != 4:
        return None
    try:
        nums = [int(p) for p in parts]
    except ValueError:
        return None
    return (nums[0], nums[1], nums[2], nums[3])


def parse_bool(value: str, default: bool) -> bool:
    if value == "":
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def get_window_rect(title: str) -> Tuple[int, int, int, int]:
    if win32gui is None:
        raise RuntimeError("pywin32 is required for window capture")

    hwnd = _find_best_window(title)

    if not hwnd:
        raise RuntimeError(f"Window not found: {title}")

    # Restore if minimized
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    rect = win32gui.GetWindowRect(hwnd)
    # rect is (left, top, right, bottom)
    return rect

def _find_best_window(title: str) -> int:
    if win32gui is None:
        return 0
    candidates = []

    def enum_handler(h, result):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            if title in t:
                result.append(h)

    results = []
    win32gui.EnumWindows(enum_handler, results)
    if not results:
        hwnd = win32gui.FindWindow(None, title)
        return hwnd if hwnd else 0

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
        raise RuntimeError("pywin32 is required for window capture")

    hwnd = _find_best_window(title)

    if not hwnd:
        raise RuntimeError(f"Window not found: {title}")

    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    pt_left_top = win32gui.ClientToScreen(hwnd, (left, top))
    pt_right_bottom = win32gui.ClientToScreen(hwnd, (right, bottom))
    return (pt_left_top[0], pt_left_top[1], pt_right_bottom[0], pt_right_bottom[1])


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


def grab_window_image(rect: Tuple[int, int, int, int]) -> Image.Image:
    left, top, right, bottom = rect
    with mss.mss() as sct:
        # 仮想スクリーン範囲にクランプしてGDIエラーを回避
        v = sct.monitors[0]
        v_left, v_top = v["left"], v["top"]
        v_right = v_left + v["width"]
        v_bottom = v_top + v["height"]

        left = max(left, v_left)
        top = max(top, v_top)
        right = min(right, v_right)
        bottom = min(bottom, v_bottom)

        width = right - left
        height = bottom - top
        if width <= 0 or height <= 0:
            raise RuntimeError(f"Invalid capture rect: {rect}")

        monitor = {"left": left, "top": top, "width": width, "height": height}
        sct_img = sct.grab(monitor)
        img = Image.frombytes("RGB", sct_img.size, sct_img.rgb)
    return img


def apply_crop(img: Image.Image, crop_rect: Optional[Tuple[int, int, int, int]]) -> Image.Image:
    if not crop_rect:
        return img
    left, top, right, bottom = crop_rect
    # Clamp to image bounds to avoid invalid crop
    left = max(0, min(left, img.width - 1))
    top = max(0, min(top, img.height - 1))
    right = max(left + 1, min(right, img.width))
    bottom = max(top + 1, min(bottom, img.height))
    return img.crop((left, top, right, bottom))


def preprocess_image(img: Image.Image, threshold: int, invert: bool) -> Image.Image:
    if img.mode != "RGB":
        img = img.convert("RGB")
    # 黒背景向けの反転（白背景ならfalse推奨）
    if invert:
        img = ImageOps.invert(img)
    gray = ImageOps.grayscale(img)
    # Auto-contrast to improve separation
    gray = ImageOps.autocontrast(gray)
    # Simple binary threshold
    bw = gray.point(lambda x: 255 if x > threshold else 0, mode="1")
    return bw.convert("L")


def ocr_image(img: Image.Image, lang: str, psm: int) -> str:
    # PSMは文字配置の仮定（6=ブロック本文）
    config = f"--psm {psm}"
    text = pytesseract.image_to_string(img, lang=lang, config=config)
    return text


def normalize_text(text: str, keep_newlines: bool) -> str:
    text = text.strip()
    if keep_newlines:
        # collapse multiple blank lines
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text
    # collapse whitespace to single spaces
    text = re.sub(r"\s+", " ", text)
    return text


def format_message(text: str, add_timestamp: bool) -> str:
    if not add_timestamp:
        return text
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return f"[{ts}]\n{text}"


def write_latest_text(path: str, text: str, max_chars: int) -> None:
    if not text:
        return
    if len(text) > max_chars:
        text = text[:max_chars] + "..."
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"[{ts}]\n{text}\n")


def append_recent_log(path: str, text: str, max_lines: int) -> None:
    if not text:
        return
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{ts}] {text}".replace("\n", " ")
    lines = []
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
        except Exception:
            lines = []
    lines.append(entry)
    if len(lines) > max_lines:
        lines = lines[-max_lines:]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def save_layout_tsv(img: Image.Image, lang: str, out_dir: str) -> None:
    os.makedirs(out_dir, exist_ok=True)
    tsv = pytesseract.image_to_data(img, lang=lang, output_type=Output.STRING)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(out_dir, f"layout_{ts}.tsv")
    with open(path, "w", encoding="utf-8") as f:
        f.write(tsv)


def save_screenshots(img_full: Image.Image, img_crop: Image.Image, out_dir: str) -> None:
    os.makedirs(out_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    full_path = os.path.join(out_dir, f"line_full_{ts}.png")
    crop_path = os.path.join(out_dir, f"line_crop_{ts}.png")
    img_full.save(full_path)
    img_crop.save(crop_path)
    print(f"Saved screenshots: {full_path} / {crop_path}")


def write_status(path: str, state: str) -> None:
    ts = int(time.time())
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"{state}|{ts}\n")


def main() -> None:
    stop_requested = False

    def _handle_signal(_signum, _frame):
        nonlocal stop_requested
        stop_requested = True

    signal.signal(signal.SIGINT, _handle_signal)
    signal.signal(signal.SIGTERM, _handle_signal)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    enable_dpi_awareness()

    env_path = os.getenv("ENV_FILE", os.path.join(base_dir, ".env"))
    load_env_file(env_path)

    tesseract_cmd = os.getenv("TESSERACT_CMD", "")
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    cfg = Config(
        window_title=os.getenv("LINE_WINDOW_TITLE", "LINE"),
        webhook_url=os.getenv("WEBHOOK_URL", ""),
        poll_sec=float(os.getenv("POLL_SEC", "1.0")),
        ocr_lang=os.getenv("OCR_LANG", "jpn+eng"),
        crop_rect=parse_crop_rect(os.getenv("CROP_RECT", "")),
        crop_mode=os.getenv("CROP_MODE", "line").lower(),
        keep_newlines=parse_bool(os.getenv("KEEP_NEWLINES", ""), False),
        preprocess=parse_bool(os.getenv("PREPROCESS", ""), False),
        threshold=int(os.getenv("THRESHOLD", "160")),
        add_timestamp=parse_bool(os.getenv("ADD_TIMESTAMP", ""), True),
        only_on_change=parse_bool(os.getenv("ONLY_ON_CHANGE", ""), True),
        save_screenshot=parse_bool(os.getenv("SAVE_SCREENSHOT", ""), False),
        save_screenshot_once=parse_bool(os.getenv("SAVE_SCREENSHOT_ONCE", ""), False),
        screenshot_dir=os.getenv("SCREENSHOT_DIR", "screenshots"),
        stop_file=os.getenv("STOP_FILE", "STOP"),
        pid_file=os.getenv("PID_FILE", "ocr.pid"),
        status_file=os.getenv("STATUS_FILE", "ocr.status"),
        heartbeat_sec=float(os.getenv("HEARTBEAT_SEC", "5")),
        log_file=os.getenv("LOG_FILE", "ocr.latest.txt"),
        log_max_chars=int(os.getenv("LOG_MAX_CHARS", "800")),
        recent_log_file=os.getenv("RECENT_LOG_FILE", "ocr.recent.txt"),
        recent_log_lines=int(os.getenv("RECENT_LOG_LINES", "200")),
        webhook_log_file=os.getenv("WEBHOOK_LOG_FILE", "ocr.webhook.txt"),
        webhook_log_lines=int(os.getenv("WEBHOOK_LOG_LINES", "200")),
        ocr_scale=float(os.getenv("OCR_SCALE", "1.0")),
        median_size=int(os.getenv("MEDIAN_SIZE", "0")),
        sharpen=parse_bool(os.getenv("SHARPEN", ""), False),
        invert=parse_bool(os.getenv("INVERT", ""), False),
        psm=int(os.getenv("PSM", "6")),
        save_layout_tsv=parse_bool(os.getenv("SAVE_LAYOUT_TSV", ""), False),
        layout_tsv_dir=os.getenv("LAYOUT_TSV_DIR", "layout_tsv"),
    )

    def resolve_path(p: str) -> str:
        if not p:
            return p
        return p if os.path.isabs(p) else os.path.join(base_dir, p)

    cfg.screenshot_dir = resolve_path(cfg.screenshot_dir)
    cfg.stop_file = resolve_path(cfg.stop_file)
    cfg.pid_file = resolve_path(cfg.pid_file)
    cfg.status_file = resolve_path(cfg.status_file)
    cfg.log_file = resolve_path(cfg.log_file)
    cfg.recent_log_file = resolve_path(cfg.recent_log_file)
    cfg.webhook_log_file = resolve_path(cfg.webhook_log_file)
    cfg.layout_tsv_dir = resolve_path(cfg.layout_tsv_dir)

    if not cfg.webhook_url:
        raise RuntimeError("WEBHOOK_URL env var is required")

    sender = WebhookSender(cfg.webhook_url, cfg.webhook_log_file, cfg.webhook_log_lines)

    last_text = ""
    saved_once = False
    last_heartbeat = 0.0

    try:
        with open(cfg.pid_file, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        write_status(cfg.status_file, "running")

        print("Starting OCR loop...")
        while True:
            if stop_requested:
                print("Stop requested by signal. Exiting.")
                break
            if cfg.stop_file and os.path.exists(cfg.stop_file):
                print(f"Stop file detected: {cfg.stop_file}. Exiting.")
                break

            # screen: 画面絶対座標 / line: LINEクライアント相対
            if cfg.crop_rect and cfg.crop_mode == "screen":
                rect = cfg.crop_rect
                img_full = grab_window_image(rect)
                img_crop = img_full
            else:
                client_rect = get_client_rect(cfg.window_title)
                cl, ct, cr, cb = client_rect
                if cfg.crop_rect:
                    l, t, r, b = cfg.crop_rect
                    # clamp crop rect to client size
                    max_w = cr - cl
                    max_h = cb - ct
                    l = max(0, min(l, max_w - 1))
                    t = max(0, min(t, max_h - 1))
                    r = max(l + 1, min(r, max_w))
                    b = max(t + 1, min(b, max_h))
                    rect = (cl + l, ct + t, cl + r, ct + b)
                    if rect[2] <= rect[0] or rect[3] <= rect[1]:
                        print(f"Invalid crop rect: {cfg.crop_rect}")
                        time.sleep(cfg.poll_sec)
                        continue
                    img_full = grab_window_image(rect)
                    img_crop = img_full
                else:
                    img_full = grab_window_image(client_rect)
                    img_crop = img_full

            if cfg.ocr_scale and cfg.ocr_scale > 1.0:
                w, h = img_crop.size
                img_crop = img_crop.resize(
                    (int(w * cfg.ocr_scale), int(h * cfg.ocr_scale)),
                    Image.LANCZOS,
                )
            if cfg.median_size and cfg.median_size > 1:
                img_crop = img_crop.filter(ImageFilter.MedianFilter(size=cfg.median_size))
            if cfg.sharpen:
                img_crop = img_crop.filter(ImageFilter.SHARPEN)
            if cfg.preprocess:
                img_ocr = preprocess_image(img_crop, cfg.threshold, cfg.invert)
            else:
                img_ocr = img_crop

            if cfg.save_layout_tsv:
                save_layout_tsv(img_ocr, cfg.ocr_lang, cfg.layout_tsv_dir)

            text = ocr_image(img_ocr, cfg.ocr_lang, cfg.psm)
            if cfg.normalize_whitespace:
                text = normalize_text(text, cfg.keep_newlines)

            if cfg.log_file:
                write_latest_text(cfg.log_file, text, cfg.log_max_chars)
            if cfg.recent_log_file:
                append_recent_log(cfg.recent_log_file, text, cfg.recent_log_lines)

            if cfg.save_screenshot or cfg.save_screenshot_once:
                if not saved_once or cfg.save_screenshot:
                    save_screenshots(img_full, img_crop, cfg.screenshot_dir)
                    saved_once = True
                if cfg.save_screenshot_once:
                    break

            if cfg.only_on_change:
                if text and text != last_text:
                    sender.send(format_message(text, cfg.add_timestamp))
                    last_text = text
            else:
                if text:
                    sender.send(format_message(text, cfg.add_timestamp))

            now = time.time()
            if now - last_heartbeat >= cfg.heartbeat_sec:
                write_status(cfg.status_file, "running")
                last_heartbeat = now

            time.sleep(cfg.poll_sec)
    finally:
        try:
            if cfg.status_file:
                write_status(cfg.status_file, "stopped")
        except Exception:
            pass
        try:
            if cfg.pid_file and os.path.exists(cfg.pid_file):
                os.remove(cfg.pid_file)
        except Exception:
            pass


if __name__ == "__main__":
    main()
