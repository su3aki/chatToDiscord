# OCR -> Discord Webhook (LINE PC OpenChat)
# Minimal, production-ready-ish scaffold with pluggable sender.
# ASCII-only comments, Japanese text can be OCR'ed.

import os
import time
import re
from dataclasses import dataclass
from typing import Optional, Tuple
from datetime import datetime

import requests
from PIL import Image, ImageOps
import pytesseract
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


class Sender:
    def send(self, text: str) -> None:
        raise NotImplementedError


class WebhookSender(Sender):
    def __init__(self, webhook_url: str):
        self.webhook_url = webhook_url

    def send(self, text: str) -> None:
        # Discord webhook message max ~2000 chars
        if not text:
            return
        if len(text) > 1900:
            text = text[:1900] + "..."
        resp = requests.post(self.webhook_url, json={"content": text})
        if resp.status_code >= 400:
            raise RuntimeError(f"Webhook failed: {resp.status_code} {resp.text}")


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

    hwnd = win32gui.FindWindow(None, title)
    if not hwnd:
        # Try partial match
        def enum_handler(h, result):
            if win32gui.IsWindowVisible(h):
                t = win32gui.GetWindowText(h)
                if title in t:
                    result.append(h)
        results = []
        win32gui.EnumWindows(enum_handler, results)
        hwnd = results[0] if results else None

    if not hwnd:
        raise RuntimeError(f"Window not found: {title}")

    # Restore if minimized
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    rect = win32gui.GetWindowRect(hwnd)
    # rect is (left, top, right, bottom)
    return rect


def grab_window_image(rect: Tuple[int, int, int, int]) -> Image.Image:
    left, top, right, bottom = rect
    width = right - left
    height = bottom - top
    with mss.mss() as sct:
        monitor = {"left": left, "top": top, "width": width, "height": height}
        sct_img = sct.grab(monitor)
        img = Image.frombytes("RGB", sct_img.size, sct_img.rgb)
    return img


def apply_crop(img: Image.Image, crop_rect: Optional[Tuple[int, int, int, int]]) -> Image.Image:
    if not crop_rect:
        return img
    return img.crop(crop_rect)


def preprocess_image(img: Image.Image, threshold: int) -> Image.Image:
    gray = ImageOps.grayscale(img)
    # Auto-contrast to improve separation
    gray = ImageOps.autocontrast(gray)
    # Simple binary threshold
    bw = gray.point(lambda x: 255 if x > threshold else 0, mode="1")
    return bw.convert("L")


def ocr_image(img: Image.Image, lang: str) -> str:
    text = pytesseract.image_to_string(img, lang=lang)
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


def save_screenshots(img_full: Image.Image, img_crop: Image.Image, out_dir: str) -> None:
    os.makedirs(out_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    full_path = os.path.join(out_dir, f"line_full_{ts}.png")
    crop_path = os.path.join(out_dir, f"line_crop_{ts}.png")
    img_full.save(full_path)
    img_crop.save(crop_path)
    print(f"Saved screenshots: {full_path} / {crop_path}")


def main() -> None:
    env_path = os.getenv("ENV_FILE", os.path.join(os.path.dirname(__file__), ".env"))
    load_env_file(env_path)

    cfg = Config(
        window_title=os.getenv("LINE_WINDOW_TITLE", "LINE"),
        webhook_url=os.getenv("WEBHOOK_URL", ""),
        poll_sec=float(os.getenv("POLL_SEC", "1.0")),
        ocr_lang=os.getenv("OCR_LANG", "jpn+eng"),
        crop_rect=parse_crop_rect(os.getenv("CROP_RECT", "")),
        keep_newlines=parse_bool(os.getenv("KEEP_NEWLINES", ""), False),
        preprocess=parse_bool(os.getenv("PREPROCESS", ""), False),
        threshold=int(os.getenv("THRESHOLD", "160")),
        add_timestamp=parse_bool(os.getenv("ADD_TIMESTAMP", ""), True),
        save_screenshot=parse_bool(os.getenv("SAVE_SCREENSHOT", ""), False),
        save_screenshot_once=parse_bool(os.getenv("SAVE_SCREENSHOT_ONCE", ""), False),
        screenshot_dir=os.getenv("SCREENSHOT_DIR", "screenshots"),
    )

    if not cfg.webhook_url:
        raise RuntimeError("WEBHOOK_URL env var is required")

    sender = WebhookSender(cfg.webhook_url)

    last_text = ""
    saved_once = False
    print("Starting OCR loop...")
    while True:
        rect = get_window_rect(cfg.window_title)
        img_full = grab_window_image(rect)
        img_crop = apply_crop(img_full, cfg.crop_rect)
        if cfg.preprocess:
            img_ocr = preprocess_image(img_crop, cfg.threshold)
        else:
            img_ocr = img_crop
        text = ocr_image(img_ocr, cfg.ocr_lang)
        if cfg.normalize_whitespace:
            text = normalize_text(text, cfg.keep_newlines)

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

        time.sleep(cfg.poll_sec)


if __name__ == "__main__":
    main()
