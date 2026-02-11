import time
import os
from typing import Tuple

try:
    import win32gui
    import win32con
    import win32api
    import win32process
except Exception:
    win32gui = None
    win32con = None
    win32api = None
    win32process = None


def get_window_rect(title: str) -> Tuple[int, int, int, int]:
    if win32gui is None:
        raise RuntimeError("pywin32 is required")

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

    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    return win32gui.GetWindowRect(hwnd)


def get_client_rect(title: str) -> Tuple[int, int, int, int]:
    if win32gui is None:
        raise RuntimeError("pywin32 is required")

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

    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    pt_left_top = win32gui.ClientToScreen(hwnd, (left, top))
    pt_right_bottom = win32gui.ClientToScreen(hwnd, (right, bottom))
    return (pt_left_top[0], pt_left_top[1], pt_right_bottom[0], pt_right_bottom[1])


def enable_dpi_awareness() -> None:
    try:
        import ctypes
        shcore = ctypes.WinDLL("shcore")
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            import ctypes
            user32 = ctypes.WinDLL("user32")
            user32.SetProcessDPIAware()
        except Exception:
            pass


def wait_key(vk: int) -> None:
    # wait for key press transition
    while True:
        if win32api.GetAsyncKeyState(vk) & 0x1:
            break
        time.sleep(0.05)

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


def main() -> None:
    if win32api is None:
        raise RuntimeError("pywin32 is required")

    enable_dpi_awareness()

    title = os.getenv("LINE_WINDOW_TITLE", "LINE")
    print(f"Target window: {title}")
    print("Move mouse to TOP-LEFT of chat area, then press F8")
    wait_key(win32con.VK_F8)
    x1, y1 = win32api.GetCursorPos()

    print("Move mouse to BOTTOM-RIGHT of chat area, then press F9")
    wait_key(win32con.VK_F9)
    x2, y2 = win32api.GetCursorPos()

    # Use client rect to avoid frame/border offsets
    left, top, right, bottom = get_client_rect(title)
    rx1, ry1 = x1 - left, y1 - top
    rx2, ry2 = x2 - left, y2 - top

    l = min(rx1, rx2)
    t = min(ry1, ry2)
    r = max(rx1, rx2)
    b = max(ry1, ry2)

    print(f"CROP_RECT={l},{t},{r},{b}")
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    update_env_value(env_path, "CROP_RECT", f"{l},{t},{r},{b}")
    print(f"Updated .env: {env_path}")


if __name__ == "__main__":
    main()
