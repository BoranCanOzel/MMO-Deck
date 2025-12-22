"""
Reliable API-based actions (no Win-key simulation).

Hotkeys:
  F13               -> cycle LEFT widths
  F14               -> Maximize/Restore active window (ShowWindow)
  F15               -> cycle RIGHT widths
  F16               -> Hard Refresh (Ctrl+F5)
  F17               -> Prev tab  (Ctrl+Shift+Tab)
  F18               -> Next tab  (Ctrl+Tab)
  F21               -> Toggle Desktop (Shell.Application.ToggleDesktop)
  F19               -> Volume Down (direct)
  F20               -> Volume Up (direct)

Install:
  pip install keyboard pywin32 pycaw comtypes
"""

import time
import ctypes
import threading
import keyboard
import win32gui
import win32con
import win32api
import win32com.client

try:
    from ctypes import POINTER, cast
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    _HAS_CORE_AUDIO = True
except ImportError:
    _HAS_CORE_AUDIO = False
    POINTER = None  # type: ignore
    cast = None  # type: ignore
    CLSCTX_ALL = None  # type: ignore
    AudioUtilities = None  # type: ignore
    IAudioEndpointVolume = None  # type: ignore

# ---------------- HOTKEYS ----------------
LEFT_HOTKEY  = "f13"
MAX_HOTKEY   = "f14"
RIGHT_HOTKEY = "f15"

REFRESH_HOTKEY   = "f16"
TAB_PREV_HOTKEY  = "f17"
TAB_NEXT_HOTKEY  = "f18"
DESKTOP_HOTKEY   = "f21"
VOL_DOWN_HOTKEY  = "f19"
VOL_UP_HOTKEY    = "f20"

# ---------------- WINDOW CYCLE SETTINGS ----------------
WIDTHS = [0.5040, 0.3372, 0.6707]
TOL_PX = 2
DEBOUNCE_SEC = 0.10
_last_trigger = 0.0

# ---------------- VOLUME REPEAT ----------------
VOL_REPEAT_START_SEC = 0.0
VOL_REPEAT_INTERVAL_SEC = 0.03
VOL_STEP = 0.01  # 1% per step
_vol_repeat = {
    "down": {"thread": None, "stop_event": None},
    "up": {"thread": None, "stop_event": None},
}

_endpoint_volume = None
KEYEVENTF_KEYUP = 0x0002
VK_CONTROL = 0x11
VK_SHIFT = 0x10
VK_TAB = 0x09
VK_F5 = 0x74


def _debounced() -> bool:
    global _last_trigger
    now = time.time()
    if now - _last_trigger < DEBOUNCE_SEC:
        return False
    _last_trigger = now
    return True


def _get_foreground_window():
    hwnd = win32gui.GetForegroundWindow()
    return hwnd if hwnd else None


def _is_ignorable_window(hwnd: int) -> bool:
    cls = win32gui.GetClassName(hwnd)
    return cls in ("Progman", "WorkerW", "Shell_TrayWnd")


def _get_monitor_work_area_for_window(hwnd: int):
    monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
    info = win32api.GetMonitorInfo(monitor)
    return info["Work"]  # (l,t,r,b)


def _get_window_rect(hwnd: int):
    return win32gui.GetWindowRect(hwnd)


def _rect_close(a, b, tol=TOL_PX) -> bool:
    return all(abs(a[i] - b[i]) <= tol for i in range(4))


def _make_target_rect(work_area, width_ratio: float, side: str):
    wl, wt, wr, wb = work_area
    work_w = wr - wl
    work_h = wb - wt

    target_w = int(round(work_w * width_ratio))
    t = wt
    b = wt + work_h

    if side == "left":
        l = wl
        r = l + target_w
    elif side == "right":
        r = wr
        l = r - target_w
    else:
        raise ValueError("side must be 'left' or 'right'")

    return (l, t, r, b)


def _set_window_rect(hwnd: int, rect):
    l, t, r, b = rect
    w = r - l
    h = b - t
    win32gui.SetWindowPos(
        hwnd,
        None,
        l, t, w, h,
        win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE
    )


def _cycle_widths(side: str):
    if not _debounced():
        return

    hwnd = _get_foreground_window()
    if not hwnd or _is_ignorable_window(hwnd):
        return

    work_area = _get_monitor_work_area_for_window(hwnd)
    targets = [_make_target_rect(work_area, w, side) for w in WIDTHS]

    # IMPORTANT: if maximized, restart at 50.40% (targets[0])
    placement = win32gui.GetWindowPlacement(hwnd)
    is_maximized = (placement[1] == win32con.SW_SHOWMAXIMIZED)
    if is_maximized:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        _set_window_rect(hwnd, targets[0])
        return

    current = _get_window_rect(hwnd)

    next_rect = targets[0]
    for i, tr in enumerate(targets):
        if _rect_close(current, tr):
            next_rect = targets[(i + 1) % len(targets)]
            break

    _set_window_rect(hwnd, next_rect)


def _cycle_left():
    _cycle_widths("left")


def _cycle_right():
    _cycle_widths("right")


def _maximize_restore_active_window():
    hwnd = _get_foreground_window()
    if not hwnd or _is_ignorable_window(hwnd):
        return

    placement = win32gui.GetWindowPlacement(hwnd)
    show_cmd = placement[1]

    # Toggle maximize/restore reliably via ShowWindow
    if show_cmd == win32con.SW_SHOWMAXIMIZED:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    else:
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)


def _toggle_desktop():
    # Reliable "Show Desktop" toggle via Shell COM
    shell = win32com.client.Dispatch("Shell.Application")
    shell.ToggleDesktop()


def _get_endpoint_volume():
    global _endpoint_volume
    if not _HAS_CORE_AUDIO:
        return None
    if _endpoint_volume is not None:
        return _endpoint_volume
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        _endpoint_volume = cast(interface, POINTER(IAudioEndpointVolume))
    except Exception:
        _endpoint_volume = None
    return _endpoint_volume


def _volume_step(direction: str):
    endpoint = _get_endpoint_volume()
    if not endpoint:
        return
    try:
        current = endpoint.GetMasterVolumeLevelScalar()
        delta = VOL_STEP if direction == "up" else -VOL_STEP
        target = min(1.0, max(0.0, current + delta))
        endpoint.SetMasterVolumeLevelScalar(target, None)
    except Exception:
        pass


def _volume_repeat_runner(direction: str, step_fn):
    stop_event = _vol_repeat[direction]["stop_event"]
    # Repeat as long as the key stays down; no initial lag
    if stop_event and stop_event.wait(VOL_REPEAT_START_SEC):
        return
    while stop_event and not stop_event.wait(VOL_REPEAT_INTERVAL_SEC):
        step_fn()
    _vol_repeat[direction]["thread"] = None
    _vol_repeat[direction]["stop_event"] = None


def _start_volume_repeat(direction: str, step_fn):
    state = _vol_repeat[direction]
    if state["thread"] and state["thread"].is_alive():
        return
    stop_event = threading.Event()
    state["stop_event"] = stop_event
    step_fn()  # immediate first step
    t = threading.Thread(
        target=_volume_repeat_runner,
        args=(direction, step_fn),
        daemon=True,
    )
    state["thread"] = t
    t.start()


def _stop_volume_repeat(direction: str):
    state = _vol_repeat[direction]
    stop_event = state.get("stop_event")
    if stop_event:
        stop_event.set()
    t = state.get("thread")
    if t and t.is_alive():
        t.join(timeout=0.35)
    state["thread"] = None
    state["stop_event"] = None


def _volume_down_press():
    _start_volume_repeat("down", lambda: _volume_step("down"))


def _volume_down_release():
    _stop_volume_repeat("down")


def _volume_up_press():
    _start_volume_repeat("up", lambda: _volume_step("up"))


def _volume_up_release():
    _stop_volume_repeat("up")


def _key_event(vk: int, up: bool = False):
    flags = KEYEVENTF_KEYUP if up else 0
    ctypes.windll.user32.keybd_event(vk, 0, flags, 0)


def _send_tab_combo(reverse: bool):
    if reverse:
        _key_event(VK_CONTROL)
        _key_event(VK_SHIFT)
        _key_event(VK_TAB)
        _key_event(VK_TAB, up=True)
        _key_event(VK_SHIFT, up=True)
        _key_event(VK_CONTROL, up=True)
    else:
        _key_event(VK_SHIFT, up=True)  # ensure shift not held
        _key_event(VK_CONTROL)
        _key_event(VK_TAB)
        _key_event(VK_TAB, up=True)
        _key_event(VK_CONTROL, up=True)


def _next_tab():
    _send_tab_combo(reverse=False)


def _prev_tab():
    _send_tab_combo(reverse=True)


def _hard_refresh():
    _key_event(VK_CONTROL)
    _key_event(VK_F5)
    _key_event(VK_F5, up=True)
    _key_event(VK_CONTROL, up=True)


def main():
    if not _HAS_CORE_AUDIO:
        print("Volume hotkeys need pycaw + comtypes installed for direct control.")

    keyboard.add_hotkey(LEFT_HOTKEY, _cycle_left)
    keyboard.add_hotkey(MAX_HOTKEY, _maximize_restore_active_window)
    keyboard.add_hotkey(RIGHT_HOTKEY, _cycle_right)

    keyboard.add_hotkey(REFRESH_HOTKEY, _hard_refresh)
    keyboard.add_hotkey(TAB_PREV_HOTKEY, _prev_tab)
    keyboard.add_hotkey(TAB_NEXT_HOTKEY, _next_tab)
    keyboard.add_hotkey(DESKTOP_HOTKEY, _toggle_desktop)
    keyboard.add_hotkey(VOL_DOWN_HOTKEY, _volume_down_press)
    keyboard.add_hotkey(VOL_DOWN_HOTKEY, _volume_down_release, trigger_on_release=True)
    keyboard.add_hotkey(VOL_UP_HOTKEY, _volume_up_press)
    keyboard.add_hotkey(VOL_UP_HOTKEY, _volume_up_release, trigger_on_release=True)

    print("Hotkeys active:")
    print("  F13              LEFT cycle")
    print("  F14              Maximize/Restore (API)")
    print("  F15              RIGHT cycle")
    print("  F16              Hard Refresh (Ctrl+F5)")
    print("  F17              Prev tab (Ctrl+Shift+Tab)")
    print("  F18              Next tab (Ctrl+Tab)")
    print("  F21              Toggle Desktop (API)")
    print("  F19              Volume Down")
    print("  F20              Volume Up")
    print("Ctrl+C to exit.")
    keyboard.wait()


if __name__ == "__main__":
    main()
