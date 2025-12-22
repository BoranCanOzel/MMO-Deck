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
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# ---------------- HOTKEYS ----------------
LEFT_HOTKEY  = "f13"
MAX_HOTKEY   = "f14"
RIGHT_HOTKEY = "f15"

REFRESH_HOTKEY   = "f16"
PREV_TAB_HOTKEY  = "f17"
NEXT_TAB_HOTKEY  = "f18"
VOLUME_DOWN_HOTKEY = "f19"
VOLUME_UP_HOTKEY   = "f20"
TOGGLE_DESKTOP_HOTKEY = "f21"

# ---------------- WINDOW CYCLE SETTINGS ----------------
WIDTHS = [0.5040, 0.3372, 0.6707]
TOL_PX = 2
DEBOUNCE_SEC = 0.10
TAB_REPEAT_INITIAL_SEC = 0.35
TAB_REPEAT_SEC = 0.12
_last_trigger = 0.0

KEYEVENTF_KEYUP = 0x0002
VK_CONTROL = 0x11
VK_F5 = 0x74
VK_SHIFT = 0x10
VK_TAB = 0x09

_volume_endpoint = None
_tab_state = {}


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


def _key_event(vk: int, up: bool = False):
    flags = KEYEVENTF_KEYUP if up else 0
    ctypes.windll.user32.keybd_event(vk, 0, flags, 0)


def _hard_refresh():
    _key_event(VK_CONTROL)
    _key_event(VK_F5)
    _key_event(VK_F5, up=True)
    _key_event(VK_CONTROL, up=True)


def _send_tab_combo(shift: bool):
    mods = [VK_CONTROL]
    if shift:
        mods.append(VK_SHIFT)

    for key in mods:
        _key_event(key)

    _key_event(VK_TAB)
    _key_event(VK_TAB, up=True)

    for key in reversed(mods):
        _key_event(key, up=True)


def _prev_tab():
    _send_tab_combo(shift=True)


def _next_tab():
    _send_tab_combo(shift=False)


def _tab_press(name: str, shift: bool):
    # Ignore auto-repeat while held; only fire once per physical press
    if name in _tab_state:
        return
    stop_evt = threading.Event()
    _tab_state[name] = stop_evt
    print(f"{name.upper()} pressed -> {'Prev' if shift else 'Next'} tab (holding sends repeats)")

    def _runner():
        delay = TAB_REPEAT_INITIAL_SEC
        while not stop_evt.wait(delay):
            print(f"{name.upper()} tick -> {'Prev' if shift else 'Next'} tab")
            _send_tab_combo(shift)
            delay = TAB_REPEAT_SEC

    # Fire immediately, then repeat in background while held
    _send_tab_combo(shift)
    threading.Thread(target=_runner, daemon=True).start()


def _tab_release(name: str):
    stop_evt = _tab_state.pop(name, None)
    if stop_evt is None:
        return
    stop_evt.set()
    print(f"{name.upper()} released")


def _toggle_desktop():
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        shell.ToggleDesktop()
    except Exception:
        # Keep the hotkey resilient even if the COM object fails
        pass


def _get_volume_endpoint():
    global _volume_endpoint
    if _volume_endpoint is not None:
        return _volume_endpoint

    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        _volume_endpoint = cast(interface, POINTER(IAudioEndpointVolume))
    except Exception:
        _volume_endpoint = None
    return _volume_endpoint


def _volume_step(up: bool):
    endpoint = _get_volume_endpoint()
    if not endpoint:
        return

    try:
        if up:
            endpoint.VolumeStepUp(None)
        else:
            endpoint.VolumeStepDown(None)
    except Exception:
        # Silently ignore audio failures to keep the hotkey loop alive
        pass


def main():
    keyboard.add_hotkey(LEFT_HOTKEY, _cycle_left)
    keyboard.add_hotkey(MAX_HOTKEY, _maximize_restore_active_window)
    keyboard.add_hotkey(RIGHT_HOTKEY, _cycle_right)

    keyboard.add_hotkey(REFRESH_HOTKEY, _hard_refresh)
    keyboard.on_press_key(PREV_TAB_HOTKEY, lambda e: _tab_press(PREV_TAB_HOTKEY, shift=True))
    keyboard.on_release_key(PREV_TAB_HOTKEY, lambda e: _tab_release(PREV_TAB_HOTKEY))
    keyboard.on_press_key(NEXT_TAB_HOTKEY, lambda e: _tab_press(NEXT_TAB_HOTKEY, shift=False))
    keyboard.on_release_key(NEXT_TAB_HOTKEY, lambda e: _tab_release(NEXT_TAB_HOTKEY))
    keyboard.add_hotkey(TOGGLE_DESKTOP_HOTKEY, _toggle_desktop)
    keyboard.add_hotkey(VOLUME_DOWN_HOTKEY, lambda: _volume_step(up=False))
    keyboard.add_hotkey(VOLUME_UP_HOTKEY, lambda: _volume_step(up=True))

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
