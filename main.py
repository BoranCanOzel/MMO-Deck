"""
Hotkeys for window sizing, tab navigation, volume, refresh, and desktop toggle.

Hotkeys:
  F13               -> cycle LEFT widths
  F14               -> Maximize/Restore active window (ShowWindow)
  F15               -> cycle RIGHT widths
  F16               -> Hard Refresh (Ctrl+F5)
  F17               -> Prev tab  (Ctrl+Shift+Tab)
  F18               -> Next tab  (Ctrl+Tab)
  Shift+F23         -> Browser Back (Alt+Left)
  Shift+F24         -> Browser Forward (Alt+Right)
  F22               -> Toggle Desktop (Win+D)
  F23               -> Volume Down (direct)
  F24               -> Volume Up (direct)

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
import pythoncom
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
VOLUME_DOWN_HOTKEY = "f23"
VOLUME_UP_HOTKEY   = "f24"
TOGGLE_DESKTOP_HOTKEY = "f22"
BROWSER_BACK_HOTKEY = "shift+f23"
BROWSER_FORWARD_HOTKEY = "shift+f24"

# ---------------- TUNING KNOBS ----------------
# Window sizing
WINDOW_WIDTHS = [0.5040, 0.3372, 0.6707]
WINDOW_POS_TOL_PX = 2
WINDOW_CYCLE_DEBOUNCE_SEC = 0.10

# Tab navigation
TAB_REPEAT_INITIAL_SEC = 0.35
TAB_REPEAT_SEC = 0.12

# Volume
VOLUME_REPEAT_INITIAL_SEC = 0.35
VOLUME_REPEAT_SEC = 0.03
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
_last_trigger = 0.0

KEYEVENTF_KEYUP = 0x0002
VK_CONTROL = 0x11
VK_F5 = 0x74
VK_SHIFT = 0x10
VK_TAB = 0x09
VK_MENU = 0x12
VK_LEFT = 0x25
VK_RIGHT = 0x27
VK_LWIN = 0x5B
VK_D = 0x44
VK_VOLUME_UP = 0xAF
VK_VOLUME_DOWN = 0xAE

_volume_endpoint = None
_tab_state = {}
_volume_state = {}
_toggle_state = set()
_shell_app = None


def _debounced() -> bool:
    global _last_trigger
    now = time.time()
    if now - _last_trigger < WINDOW_CYCLE_DEBOUNCE_SEC:
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


def _rect_close(a, b, tol=WINDOW_POS_TOL_PX) -> bool:
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
    targets = [_make_target_rect(work_area, w, side) for w in WINDOW_WIDTHS]

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


def _browser_nav(back: bool):
    # Alt+Left / Alt+Right for browser navigation
    _key_event(VK_MENU)
    _key_event(VK_LEFT if back else VK_RIGHT)
    _key_event(VK_LEFT if back else VK_RIGHT, up=True)
    _key_event(VK_MENU, up=True)


def _tab_press(name: str, shift: bool):
    if name in _tab_state:
        return
    stop_evt = threading.Event()
    _tab_state[name] = stop_evt
    _send_tab_combo(shift)

    def _runner():
        delay = TAB_REPEAT_INITIAL_SEC
        while not stop_evt.wait(delay):
            _send_tab_combo(shift)
            delay = TAB_REPEAT_SEC

    threading.Thread(target=_runner, daemon=True).start()


def _tab_release(name: str):
    stop_evt = _tab_state.pop(name, None)
    if stop_evt:
        stop_evt.set()


def _win_d_chord():
    # Send Win+D with aggressive key-up to avoid Win sticking (and Win+P)
    _key_event(VK_D, up=True)
    _key_event(VK_LWIN, up=True)
    time.sleep(0.005)
    _key_event(VK_LWIN)
    time.sleep(0.005)
    _key_event(VK_D)
    time.sleep(0.015)  # let Windows register the chord
    _key_event(VK_D, up=True)
    _key_event(VK_LWIN, up=True)
    _key_event(VK_LWIN, up=True)  # extra release guard


def _get_shell_app():
    global _shell_app
    if _shell_app is not None:
        return _shell_app
    try:
        _shell_app = win32com.client.Dispatch("Shell.Application")
    except Exception:
        _shell_app = None
    return _shell_app


def _toggle_desktop():
    # Prefer Shell.ToggleDesktop for proper toggle; fall back to a Win+D chord
    shell = _get_shell_app()
    if shell:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            shell.ToggleDesktop()
            return
        except Exception:
            pass
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    try:
        _win_d_chord()
    except Exception:
        pass  # keep hotkey resilient


def _toggle_desktop_press(name: str):
    # Fire only once per physical press (ignore OS key auto-repeat)
    if name in _toggle_state:
        return
    _toggle_state.add(name)
    _toggle_desktop()


def _toggle_desktop_release(name: str):
    _toggle_state.discard(name)


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


def _volume_keypress(up: bool):
    # Hardware-style key events as a fallback if pycaw fails
    vk = VK_VOLUME_UP if up else VK_VOLUME_DOWN
    _key_event(vk)


def _volume_step(up: bool):
    endpoint = _get_volume_endpoint()
    if not endpoint:
        _volume_keypress(up)
        return

    try:
        if up:
            endpoint.VolumeStepUp(None)
        else:
            endpoint.VolumeStepDown(None)
    except Exception:
        _volume_keypress(up)


def _volume_press(name: str, up: bool):
    if name in _volume_state:
        return
    stop_evt = threading.Event()
    _volume_state[name] = stop_evt
    _volume_step(up)

    def _runner():
        delay = VOLUME_REPEAT_INITIAL_SEC
        while not stop_evt.wait(delay):
            _volume_step(up)
            delay = VOLUME_REPEAT_SEC

    threading.Thread(target=_runner, daemon=True).start()


def _volume_release(name: str):
    stop_evt = _volume_state.pop(name, None)
    if stop_evt:
        stop_evt.set()


def _prevent_sleep():
    # Keep the system awake while the hotkey listener runs
    kernel32 = ctypes.windll.kernel32
    prev = kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )
    return prev


def _allow_sleep(prev_state):
    # Restore previous execution state on exit
    ctypes.windll.kernel32.SetThreadExecutionState(prev_state or ES_CONTINUOUS)


def main():
    prev_state = _prevent_sleep()

    keyboard.add_hotkey(LEFT_HOTKEY, _cycle_left)
    keyboard.add_hotkey(MAX_HOTKEY, _maximize_restore_active_window)
    keyboard.add_hotkey(RIGHT_HOTKEY, _cycle_right)

    keyboard.add_hotkey(REFRESH_HOTKEY, _hard_refresh)
    keyboard.on_press_key(PREV_TAB_HOTKEY, lambda e: _tab_press(PREV_TAB_HOTKEY, shift=True))
    keyboard.on_release_key(PREV_TAB_HOTKEY, lambda e: _tab_release(PREV_TAB_HOTKEY))
    keyboard.on_press_key(NEXT_TAB_HOTKEY, lambda e: _tab_press(NEXT_TAB_HOTKEY, shift=False))
    keyboard.on_release_key(NEXT_TAB_HOTKEY, lambda e: _tab_release(NEXT_TAB_HOTKEY))
    keyboard.add_hotkey(BROWSER_BACK_HOTKEY, lambda: _browser_nav(back=True))
    keyboard.add_hotkey(BROWSER_FORWARD_HOTKEY, lambda: _browser_nav(back=False))
    keyboard.on_press_key(TOGGLE_DESKTOP_HOTKEY, lambda e: _toggle_desktop_press(TOGGLE_DESKTOP_HOTKEY))
    keyboard.on_release_key(TOGGLE_DESKTOP_HOTKEY, lambda e: _toggle_desktop_release(TOGGLE_DESKTOP_HOTKEY))
    keyboard.on_press_key(VOLUME_DOWN_HOTKEY, lambda e: _volume_press(VOLUME_DOWN_HOTKEY, up=False))
    keyboard.on_release_key(VOLUME_DOWN_HOTKEY, lambda e: _volume_release(VOLUME_DOWN_HOTKEY))
    keyboard.on_press_key(VOLUME_UP_HOTKEY, lambda e: _volume_press(VOLUME_UP_HOTKEY, up=True))
    keyboard.on_release_key(VOLUME_UP_HOTKEY, lambda e: _volume_release(VOLUME_UP_HOTKEY))

    print("Hotkeys active:")
    print("  F13              LEFT cycle")
    print("  F14              Maximize/Restore (API)")
    print("  F15              RIGHT cycle")
    print("  F16              Hard Refresh (Ctrl+F5)")
    print("  F17              Prev tab (Ctrl+Shift+Tab)")
    print("  F18              Next tab (Ctrl+Tab)")
    print("  Shift+F23        Browser Back (Alt+Left)")
    print("  Shift+F24        Browser Forward (Alt+Right)")
    print("  F22              Toggle Desktop (Win+D)")
    print("  F23              Volume Down")
    print("  F24              Volume Up")
    print("Ctrl+C to exit.")
    try:
        keyboard.wait()
    finally:
        _allow_sleep(prev_state)


if __name__ == "__main__":
    main()
