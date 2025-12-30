"""
Hotkeys for window sizing, tab navigation, volume, refresh, and desktop toggle.

Hotkeys:
  F13               -> cycle LEFT widths
  F14               -> Maximize/Restore active window (ShowWindow)
  F15               -> cycle RIGHT widths
  Shift+F13         -> Cycle BOTTOM heights (Y axis)
  Shift+F15         -> Cycle TOP heights (Y axis)
  F16               -> Tap: Refresh (Ctrl+R / Ctrl+/), Hold: Hard Refresh (Ctrl+F5 or Ctrl+/)
  F17               -> Prev tab  (Ctrl+Shift+Tab)
  F18               -> Next tab  (Ctrl+Tab)
  F19               -> Print Screen
  Ctrl+F23          -> Switch desktop left (Win+Ctrl+Left)
  Ctrl+F24          -> Switch desktop right (Win+Ctrl+Right)
  F21               -> Open This PC
  F22               -> Toggle Desktop (Win+D)
  F23               -> Volume Down (direct)
  F24               -> Volume Up (direct)

Install:
  pip install keyboard pywin32 pycaw comtypes
"""

import os
import sys
import time
import ctypes
import threading
import tkinter as tk
from tkinter import ttk
import subprocess

import keyboard
import win32gui
import win32con
import win32api
import win32com.client
import pythoncom
import win32process
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

try:
    import pystray
    from PIL import Image, ImageDraw
except Exception:
    pystray = None
    Image = None
    ImageDraw = None

# ---------------- HOTKEYS ----------------
LEFT_HOTKEY  = "f13"
MAX_HOTKEY   = "f14"
RIGHT_HOTKEY = "f15"

REFRESH_HOTKEY   = "f16"
PREV_TAB_HOTKEY  = "f17"
NEXT_TAB_HOTKEY  = "f18"
PRINT_SCREEN_HOTKEY = "f19"

BROWSER_BACK_HOTKEY = "shift+f23"
BROWSER_FORWARD_HOTKEY = "shift+f24"
VOLUME_DOWN_HOTKEY = "f23"
VOLUME_UP_HOTKEY   = "f24"
TOGGLE_DESKTOP_HOTKEY = "f22"
OPEN_THIS_PC_HOTKEY = "f21"

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

# Browser detection
BROWSER_PROCESSES = {"chrome.exe"}

REFRESH_HOLD_THRESHOLD_SEC = 0.40
APP_NAME = "MMO Deck"
STARTUP_LINK_NAME = "MMO Deck.lnk"
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
_last_trigger = 0.0

KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_EXTENDEDKEY = 0x0001
VK_CONTROL = 0x11
VK_F5 = 0x74
VK_SHIFT = 0x10
VK_TAB = 0x09
VK_PRIOR = 0x21  # Page Up
VK_NEXT = 0x22   # Page Down
VK_MENU = 0x12
VK_LEFT = 0x25
VK_RIGHT = 0x27
VK_OEM_2 = 0xBF  # '/' key
VK_LWIN = 0x5B
VK_D = 0x44
VK_VOLUME_UP = 0xAF
VK_VOLUME_DOWN = 0xAE
VK_SNAPSHOT = 0x2C  # Print Screen

_volume_endpoint = None
_tab_state = {}
_volume_state = {}
_toggle_state = set()
_this_pc_state = set()
_maximize_state = set()
_shell_app = None
_refresh_state = {}
_tray_icon = None
_root = None


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


def _get_foreground_process_name():
    hwnd = _get_foreground_window()
    if not hwnd:
        return None
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
            False,
            pid,
        )
        try:
            exe = win32process.GetModuleFileNameEx(handle, 0)
            return os.path.basename(exe).lower()
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        return None


def _is_browser_window():
    proc = _get_foreground_process_name()
    return proc in BROWSER_PROCESSES if proc else False


def _startup_shortcut_path():
    startup_dir = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
    return os.path.join(startup_dir, STARTUP_LINK_NAME)


def _add_to_startup():
    try:
        path = _startup_shortcut_path()
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(path)
        shortcut.TargetPath = sys.executable
        shortcut.Arguments = f'"{os.path.abspath(sys.argv[0])}"'
        shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(sys.argv[0]))
        shortcut.IconLocation = sys.executable
        shortcut.Save()
        print(f"Startup: added shortcut at {path}")
    except Exception as exc:
        print(f"Startup: failed to add ({exc})")


def _remove_from_startup():
    try:
        path = _startup_shortcut_path()
        if os.path.exists(path):
            os.remove(path)
            print(f"Startup: removed shortcut at {path}")
        else:
            print("Startup: no shortcut to remove")
    except Exception as exc:
        print(f"Startup: failed to remove ({exc})")


def _send_ctrl_combo(key: str):
    # Temporarily release Shift so we don't send Ctrl+Shift+key
    had_shift = keyboard.is_pressed("shift")
    if had_shift:
        keyboard.release("shift")
    try:
        keyboard.send(f"ctrl+{key}")
    finally:
        if had_shift:
            keyboard.press("shift")


def _send_ctrl_slash():
    # Send Ctrl + '/' reliably, releasing Shift if held
    had_shift = keyboard.is_pressed("shift")
    if had_shift:
        keyboard.release("shift")
    try:
        _key_event(VK_CONTROL)
        _key_event(VK_OEM_2)
        _key_event(VK_OEM_2, up=True)
        _key_event(VK_CONTROL, up=True)
    finally:
        if had_shift:
            keyboard.press("shift")


def _create_tray_image():
    if Image is None:
        return None
    icon_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "icon.ico")
    try:
        return Image.open(icon_path)
    except Exception as exc:
        print(f"Tray icon: failed to load {icon_path} ({exc}); using fallback.")
        if ImageDraw is None:
            return None
        img = Image.new("RGB", (64, 64), (43, 119, 232))
        d = ImageDraw.Draw(img)
        d.rectangle([18, 18, 46, 46], fill=(255, 255, 255))
        return img


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


def _clamp_width_to_work_area(rect, work_area):
    wl, _, wr, _ = work_area
    l, t, r, b = rect
    width = r - l
    max_width = wr - wl

    if width > max_width:
        width = max_width
    l = max(wl, l)
    if l + width > wr:
        l = wr - width
    r = l + width
    return (l, t, r, b)


def _make_vertical_target_rects(work_area, height_ratios, anchor: str, current_rect):
    wl, wt, wr, wb = work_area
    work_h = wb - wt
    l, _, r, _ = _clamp_width_to_work_area(current_rect, work_area)
    targets = []
    for hr in height_ratios:
        h = int(round(work_h * hr))
        h = min(h, work_h)
        if anchor == "top":
            t = wt
            b = wt + h
        elif anchor == "bottom":
            b = wb
            t = wb - h
        else:
            raise ValueError("anchor must be 'top' or 'bottom'")
        targets.append((l, t, r, b))
    return targets


def _cycle_heights(anchor: str):
    if not _debounced():
        return

    hwnd = _get_foreground_window()
    if not hwnd or _is_ignorable_window(hwnd):
        return

    placement = win32gui.GetWindowPlacement(hwnd)
    is_maximized = (placement[1] == win32con.SW_SHOWMAXIMIZED)
    if is_maximized:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        current = _get_window_rect(hwnd)
    else:
        current = _get_window_rect(hwnd)

    work_area = _get_monitor_work_area_for_window(hwnd)
    targets = _make_vertical_target_rects(work_area, WINDOW_WIDTHS, anchor, current)

    next_rect = targets[0]
    for i, tr in enumerate(targets):
        if _rect_close(current, tr):
            next_rect = targets[(i + 1) % len(targets)]
            break

    _set_window_rect(hwnd, next_rect)


def _cycle_bottom_heights():
    _cycle_heights("bottom")


def _cycle_top_heights():
    _cycle_heights("top")


def _set_vertical_position(position: str):
    if not _debounced():
        return

    hwnd = _get_foreground_window()
    if not hwnd or _is_ignorable_window(hwnd):
        return

    if position != "full":
        placement = win32gui.GetWindowPlacement(hwnd)
        if placement[1] == win32con.SW_SHOWMAXIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    work_area = _get_monitor_work_area_for_window(hwnd)
    wl, wt, wr, wb = work_area
    work_h = wb - wt

    cur_rect = _get_window_rect(hwnd)
    cur_rect = _clamp_width_to_work_area(cur_rect, work_area)
    l, _, r, _ = cur_rect

    if position == "top":
        t = wt
        b = wt + int(round(work_h * 0.5))
    elif position == "bottom":
        b = wb
        t = wb - int(round(work_h * 0.5))
    elif position == "full":
        t = wt
        b = wb
    else:
        raise ValueError("position must be 'top', 'bottom', or 'full'")

    _set_window_rect(hwnd, (l, t, r, b))


def _set_top_half():
    _set_vertical_position("top")


def _set_bottom_half():
    _set_vertical_position("bottom")


def _set_full_height():
    _set_vertical_position("full")


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
    if _is_browser_window():
        _key_event(VK_CONTROL)
        _key_event(VK_F5)
        _key_event(VK_F5, up=True)
        _key_event(VK_CONTROL, up=True)
    else:
        _send_ctrl_slash()


def _refresh_tap():
    if _is_browser_window():
        print("Refresh: browser (Ctrl+R)")
        keyboard.send("ctrl+r")
    else:
        print("Refresh: non-browser (Ctrl+/)")
        _send_ctrl_slash()


def _refresh_hold():
    if _is_browser_window():
        print("Hard refresh: browser (Ctrl+F5)")
        _hard_refresh()
    else:
        print("Refresh (hold): non-browser (Ctrl+/)")
        _send_ctrl_slash()


def _print_screen():
    # Release modifiers to avoid Win/Alt/Shift altering the snapshot behavior
    released = []
    for mod in ("windows", "alt", "ctrl", "shift"):
        if keyboard.is_pressed(mod):
            keyboard.release(mod)
            released.append(mod)
    try:
        # Simple, reliable: ask keyboard to send Print Screen; fallback to direct key events
        keyboard.send("print screen")
    except Exception:
        _key_event(VK_SNAPSHOT)
        _key_event(VK_SNAPSHOT, up=True)
    finally:
        for mod in released:
            keyboard.press(mod)


def _send_tab_combo(shift: bool):
    if _is_browser_window():
        # Browser: Ctrl+Tab / Ctrl+Shift+Tab
        mods = [VK_CONTROL]
        if shift:
            mods.append(VK_SHIFT)
        for key in mods:
            _key_event(key)
        _key_event(VK_TAB)
        _key_event(VK_TAB, up=True)
        for key in reversed(mods):
            _key_event(key, up=True)
    else:
        # Non-browser: Ctrl+PgUp / Ctrl+PgDn via keyboard for reliability
        keyboard.send("ctrl+page up" if shift else "ctrl+page down")


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


def _switch_virtual_desktop(back: bool):
    # Send Win+Ctrl+Left/Right; temporarily release Shift so it doesn't move windows
    had_shift = keyboard.is_pressed("shift")
    if had_shift:
        keyboard.release("shift")
    try:
        _key_event(VK_LWIN)
        _key_event(VK_CONTROL)
        _key_event(VK_LEFT if back else VK_RIGHT)
        _key_event(VK_LEFT if back else VK_RIGHT, up=True)
        _key_event(VK_CONTROL, up=True)
        _key_event(VK_LWIN, up=True)
    finally:
        if had_shift:
            keyboard.press("shift")


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


def _maximize_press(name: str):
    # Fire once per physical press to avoid rapid toggle on key repeat
    if name in _maximize_state:
        return
    _maximize_state.add(name)
    _maximize_restore_active_window()


def _maximize_release(name: str):
    _maximize_state.discard(name)


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


def _legacy_shift_f23_action():
    # Preserved old behavior: browser back or Ctrl+Z
    if _is_browser_window():
        _browser_nav(back=True)
    else:
        _send_ctrl_combo("z")


def _legacy_shift_f24_action():
    # Preserved old behavior: browser forward or Ctrl+Y
    if _is_browser_window():
        _browser_nav(back=False)
    else:
        _send_ctrl_combo("y")


def _handle_f23_press(e):
    # Debug logging noisy during normal use; keep commented for troubleshooting.
    # proc = _get_foreground_process_name()
    # print(f"Active process: {proc or 'unknown'}")
    if keyboard.is_pressed("ctrl"):
        _switch_virtual_desktop(back=True)
        return
    _volume_press(VOLUME_DOWN_HOTKEY, up=False)


def _handle_f23_release(e):
    if keyboard.is_pressed("ctrl"):
        return
    _volume_release(VOLUME_DOWN_HOTKEY)


def _handle_f24_press(e):
    # Debug logging noisy during normal use; keep commented for troubleshooting.
    # proc = _get_foreground_process_name()
    # print(f"Active process: {proc or 'unknown'}")
    if keyboard.is_pressed("ctrl"):
        _switch_virtual_desktop(back=False)
        return
    _volume_press(VOLUME_UP_HOTKEY, up=True)


def _handle_f24_release(e):
    if keyboard.is_pressed("ctrl"):
        return
    _volume_release(VOLUME_UP_HOTKEY)


def _open_this_pc():
    try:
        os.startfile("shell:MyComputerFolder")
        return
    except Exception:
        pass
    try:
        subprocess.Popen(["explorer.exe", "shell:MyComputerFolder"])
    except Exception as exc:
        print(f"This PC: failed to open ({exc})")


def _open_this_pc_press(name: str):
    # Fire once per physical press to avoid repeat-open on key auto-repeat
    if name in _this_pc_state:
        return
    _this_pc_state.add(name)
    _open_this_pc()


def _open_this_pc_release(name: str):
    _this_pc_state.discard(name)


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


def _refresh_press(name: str):
    global _refresh_state
    if _refresh_state.get("active"):
        return
    _refresh_state = {"active": True, "hold": False}

    def _hold_action():
        # If still active when timer fires, treat as hold
        if _refresh_state.get("active"):
            _refresh_state["hold"] = True
            _refresh_hold()

    timer = threading.Timer(REFRESH_HOLD_THRESHOLD_SEC, _hold_action)
    _refresh_state["timer"] = timer
    timer.start()


def _refresh_release(name: str):
    global _refresh_state
    state = _refresh_state if isinstance(_refresh_state, dict) else {}
    timer = state.get("timer")
    if timer:
        timer.cancel()
    if state.get("active") and not state.get("hold"):
        _refresh_tap()
    _refresh_state = {}


def _hide_window(auto: bool = False):
    if _root:
        # On auto-hide, don't disappear if tray isn't available
        if auto and not (pystray and Image):
            print("Tray icon not available (pystray/Pillow missing); keeping window visible.")
            return
        _root.withdraw()  # hide from taskbar
        if pystray and Image:
            _start_tray()
        else:
            print("Tray icon not available (pystray/Pillow missing); window hidden.")


def _show_window():
    global _tray_icon
    if _root:
        _root.deiconify()
        _root.lift()
        _root.focus_force()
    if _tray_icon:
        _tray_icon.stop()
        _tray_icon = None


def _tray_quit(icon, item):
    if _root:
        _root.after(0, _root.quit)


def _start_tray():
    global _tray_icon
    if _tray_icon or not pystray or not Image:
        return

    def on_show(icon, item):
        _show_window()

    def on_quit(icon, item):
        _tray_quit(icon, item)

    image = _create_tray_image()
    _tray_icon = pystray.Icon(APP_NAME, image, APP_NAME, menu=pystray.Menu(
        pystray.MenuItem("Show", on_show, default=True),  # double-click default
        pystray.MenuItem("Quit", on_quit),
    ))
    threading.Thread(target=_tray_icon.run, daemon=True).start()


def _auto_hide_on_start():
    # Hide to tray right after launch if tray is available
    if pystray and Image:
        _hide_window(auto=True)
    else:
        print("Startup: tray dependencies missing; window will stay visible.")


def _build_gui():
    global _root
    _root = tk.Tk()
    _root.title(APP_NAME)
    _root.geometry("360x230")
    _root.protocol("WM_DELETE_WINDOW", _hide_window)

    frame = ttk.Frame(_root, padding=12)
    frame.pack(fill="both", expand=True)

    ttk.Label(frame, text="MMO Deck Controls").pack(anchor="w")

    ttk.Button(frame, text="Hide to tray", command=_hide_window).pack(fill="x", pady=4)
    ttk.Button(frame, text="Add to Startup", command=_add_to_startup).pack(fill="x", pady=4)
    ttk.Button(frame, text="Remove from Startup", command=_remove_from_startup).pack(fill="x", pady=4)
    ttk.Button(frame, text="Quit", command=_root.quit).pack(fill="x", pady=12)

    return _root


def main():
    prev_state = _prevent_sleep()

    keyboard.on_press_key(LEFT_HOTKEY, lambda e: _cycle_bottom_heights() if keyboard.is_pressed("shift") else _cycle_left())
    keyboard.on_press_key(MAX_HOTKEY, lambda e: _maximize_press(MAX_HOTKEY))
    keyboard.on_release_key(MAX_HOTKEY, lambda e: _maximize_release(MAX_HOTKEY))
    keyboard.on_press_key(RIGHT_HOTKEY, lambda e: _cycle_top_heights() if keyboard.is_pressed("shift") else _cycle_right())

    keyboard.on_press_key(REFRESH_HOTKEY, lambda e: _refresh_press(REFRESH_HOTKEY))
    keyboard.on_release_key(REFRESH_HOTKEY, lambda e: _refresh_release(REFRESH_HOTKEY))
    keyboard.on_press_key(PREV_TAB_HOTKEY, lambda e: _tab_press(PREV_TAB_HOTKEY, shift=True))
    keyboard.on_release_key(PREV_TAB_HOTKEY, lambda e: _tab_release(PREV_TAB_HOTKEY))
    keyboard.on_press_key(NEXT_TAB_HOTKEY, lambda e: _tab_press(NEXT_TAB_HOTKEY, shift=False))
    keyboard.on_release_key(NEXT_TAB_HOTKEY, lambda e: _tab_release(NEXT_TAB_HOTKEY))
    keyboard.on_press_key(PRINT_SCREEN_HOTKEY, lambda e: _print_screen())
    keyboard.on_press_key(OPEN_THIS_PC_HOTKEY, lambda e: _open_this_pc_press(OPEN_THIS_PC_HOTKEY))
    keyboard.on_release_key(OPEN_THIS_PC_HOTKEY, lambda e: _open_this_pc_release(OPEN_THIS_PC_HOTKEY))
    keyboard.on_press_key(TOGGLE_DESKTOP_HOTKEY, lambda e: _toggle_desktop_press(TOGGLE_DESKTOP_HOTKEY))
    keyboard.on_release_key(TOGGLE_DESKTOP_HOTKEY, lambda e: _toggle_desktop_release(TOGGLE_DESKTOP_HOTKEY))
    keyboard.on_press_key(VOLUME_DOWN_HOTKEY, _handle_f23_press)
    keyboard.on_release_key(VOLUME_DOWN_HOTKEY, _handle_f23_release)
    keyboard.on_press_key(VOLUME_UP_HOTKEY, _handle_f24_press)
    keyboard.on_release_key(VOLUME_UP_HOTKEY, _handle_f24_release)

    print("Hotkeys active:")
    print("  F13              LEFT cycle")
    print("  F14              Maximize/Restore (API)")
    print("  F15              RIGHT cycle")
    print("  Shift+F13        Cycle BOTTOM heights (Y axis)")
    print("  Shift+F15        Cycle TOP heights (Y axis)")
    print("  F16              Tap: Refresh / Hold: Hard Refresh")
    print("  F17              Prev tab (Ctrl+Shift+Tab)")
    print("  F18              Next tab (Ctrl+Tab)")
    print("  F19              Print Screen")
    print("  F21              Open This PC")
    print("  Ctrl+F23         Switch desktop left (Win+Ctrl+Left)")
    print("  Ctrl+F24         Switch desktop right (Win+Ctrl+Right)")
    print("  F22              Toggle Desktop (Win+D)")
    print("  F23              Volume Down")
    print("  F24              Volume Up")
    print("Close/hide via the GUI (tray) or Quit button.")

    gui = _build_gui()
    _auto_hide_on_start()
    try:
        gui.mainloop()
    finally:
        if _tray_icon:
            try:
                _tray_icon.stop()
            except Exception:
                pass
        _allow_sleep(prev_state)


if __name__ == "__main__":
    main()
