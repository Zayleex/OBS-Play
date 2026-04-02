import win32gui, win32process, win32api
import psutil
import obspython as obs
import os
import shutil
import ctypes
import re

# Constants for SHQueryUserNotificationState (from shellapi.h)
QUNS_BUSY = 2
QUNS_PRESENTATION_MODE = 3
QUNS_RUNNING_D3D_FULL_SCREEN = 4


def is_fullscreen_state():
    """
    Checks via the Windows API if an app is running in a fullscreen state.
    """
    state = ctypes.c_int()
    try:
        ctypes.windll.shell32.SHQueryUserNotificationState(ctypes.byref(state))
        if state.value in (QUNS_BUSY, QUNS_PRESENTATION_MODE, QUNS_RUNNING_D3D_FULL_SCREEN):
            return True
    except Exception as e:
        print(f"Error checking fullscreen state: {e}")
    return False


def is_on_primary_display(hwnd):
    """
    Checks if the specified window covers the entire primary monitor.
    """
    try:
        monitor_width = win32api.GetSystemMetrics(0)
        monitor_height = win32api.GetSystemMetrics(1)
        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect
        if left <= 0 and top <= 0 and right >= monitor_width and bottom >= monitor_height:
            return True
    except Exception:
        pass
    return False


def get_active_context():
    """
    Determines the folder name. If it's a browser or not a primary fullscreen app,
    it returns 'Desktop'. Otherwise, it returns the clean window title.
    """
    hwnd = win32gui.GetForegroundWindow()

    # 1. Get the process name (.exe) to identify browsers
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        exe_name = psutil.Process(pid).name().lower()
    except Exception:
        exe_name = ""

    # List of common browsers
    browser_exes = ["chrome.exe", "firefox.exe", "msedge.exe", "opera.exe", "brave.exe", "vivaldi.exe"]

    # 2. Rule: If it's a browser, it's always 'Desktop'
    if exe_name in browser_exes:
        return "Desktop"

    # 3. Rule: If it's NOT fullscreen or NOT on the main monitor, it's 'Desktop'
    if not is_fullscreen_state() or not is_on_primary_display(hwnd):
        return "Desktop"

    # 4. It's a game/fullscreen app: Get the window title
    title = win32gui.GetWindowText(hwnd)
    if not title:
        title = exe_name.replace(".exe", "").capitalize() if exe_name else "Unknown"

    # Clean the title for Windows folder naming conventions
    clean_title = re.sub(r'[\\/*?:"<>|]', "", title).strip()
    return clean_title if clean_title else "Unknown"


def on_event(event):
    """
    Main trigger for the OBS replay buffer event.
    """
    if event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_SAVED:
        app_name = get_active_context()
        replay_path = obs.obs_frontend_get_last_replay()

        if replay_path and os.path.exists(replay_path):
            base_dir = os.path.dirname(replay_path)
            filename = os.path.basename(replay_path)

            # Create target: .../Replays/VALORANT/ or .../Replays/Desktop/
            target_dir = os.path.join(base_dir, "Replays", app_name)
            os.makedirs(target_dir, exist_ok=True)

            target_path = os.path.join(target_dir, filename)

            try:
                shutil.move(replay_path, target_path)
                print(f"[Auto-Sorter] Sorted '{filename}' into '{app_name}'")
            except Exception as e:
                print(f"[Auto-Sorter] Error moving file: {e}")


def script_load(settings):
    obs.obs_frontend_add_event_callback(on_event)
    print("Replay Buffer Auto-Sorter (Browser-Fix) loaded.")


def script_unload():
    obs.obs_frontend_remove_event_callback(on_event)


def script_description():
    return "Sorts Replay Buffer clips. Categorizes Browsers and non-fullscreen apps as 'Desktop'. Fullscreen games get their own folder."