import ctypes
import os
import re
import shutil

import obspython as obs
import psutil
import win32api
import win32gui
import win32process

# Constants for SHQueryUserNotificationState (from shellapi.h)
QUNS_BUSY = 2
QUNS_PRESENTATION_MODE = 3
QUNS_RUNNING_D3D_FULL_SCREEN = 4


def is_fullscreen_state():
    """
    Checks via the Windows API if a game/app is running in fullscreen mode.
    This mimics the behavior of SHQueryUserNotificationState to detect true fullscreen.
    """
    state = ctypes.c_int()
    try:
        # Call the Windows Shell API
        ctypes.windll.shell32.SHQueryUserNotificationState(ctypes.byref(state))

        # Check if the state matches any of the known fullscreen/presentation states
        if state.value in (
            QUNS_BUSY,
            QUNS_PRESENTATION_MODE,
            QUNS_RUNNING_D3D_FULL_SCREEN,
        ):
            return True
    except Exception as e:
        print(f"Error checking fullscreen state: {e}")
    return False


def is_on_primary_display(hwnd):
    """
    Checks if the specified window covers the entire primary monitor.
    """
    try:
        # SM_CXSCREEN = 0, SM_CYSCREEN = 1 (Primary monitor width and height)
        monitor_width = win32api.GetSystemMetrics(0)
        monitor_height = win32api.GetSystemMetrics(1)

        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect

        if (
            left <= 0
            and top <= 0
            and right >= monitor_width
            and bottom >= monitor_height
        ):
            return True
    except Exception:
        pass
    return False


def get_clean_window_name(hwnd):
    """
    Retrieves the window title (e.g., 'VALORANT') and strips out invalid folder characters.
    """
    title = win32gui.GetWindowText(hwnd)

    # Fallback to the .exe process name if the window has no title
    if not title:
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            title = psutil.Process(pid).name().replace(".exe", "")
        except Exception:
            title = "Unknown"

    # Regex: Removes characters that are not allowed in Windows folder names (\ / : * ? " < > |)
    clean_title = re.sub(r'[\\/*?:"<>|]', "", title)
    return clean_title.strip()


def get_active_context():
    """
    Determines the current context. If the active window is fullscreen and on
    the primary display, it is assumed to be a game. Otherwise, it defaults to 'Desktop'.
    """
    hwnd = win32gui.GetForegroundWindow()

    if not is_fullscreen_state() or not is_on_primary_display(hwnd):
        return "Desktop"

    return get_clean_window_name(hwnd)


def on_event(event):
    """
    Callback function triggered by OBS frontend events.
    Listens for the replay buffer save event and sorts the file accordingly.
    """
    if event == obs.OBS_FRONTEND_EVENT_REPLAY_BUFFER_SAVED:
        app_name = get_active_context()
        replay_path = obs.obs_frontend_get_last_replay()

        if replay_path and os.path.exists(replay_path):
            base_dir = os.path.dirname(replay_path)
            filename = os.path.basename(replay_path)

            # Build the target directory path (e.g., "Videos/Replays/Desktop" or "Videos/Replays/VALORANT")
            target_dir = os.path.join(base_dir, "Replays", app_name)
            os.makedirs(target_dir, exist_ok=True)

            target_path = os.path.join(target_dir, filename)

            try:
                shutil.move(replay_path, target_path)
                print(f"Clip successfully sorted to: {target_path}")
            except Exception as e:
                print(f"Error moving the clip: {e}")


def script_load(settings):
    """Called when the script is loaded into OBS."""
    obs.obs_frontend_add_event_callback(on_event)
    print("Replay Buffer Advanced Auto-Sorter loaded.")


def script_unload():
    """Called when the script is unloaded from OBS."""
    obs.obs_frontend_remove_event_callback(on_event)
    print("Replay Buffer Advanced Auto-Sorter unloaded.")


def script_description():
    """Returns the description of the script for the OBS UI."""
    return "Automatically sorts Replay Buffer clips into subfolders. Uses the Windows API to detect true fullscreen applications (games) and names the folder after the window title. Defaults to 'Desktop' otherwise."
