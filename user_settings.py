import os
import json
import getpass

# ----------- Constants -----------
USER_DATA_FOLDER = "../user_data"
CHECKLIST_SAVE_PATH = r"P:\ENGINEERING\Design Checklist\json_files"

# ----------- User Configuration ----------- 

def get_username():
    """Return the current user's login name."""
    try:
        return os.getlogin()
    except Exception:
        # Fallback if os.getlogin() fails (sometimes on services)
        return getpass.getuser()

def get_user_config_path():
    """Get the path to the current user's settings file."""
    username = get_username()
    os.makedirs(USER_DATA_FOLDER, exist_ok=True)
    return os.path.join(USER_DATA_FOLDER, f"{username}.json")

def load_user_settings():
    """Load user settings from the user's config file."""
    config_path = get_user_config_path()
    if os.path.exists(config_path):
        with open(config_path, "r") as f:
            return json.load(f)
    return {}

def save_user_settings(settings):
    """Save user settings to the user's config file."""
    config_path = get_user_config_path()
    with open(config_path, "w") as f:
        json.dump(settings, f, indent=4)

# ----------- Checklist File Operations -----------

def save_combined_data(path, data):
    data["last_user"] = getpass.getuser()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, separators=(",", ":"))

def load_combined_data(path):
    """Load checklist data from the given path. Returns dict or None."""
    if path and os.path.exists(path):
        try:
            with open(path, "r") as f:
                return json.load(f)
        except Exception as e:
            # Don't handle errors here; let UI show errors if needed
            return None
    return None

