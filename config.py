import os
import json

CONFIG_FILE = "config.json"

DEFAULT_CONFIG = {
    "outlook_folder": "HP Scan",
    "sender_filter": "scanner@company.com",
    "subject_filter": "onboarding, spare, replacement",
    "rename_format": "onboarding_{date}_{filename}",
    "outlook_source": "local",  # "local" (Classic Outlook) or "graph" (New Outlook / Cloud)
    "sync_method": "graph",  # "graph" or "local"
    "sharepoint_site": "https://company.sharepoint.com/sites/yoursite",
    "sharepoint_folder": "Shared Documents/Backups/2026",
    "local_sync_path": "",
    "client_id": "your-azure-client-id-uuid",
    "tenant_id": "common",
    "check_interval": 15,
    "run_minimized": False
}

def get_config_path():
    # If compiled as exe, get the directory of the exe
    import sys
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, CONFIG_FILE)

def load_config():
    path = get_config_path()
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Merge with default config to ensure new fields are present
                merged = DEFAULT_CONFIG.copy()
                merged.update(config)
                return merged
        except Exception as e:
            print(f"Error loading config, using defaults: {e}")
            return DEFAULT_CONFIG.copy()
    else:
        return DEFAULT_CONFIG.copy()

def save_config(config_data):
    path = get_config_path()
    try:
        # Filter config data to only write expected keys
        to_save = {}
        for key in DEFAULT_CONFIG:
            if key in config_data:
                to_save[key] = config_data[key]
            else:
                to_save[key] = DEFAULT_CONFIG[key]
                
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(to_save, f, indent=4)
        return True
    except Exception as e:
        print(f"Error saving config: {e}")
        return False
