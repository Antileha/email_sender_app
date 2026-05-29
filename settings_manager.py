import json
import os

SETTINGS_FILE = "settings.json"

DEFAULT_SETTINGS = {
    "default_cc": "",
    "default_bcc": "",
    "last_sender_email": "",
    "last_smtp_service": "Yandex",
    "delay_seconds": 1
}


def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        return DEFAULT_SETTINGS.copy()

    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)

        settings = DEFAULT_SETTINGS.copy()
        settings.update(data)
        return settings

    except Exception:
        return DEFAULT_SETTINGS.copy()


def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as file:
        json.dump(settings, file, ensure_ascii=False, indent=2)