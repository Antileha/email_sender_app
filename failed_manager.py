import json
import os

FAILED_FILE = "failed_recipients.json"


def save_failed_recipients(recipients):
    with open(FAILED_FILE, "w", encoding="utf-8") as file:
        json.dump(recipients, file, ensure_ascii=False, indent=2)


def load_failed_recipients():
    if not os.path.exists(FAILED_FILE):
        return []

    with open(FAILED_FILE, "r", encoding="utf-8") as file:
        return json.load(file)


def clear_failed_recipients():
    if os.path.exists(FAILED_FILE):
        os.remove(FAILED_FILE)