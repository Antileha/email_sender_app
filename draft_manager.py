import json
import os
from datetime import datetime

DRAFTS_DIR = "drafts"


def ensure_drafts_dir():
    os.makedirs(DRAFTS_DIR, exist_ok=True)


def save_draft(data):
    ensure_drafts_dir()

    created_at = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"draft_{created_at}.json"
    path = os.path.join(DRAFTS_DIR, filename)

    data["created_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)

    return path


def load_draft(path):
    with open(path, "r", encoding="utf-8") as file:
        return json.load(file)