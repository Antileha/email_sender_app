import json
import os

DRAFTS_DIR = "drafts"


def ensure_drafts_dir():
    os.makedirs(DRAFTS_DIR, exist_ok=True)


def save_draft(data, file_path):
    ensure_drafts_dir()

    if not file_path.lower().endswith(".json"):
        file_path += ".json"

    with open(file_path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)

    return file_path


def load_draft(path):
    with open(path, "r", encoding="utf-8") as file:
        return json.load(file)