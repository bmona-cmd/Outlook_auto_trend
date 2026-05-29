import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

TRACKER_FILE = (
    BASE_DIR
    / "data"
    / "processed_mails.json"
)


def load_processed():

    if not TRACKER_FILE.exists():
        return []

    try:
        with open(TRACKER_FILE, "r") as f:
            return json.load(f)

    except:
        return []


def save_processed(processed):

    TRACKER_FILE.parent.mkdir(exist_ok=True)

    with open(TRACKER_FILE, "w") as f:
        json.dump(processed, f, indent=4)


def already_processed(mail_id):

    processed = load_processed()

    return mail_id in processed


def mark_processed(mail_id):

    processed = load_processed()

    if mail_id not in processed:

        processed.append(mail_id)

        save_processed(processed)
