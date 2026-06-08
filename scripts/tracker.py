"""
tracker.py  —  Per-day processed-mail deduplication.

Structure of processed_mails.json:
{
    "date": "2025-06-07",       <- IST date this log belongs to
    "ids":  ["case_xxx", ...]   <- processed IDs for that day
}

On a new day (date mismatch), the old IDs are archived
to automation.log and the tracker resets fresh.
This allows the same case to be re-processed on Sunday
even if it appeared in Saturday's handover.
"""

import json
from pathlib  import Path
from datetime import datetime, timezone, timedelta

BASE_DIR     = Path(__file__).resolve().parent.parent
TRACKER_FILE = BASE_DIR / "data" / "processed_mails.json"
LOG_FILE     = BASE_DIR / "logs" / "automation.log"

IST = timezone(timedelta(hours=5, minutes=30))


def _today_str():
    return datetime.now(IST).strftime("%Y-%m-%d")


def _append_to_log(date_str: str, ids: list):
    """Archive yesterday's IDs into automation.log before reset."""
    if not ids:
        return
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a") as f:
        f.write(f"\n[TRACKER ARCHIVE — {date_str}]\n")
        for mid in ids:
            f.write(f"  {mid}\n")
        f.write(f"[END ARCHIVE — {len(ids)} entries]\n")


def _load_raw() -> dict:
    """Return raw JSON dict, or empty shell if missing/corrupt."""
    if not TRACKER_FILE.exists():
        return {"date": "", "ids": []}
    try:
        with open(TRACKER_FILE, "r") as f:
            data = json.load(f)
        # Handle old flat-list format (migration)
        if isinstance(data, list):
            return {"date": "", "ids": data}
        return data
    except Exception:
        return {"date": "", "ids": []}


def _save_raw(data: dict):
    TRACKER_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(TRACKER_FILE, "w") as f:
        json.dump(data, f, indent=4)


def reset_if_new_day():
    """
    Call once at the start of each scan cycle.
    If the tracker belongs to a previous date:
      1. Archive old IDs to automation.log
      2. Reset tracker to today with empty IDs
    """
    data  = _load_raw()
    today = _today_str()

    if data.get("date") == today:
        return  # same day — nothing to reset

    old_date = data.get("date") or "unknown"
    old_ids  = data.get("ids", [])

    print(
        f"\n[Tracker] New day detected "
        f"(was {old_date}, now {today}) — "
        f"archiving {len(old_ids)} entries and resetting."
    )

    _append_to_log(old_date, old_ids)
    _save_raw({"date": today, "ids": []})


def load_processed() -> list:
    data = _load_raw()
    # If stale date, treat as empty (safety net)
    if data.get("date") != _today_str():
        return []
    return data.get("ids", [])


def save_processed(ids: list):
    _save_raw({"date": _today_str(), "ids": ids})


def already_processed(mail_id: str) -> bool:
    return mail_id in load_processed()


def mark_processed(mail_id: str):
    ids = load_processed()
    if mail_id not in ids:
        ids.append(mail_id)
        save_processed(ids)