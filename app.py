"""
app.py  —  Flask web UI for Weekend Mail Automation
Place this file in the PROJECT ROOT (same level as run.py and scripts/).
Run:  python3 app.py
Then open:  http://localhost:5050
"""

from flask import Flask, render_template, jsonify, request, send_file
from pathlib import Path
import threading
import json
import builtins
import re

# ── project imports (unchanged) ──────────────────────────────────────────────
import scripts.read_mails as mail_reader
import scripts.email_report as email_report

BASE_DIR     = Path(__file__).resolve().parent
MAPPING_FILE = BASE_DIR / "customer_vertical_mapping.xlsx"
DEVICE_FILE  = BASE_DIR / "data" / "custom_devices.json"
OUTPUT_DIR   = BASE_DIR / "output"
EMAIL_RE     = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

app = Flask(__name__)

# ── in-memory log (thread-safe append) ───────────────────────────────────────
from collections import deque
import datetime

log_buffer = deque(maxlen=200)
log_lock   = threading.Lock()

def push_log(msg: str):
    ts   = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}]  {msg}"
    with log_lock:
        log_buffer.append(line)

automation_thread = None
email_thread = None

# ── patch read_mails to also push to log_buffer ───────────────────────────────
_orig_print = builtins.print

def _log_print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    push_log(msg)
    _orig_print(*args, **kwargs)

mail_reader.print = _log_print   # redirect mail-reader prints to our log too
email_report.print = _log_print   # redirect email sender prints to our log too


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def load_custom_devices() -> dict:
    DEVICE_FILE.parent.mkdir(exist_ok=True)
    if not DEVICE_FILE.exists():
        return {}
    try:
        return json.loads(DEVICE_FILE.read_text())
    except Exception:
        return {}

def save_custom_devices(data: dict):
    DEVICE_FILE.parent.mkdir(exist_ok=True)
    DEVICE_FILE.write_text(json.dumps(data, indent=2))

def load_customers() -> list:
    if not MAPPING_FILE.exists():
        return []
    try:
        import pandas as pd
        df = pd.read_excel(MAPPING_FILE, engine="openpyxl")
        df.columns = [str(c).strip().lower() for c in df.columns]
        cc = next((c for c in df.columns if "customer" in c or "company" in c), None)
        vc = next((c for c in df.columns if "vertical" in c), None)
        if not cc or not vc:
            return []
        return [{"customer": str(r[cc]).strip(), "vertical": str(r[vc]).strip()}
                for _, r in df.iterrows()
                if str(r[cc]).strip() not in ("", "nan")]
    except Exception:
        return []

def append_customer(customer: str, vertical: str) -> bool:
    try:
        import pandas as pd
        from openpyxl import load_workbook
        if not MAPPING_FILE.exists():
            pd.DataFrame(columns=["Customer", "Vertical"]).to_excel(
                MAPPING_FILE, index=False, engine="openpyxl")
        wb = load_workbook(MAPPING_FILE)
        wb.active.append([customer, vertical])
        wb.save(MAPPING_FILE)
        try:
            from scripts.vertical_lookup import load_vertical_mapping
            load_vertical_mapping()
        except Exception:
            pass
        return True
    except Exception as e:
        return False

def latest_excel():
    if not OUTPUT_DIR.exists():
        return None
    files = list(OUTPUT_DIR.glob("*.xlsx"))
    return max(files, key=lambda f: f.stat().st_mtime) if files else None

def email_status() -> dict:
    try:
        recipients = load_email_recipients()
        return {
            "configured": bool(recipients),
            "recipients": recipients,
            "message": (
                f"{len(recipients)} recipient(s) configured"
                if recipients else
                "No recipients configured"
            )
        }
    except FileNotFoundError:
        return {
            "configured": False,
            "recipients": [],
            "message": "email_config.json not found"
        }
    except Exception as e:
        return {
            "configured": False,
            "recipients": [],
            "message": f"Email config error: {e}"
        }

def load_email_config() -> dict:
    if not email_report.CONFIG_FILE.exists():
        return {"sender_email": "", "recipients": []}
    try:
        data = json.loads(email_report.CONFIG_FILE.read_text())
    except Exception:
        data = {}
    data.setdefault("sender_email", "")
    data.setdefault("recipients", [])
    return data

def save_email_config(data: dict):
    email_report.CONFIG_FILE.write_text(json.dumps(data, indent=4))

def load_email_recipients() -> list:
    config = load_email_config()
    seen = set()
    recipients = []
    for email in config.get("recipients", []):
        email = str(email).strip()
        key = email.lower()
        if email and key not in seen:
            recipients.append(email)
            seen.add(key)
    return recipients

BUILTIN_DEVICES = {
    "mx":"Routing","ptx":"Routing","acx":"Routing",
    "srx":"Security","ssg":"Security",
    "qfx":"Switching","ex":"Switching",
    "mist":"Wireless","128t":"SDWAN","software":"Software",
}


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES — PAGES
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    xl = latest_excel()
    return render_template("index.html",
                           latest_file=xl.name if xl else "—",
                           running=mail_reader.RUNNING)


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES — AUTOMATION API
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/api/start", methods=["POST"])
def api_start():
    global automation_thread
    if automation_thread and automation_thread.is_alive():
        return jsonify({"ok": False, "msg": "Already running"})
    mail_reader.RUNNING = True
    push_log("Automation started.")
    automation_thread = threading.Thread(
        target=mail_reader.run_mail_reader, daemon=True)
    automation_thread.start()
    return jsonify({"ok": True})

@app.route("/api/stop", methods=["POST"])
def api_stop():
    mail_reader.RUNNING = False
    push_log("Automation stopped by user.")
    return jsonify({"ok": True})

@app.route("/api/status")
def api_status():
    email_alive = bool(email_thread and email_thread.is_alive())
    alive = bool(automation_thread and automation_thread.is_alive())
    # sync RUNNING flag if thread died naturally
    if not alive and mail_reader.RUNNING:
        mail_reader.RUNNING = False
    with log_lock:
        logs = list(log_buffer)
    xl = latest_excel()
    return jsonify({
        "running": mail_reader.RUNNING,
        "logs":    logs,
        "latest_file": xl.name if xl else "—",
        "email": email_status(),
        "email_sending": email_alive
    })

@app.route("/api/send_report", methods=["POST"])
def api_send_report():
    global email_thread

    if email_thread and email_thread.is_alive():
        return jsonify({"ok": False, "msg": "Email send already in progress"}), 409

    status = email_status()
    if not status["configured"]:
        push_log(status["message"])
        return jsonify({"ok": False, "msg": status["message"]}), 400

    def _send():
        push_log("Manual report email requested from dashboard.")
        ok = email_report.send_report()
        if ok:
            push_log("Manual report email completed.")
        else:
            push_log("Manual report email failed.")

    email_thread = threading.Thread(target=_send, daemon=True)
    email_thread.start()
    return jsonify({"ok": True, "msg": "Email send started"})


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES — EMAIL RECIPIENTS API
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/api/email_recipients")
def api_email_recipients():
    return jsonify(load_email_recipients())

@app.route("/api/email_recipients", methods=["POST"])
def api_add_email_recipient():
    data = request.get_json() or {}
    email = (data.get("email") or "").strip()

    if not email:
        return jsonify({"ok": False, "msg": "Email address is required"}), 400

    if not EMAIL_RE.match(email):
        return jsonify({"ok": False, "msg": "Enter a valid email address"}), 400

    config = load_email_config()
    recipients = load_email_recipients()

    if email.lower() in {r.lower() for r in recipients}:
        return jsonify({"ok": False, "msg": "Recipient already exists"}), 400

    recipients.append(email)
    config["recipients"] = recipients
    save_email_config(config)
    push_log(f"Email recipient added: {email}")
    return jsonify({"ok": True})

@app.route("/api/email_recipients/remove", methods=["POST"])
def api_remove_email_recipient():
    data = request.get_json() or {}
    email = (data.get("email") or "").strip()

    recipients = load_email_recipients()
    kept = [r for r in recipients if r.lower() != email.lower()]

    if len(kept) == len(recipients):
        return jsonify({"ok": False, "msg": "Recipient not found"}), 404

    config = load_email_config()
    config["recipients"] = kept
    save_email_config(config)
    push_log(f"Email recipient removed: {email}")
    return jsonify({"ok": True})


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES — EXCEL
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/api/download_excel")
def download_excel():
    xl = latest_excel()
    if not xl:
        return jsonify({"error": "No Excel file found"}), 404
    return send_file(xl, as_attachment=True, download_name=xl.name)


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES — CUSTOMERS API
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/api/customers")
def api_customers():
    return jsonify(load_customers())

@app.route("/api/customers", methods=["POST"])
def api_add_customer():
    data     = request.get_json()
    customer = (data.get("customer") or "").strip()
    vertical = (data.get("vertical") or "").strip()
    if not customer or not vertical:
        return jsonify({"ok": False, "msg": "Customer and vertical are required"}), 400
    ok = append_customer(customer, vertical)
    if ok:
        push_log(f"Customer added: {customer} → {vertical}")
    return jsonify({"ok": ok})


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES — DEVICES API
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/api/devices")
def api_devices():
    builtin = [{"keyword": k, "technology": v, "source": "built-in"}
               for k, v in BUILTIN_DEVICES.items()]
    custom  = [{"keyword": k, "technology": v, "source": "custom"}
               for k, v in load_custom_devices().items()]
    return jsonify(builtin + custom)

@app.route("/api/devices", methods=["POST"])
def api_add_device():
    data    = request.get_json()
    keyword = (data.get("keyword") or "").strip().lower()
    tech    = (data.get("technology") or "").strip()
    if not keyword or not tech:
        return jsonify({"ok": False, "msg": "Keyword and technology are required"}), 400
    if keyword in BUILTIN_DEVICES:
        return jsonify({"ok": False, "msg": f"'{keyword}' is a built-in device"}), 400
    customs = load_custom_devices()
    customs[keyword] = tech
    save_custom_devices(customs)
    try:
        from scripts import parser as p
        p.DEVICE_TECH_MAP[keyword] = tech
    except Exception:
        pass
    push_log(f"Device added: {keyword} → {tech}")
    return jsonify({"ok": True})


# ═════════════════════════════════════════════════════════════════════════════
# STARTUP — load custom devices into parser
# ═════════════════════════════════════════════════════════════════════════════

def _startup():
    try:
        from scripts import parser as p
        for kw, tech in load_custom_devices().items():
            p.DEVICE_TECH_MAP[kw] = tech
    except Exception:
        pass
    push_log("Web UI ready. Open http://localhost:5050 in your browser.")

_startup()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False, use_reloader=False)
