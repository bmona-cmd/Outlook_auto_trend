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

# ── project imports (unchanged) ──────────────────────────────────────────────
import scripts.read_mails as mail_reader

BASE_DIR     = Path(__file__).resolve().parent
MAPPING_FILE = BASE_DIR / "customer_vertical_mapping.xlsx"
DEVICE_FILE  = BASE_DIR / "data" / "custom_devices.json"
OUTPUT_DIR   = BASE_DIR / "output"

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

# ── patch read_mails to also push to log_buffer ───────────────────────────────
_orig_print = builtins.print

def _log_print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    push_log(msg)
    _orig_print(*args, **kwargs)

mail_reader.print = _log_print   # redirect mail-reader prints to our log too


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
        "latest_file": xl.name if xl else "—"
    })


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
