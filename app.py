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
email_thread      = None

# ── patch read_mails to also push to log_buffer ───────────────────────────────
_orig_print = builtins.print

def _log_print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    push_log(msg)
    _orig_print(*args, **kwargs)

mail_reader.print  = _log_print
email_report.print = _log_print


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS  (all original — untouched)
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
# ROUTES — PAGES  (original — untouched)
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    xl = latest_excel()
    return render_template("index.html",
                           latest_file=xl.name if xl else "—",
                           running=mail_reader.RUNNING)


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES — AUTOMATION API  (original — untouched)
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
    if not alive and mail_reader.RUNNING:
        mail_reader.RUNNING = False
    with log_lock:
        logs = list(log_buffer)
    xl = latest_excel()
    return jsonify({
        "running":       mail_reader.RUNNING,
        "logs":          logs,
        "latest_file":   xl.name if xl else "—",
        "email":         email_status(),
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
# ROUTES — EMAIL RECIPIENTS API  (original — untouched)
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
# ROUTES — EXCEL  (original — untouched)
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/api/download_excel")
def download_excel():
    xl = latest_excel()
    if not xl:
        return jsonify({"error": "No Excel file found"}), 404
    return send_file(xl, as_attachment=True, download_name=xl.name)


# ═════════════════════════════════════════════════════════════════════════════
# ROUTES — CUSTOMERS API  (original — untouched)
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
# ROUTES — DEVICES API  (original — untouched)
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
# ── NEW: CHART DATA API ───────────────────────────────────────────────────
# Reads the latest Excel output file and returns count breakdowns.
# Does NOT touch read_mails.py, excel_writer.py, or any existing logic.
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/api/chart_data")
def api_chart_data():
    output_files = sorted(OUTPUT_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not output_files:
        return jsonify({
            "file": None,
            "total": 0,
            "latest": {"vertical": {}, "technology": {}, "delivery_type": {}},
            "weekly": {"labels": [], "vertical": {}, "technology": {}, "delivery_type": {}, "totals": []},
            "monthly": {"labels": [], "vertical": {}, "technology": {}, "delivery_type": {}, "totals": []}
        })
    try:
        import pandas as pd

        DEFAULT_VERTICAL = ["EMEA", "Cable", "Content", "Enterprise", "Telco", "Software"]
        DEFAULT_TECHNOLOGY = ["Routing", "Switching", "Security", "Software"]
        DEFAULT_DELIVERY = ["Dispatch P1", "Dispatch P2", "Handover"]

        def normalize_counts(df, col, defaults=None):
            if col not in df.columns:
                return {str(k): 0 for k in defaults} if defaults else {}
            s = (
                df[col].dropna().astype(str).str.strip()
                .replace("", None).dropna()
            )
            counts = {str(k): int(v) for k, v in s.value_counts().to_dict().items()}
            result = {}
            if defaults:
                for k in defaults:
                    result[str(k)] = counts.pop(k, 0)
            for k, v in sorted(counts.items()):
                result[str(k)] = int(v)
            return result

        def _parse_date_column(series):
            dates = series.astype(str).str.strip()
            parsed = pd.to_datetime(dates, format="%d-%b-%y", errors="coerce")
            if parsed.isna().all():
                parsed = pd.to_datetime(dates, errors="coerce")
            return parsed

        latest_file = output_files[-1]
        latest_df = pd.read_excel(latest_file, engine="openpyxl")

        dfs = []
        for excel_file in output_files:
            try:
                df = pd.read_excel(excel_file, engine="openpyxl")
                if "Date" in df.columns:
                    df["Date"] = _parse_date_column(df["Date"])
                else:
                    df["Date"] = pd.NaT
                dfs.append(df)
            except Exception:
                continue

        if not dfs:
            return jsonify({
                "file": latest_file.name,
                "total": len(latest_df),
                "latest": {
                    "vertical": normalize_counts(latest_df, "Vertical"),
                    "technology": normalize_counts(latest_df, "Technology"),
                    "delivery_type": normalize_counts(latest_df, "Case Delivery Type")
                },
                "weekly": {"labels": [], "vertical": {}, "technology": {}, "delivery_type": {}, "totals": []},
                "monthly": {"labels": [], "vertical": {}, "technology": {}, "delivery_type": {}, "totals": []}
            })

        combined = pd.concat(dfs, ignore_index=True)
        if "Date" in combined.columns:
            combined["Date"] = _parse_date_column(combined["Date"])
        else:
            combined["Date"] = pd.NaT

        combined = combined[combined["Date"].notna()].copy()

        def build_timeseries(df, group_col, category_col, defaults=None, labels=None, order_col=None):
            if category_col not in df.columns or df.empty:
                return {"labels": labels or [], "series": {}, "totals": []}

            grouped = (
                df.groupby([group_col, category_col]).size()
                  .unstack(fill_value=0)
            )
            categories = list(defaults) if defaults else []
            for category in grouped.columns.astype(str):
                if category not in categories:
                    categories.append(category)
            grouped = grouped.reindex(columns=categories, fill_value=0)

            if labels is not None:
                grouped = grouped.reindex(index=labels, fill_value=0)
            elif order_col and order_col in df.columns:
                ordered = (
                    df[[group_col, order_col]]
                      .drop_duplicates()
                      .sort_values(order_col)
                )
                ordered_labels = ordered[group_col].astype(str).tolist()
                grouped = grouped.reindex(index=ordered_labels, fill_value=0)
            else:
                grouped = grouped.sort_index()

            labels_out = [str(l) for l in grouped.index]
            series = {
                str(category): [int(grouped.at[label, category]) for label in grouped.index]
                for category in categories
            }
            totals = [int(x) for x in grouped.sum(axis=1).tolist()]
            return {"labels": labels_out, "series": series, "totals": totals}

        if not combined.empty:
            combined["date_label"] = combined["Date"].dt.strftime("%a %d %b")
            combined["date_order"] = combined["Date"]
            combined["weekday"] = combined["Date"].dt.weekday
            combined["day_name"] = combined["Date"].dt.strftime("%a")
            weekend_start = (
                combined["Date"] - pd.to_timedelta((combined["Date"].dt.weekday - 5) % 7, unit="d")
            )
            combined["weekend_start"] = weekend_start
            combined["week_of_month"] = (combined["Date"].dt.day - 1) // 7 + 1
            combined["week_label"] = "Week " + combined["week_of_month"].astype(str)

            latest_weekend = combined["weekend_start"].max()
            weekly_df = combined[combined["weekend_start"] == latest_weekend].copy()
            weekly_labels = ["Sat", "Sun"]
            monthly_labels = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"]
        else:
            combined["weekend"] = []
            combined["date_order"] = []
            combined["weekend_start"] = []
            combined["weekday"] = []
            combined["day_name"] = []
            combined["week_label"] = []
            weekly_df = combined
            weekly_labels = ["Sat", "Sun"]
            monthly_labels = ["Week 1", "Week 2", "Week 3", "Week 4"]

        weekly = {
            "vertical": build_timeseries(weekly_df, "day_name", "Vertical", DEFAULT_VERTICAL, labels=weekly_labels, order_col="weekday"),
            "technology": build_timeseries(weekly_df, "day_name", "Technology", DEFAULT_TECHNOLOGY, labels=weekly_labels, order_col="weekday"),
            "delivery_type": build_timeseries(weekly_df, "day_name", "Case Delivery Type", DEFAULT_DELIVERY, labels=weekly_labels, order_col="weekday")
        }
        monthly = {
            "vertical": build_timeseries(combined, "week_label", "Vertical", DEFAULT_VERTICAL, labels=monthly_labels),
            "technology": build_timeseries(combined, "week_label", "Technology", DEFAULT_TECHNOLOGY, labels=monthly_labels),
            "delivery_type": build_timeseries(combined, "week_label", "Case Delivery Type", DEFAULT_DELIVERY, labels=monthly_labels)
        }

        return jsonify({
            "file": latest_file.name,
            "total": len(latest_df),
            "latest": {
                "vertical": normalize_counts(latest_df, "Vertical"),
                "technology": normalize_counts(latest_df, "Technology"),
                "delivery_type": normalize_counts(latest_df, "Case Delivery Type")
            },
            "weekly": weekly,
            "monthly": monthly
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ═════════════════════════════════════════════════════════════════════════════
# ── NEW: ADD CHARTS TO EXCEL ──────────────────────────────────────────────
# Triggered manually from the dashboard. Adds a Charts sheet to the Excel.
# Does NOT change excel_writer.py existing functions.
# ═════════════════════════════════════════════════════════════════════════════

@app.route("/api/add_charts_to_excel", methods=["POST"])
def api_add_charts_to_excel():
    xl = latest_excel()
    if not xl:
        return jsonify({"ok": False, "msg": "No Excel file found in output folder"}), 404
    try:
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.chart import BarChart, Reference
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        df = pd.read_excel(xl, engine="openpyxl")
        wb = load_workbook(xl)

        if "Charts" in wb.sheetnames:
            del wb["Charts"]
        ws = wb.create_sheet("Charts")

        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 12

        COLORS = ["4472C4", "ED7D31", "A9D18E", "FF0000", "FFC000", "70AD47", "9E480E", "997300"]
        thin   = Side(style="thin", color="BFBFBF")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        def write_block(col_name, title, start_row):
            if col_name not in df.columns:
                return start_row
            counts = (
                df[col_name].dropna().astype(str).str.strip()
                .replace("", None).dropna()
                .value_counts().sort_index()
            )
            if counts.empty:
                return start_row

            # Section title
            t = ws.cell(row=start_row, column=1, value=title)
            t.font = Font(bold=True, size=13, color="1F3864", name="Arial")
            start_row += 1

            # Header
            for c, val in [(1, col_name), (2, "Count")]:
                cell = ws.cell(row=start_row, column=c, value=val)
                cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
                cell.fill      = PatternFill("solid", start_color="4472C4")
                cell.alignment = Alignment(horizontal="center")
                cell.border    = border

            for i, (label, count) in enumerate(counts.items(), 1):
                r = start_row + i
                lc = ws.cell(row=r, column=1, value=label)
                cc = ws.cell(row=r, column=2, value=int(count))
                for cell in [lc, cc]:
                    cell.font      = Font(name="Arial", size=10)
                    cell.border    = border
                    cell.alignment = Alignment(horizontal="left" if cell.column == 1 else "center")
                    if i % 2 == 0:
                        cell.fill = PatternFill("solid", start_color="DCE6F1")

            data_end = start_row + len(counts)

            # Chart
            chart = BarChart()
            chart.type     = "col"
            chart.grouping = "clustered"
            chart.title    = title
            chart.y_axis.title = "Cases"
            chart.width    = 18
            chart.height   = 12
            chart.style    = 10
            chart.add_data(Reference(ws, min_col=2, max_col=2,
                                     min_row=start_row, max_row=data_end),
                           titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=1,
                                           min_row=start_row+1, max_row=data_end))
            chart.series[0].graphicalProperties.solidFill = COLORS[0]
            ws.add_chart(chart, f"D{start_row}")

            return data_end + 24   # leave room for chart height

        row = 1
        row = write_block("Vertical",           "Cases by Vertical",           row)
        row = write_block("Technology",         "Cases by Technology",         row)
        row = write_block("Case Delivery Type", "Cases by Case Delivery Type", row)

        wb.save(xl)
        push_log(f"Charts sheet added to {xl.name} ✓")
        return jsonify({"ok": True, "msg": f"Charts sheet added to {xl.name}"})

    except Exception as e:
        push_log(f"Charts to Excel failed: {e}")
        return jsonify({"ok": False, "msg": str(e)}), 500


# ═════════════════════════════════════════════════════════════════════════════
# STARTUP — load custom devices into parser  (original — untouched)
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