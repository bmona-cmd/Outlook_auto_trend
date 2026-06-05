"""
email_report.py  —  Sends the weekend Excel report via Outlook desktop app
                    on Mac using AppleScript (osascript).

No SMTP. No password. No App Password needed.
Uses your already-logged-in Outlook session on Mac.

Place this file in the scripts/ folder.
Called automatically by read_mails.py at end of scan window.
"""

import json
import os
import subprocess

from pathlib  import Path
from datetime import datetime


BASE_DIR    = Path(__file__).resolve().parent.parent
CONFIG_FILE = BASE_DIR / "email_config.json"
OUTPUT_DIR  = BASE_DIR / "output"


# ==========================================
# LOAD CONFIG
# ==========================================

def load_config():

    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"email_config.json not found at {CONFIG_FILE}\n"
            "Please create it with sender_email and recipients."
        )

    with open(CONFIG_FILE, "r") as f:
        return json.load(f)


# ==========================================
# FIND LATEST EXCEL
# ==========================================

def get_latest_excel():

    if not OUTPUT_DIR.exists():
        return None

    files = list(OUTPUT_DIR.glob("*.xlsx"))

    if not files:
        return None

    return max(files, key=lambda f: f.stat().st_mtime)


# ==========================================
# BUILD EMAIL BODY
# ==========================================

def build_body(day_name, date_str, excel_file):

    if excel_file:
        return (
            f"Hi,\n\n"
            f"Please find attached the weekend cases report "
            f"for {day_name}, {date_str}.\n\n"
            f"File: {excel_file.name}\n\n"
            f"This email was sent automatically by the "
            f"Weekend Mail Automation.\n\n"
            f"Regards,\n"
            f"Weekend Automation Bot"
        )

    return (
        f"Hi,\n\n"
        f"The weekend automation completed for "
        f"{day_name}, {date_str}, "
        f"but no Excel file was found in the output folder.\n\n"
        f"This may mean no cases were processed today.\n\n"
        f"Regards,\n"
        f"Weekend Automation Bot"
    )


# ==========================================
# ESCAPE FOR APPLESCRIPT
# ==========================================

def _esc(text):
    text = str(text)
    text = text.replace("\\", "\\\\")
    text = text.replace('"', '\\"')
    text = text.replace("\n", "\" & (ASCII character 10) & \"")
    return text


def _as_applescript_text(text):
    return f'"{_esc(text)}"'


# ==========================================
# SEND VIA APPLESCRIPT (Outlook on Mac)
# ==========================================

def _send_via_applescript(recipients, subject, body, excel_file):

    recipient_lines = "\n    ".join([
        "make new recipient at theMessage with properties "
        f"{{email address:{{address:{_as_applescript_text(r.strip())}}}}}"
        for r in recipients
    ])

    if excel_file and excel_file.exists():
        attachment_line = (
            "make new attachment at theMessage with properties "
            f"{{file:(POSIX file {_as_applescript_text(str(excel_file))})}}"
        )
    else:
        attachment_line = ""

    # NOTE: removed 'synchronize object model server' —
    # not supported in all Outlook versions.
    # 'send theMessage' moves it to Outbox;
    # 'delay 3' gives Outlook time to dispatch it automatically.
    script = (
        'tell application "Microsoft Outlook"\n'
        '    activate\n'
        '    set theMessage to make new outgoing message with properties'
        ' {subject:' + _as_applescript_text(subject) + ', '
        'plain text content:' + _as_applescript_text(body) + '}\n'
        '    ' + recipient_lines + '\n'
        '    ' + attachment_line + '\n'
        '    send theMessage\n'
        '    delay 3\n'
        'end tell'
    )

    if os.environ.get("EMAIL_REPORT_DEBUG_APPLESCRIPT"):
        print("\nGenerated AppleScript:\n")
        print(script)

    result = subprocess.run(
        ["osascript"],
        input=script,
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"AppleScript error: {result.stderr.strip()}"
        )

    return True


# ==========================================
# MAIN — SEND REPORT
# ==========================================

def send_report():

    try:

        config     = load_config()
        recipients = config.get("recipients", [])

        if not recipients:
            print("Email skipped: no recipients in email_config.json")
            return False

        excel_file = get_latest_excel()

        now      = datetime.now()
        weekday  = now.weekday()
        day_name = (
            "Saturday" if weekday == 5
            else "Sunday" if weekday == 6
            else now.strftime("%A")
        )
        date_str = now.strftime("%d %b %Y")

        subject = (
            f"Weekend Cases Report — "
            f"{day_name} {date_str}"
        )

        body = build_body(day_name, date_str, excel_file)

        print(f"\nSending report email via Outlook...")
        print(f"  To:   {', '.join(recipients)}")
        print(f"  File: {excel_file.name if excel_file else 'none — no attachment'}")

        _send_via_applescript(
            recipients,
            subject,
            body,
            excel_file
        )

        print(
            f"Report email sent successfully ✓\n"
            f"  To:   {', '.join(recipients)}\n"
            f"  File: {excel_file.name if excel_file else 'none'}"
        )

        return True

    except FileNotFoundError as e:
        print(f"Email config error: {e}")
        return False

    except RuntimeError as e:
        print(f"Outlook send failed: {e}")
        print(
            "Make sure Microsoft Outlook is installed "
            "and you are logged in."
        )
        return False

    except Exception as e:
        print(f"Email sending failed: {e}")
        return False