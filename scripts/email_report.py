"""
email_report.py  —  Sends the weekend Excel report via Gmail SMTP.

Works on Mac, Linux, and Windows.
Uses a dedicated Gmail bot account — no HPE MFA issues.

Place this file in the scripts/ folder.
Called automatically by read_mails.py at end of scan window.
"""

import json
import os
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.base      import MIMEBase
from email                import encoders
from pathlib              import Path
from datetime             import datetime


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
            "Please create it with sender_email, "
            "sender_password, and recipients."
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
# SEND VIA GMAIL SMTP
# Works on Mac, Linux, Windows
# ==========================================

def _send_via_smtp(
    sender, password, recipients,
    subject, body, excel_file
):

    msg            = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    # Attach Excel if it exists
    if excel_file and excel_file.exists():

        with open(excel_file, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())

        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={excel_file.name}"
        )
        msg.attach(part)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(sender, password)
        server.sendmail(
            sender,
            recipients,
            msg.as_string()
        )


# ==========================================
# MAIN — SEND REPORT
# ==========================================

def send_report(target_date=None):

    try:

        config     = load_config()
        sender     = config.get("sender_email", "").strip()
        password   = config.get("sender_password", "").strip()
        recipients = config.get("recipients", [])

        if not sender:
            print("Email skipped: sender_email missing in email_config.json")
            return False

        if not password:
            print("Email skipped: sender_password missing in email_config.json")
            return False

        if not recipients:
            print("Email skipped: no recipients in email_config.json")
            return False

        excel_file = get_latest_excel()

        now      = target_date or datetime.now()
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

        print(f"\nSending report email via Gmail SMTP...")
        print(f"  From: {sender}")
        print(f"  To:   {', '.join(recipients)}")
        print(f"  File: {excel_file.name if excel_file else 'none — no attachment'}")

        _send_via_smtp(
            sender,
            password,
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

    except smtplib.SMTPAuthenticationError:
        print(
            "Email failed: Gmail authentication error.\n"
            "Check sender_email and sender_password.\n"
            "Make sure you are using an App Password, not your regular Gmail password.\n"
            "Generate one at: myaccount.google.com → Security → App Passwords"
        )
        return False

    except Exception as e:
        print(f"Email sending failed: {e}")
        return False