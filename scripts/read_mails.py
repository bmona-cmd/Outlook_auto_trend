from playwright.sync_api import sync_playwright

from scripts.parser import (
    extract_case_details,
    should_skip_mail
)

from scripts.excel_writer import append_to_excel

from scripts.tracker import (
    already_processed,
    mark_processed
)

from scripts.logger import logger

from pathlib import Path

from datetime import datetime, timezone, timedelta

import time

import re


BASE_DIR = Path(__file__).resolve().parent.parent

AUTH_FILE = (

    BASE_DIR
    / "auth"
    / "auth.json"

)

CHROME_PATH = (

    "/Applications/Google Chrome.app/"
    "Contents/MacOS/Google Chrome"

)


KEYWORDS = [
    "dispatch",
    "handover",
    "[ho]",
    "ho:",
    "ho created",
    "-ho",
    "[ho-mw]",
    "ho-mw"
]


BODY_SELECTORS = [

    "div[role='document']",

    "div[aria-label='Message body']",

    "div[data-app-section='MailReadCompose']"
]


# ==========================================
# IST TIMEZONE
# ==========================================

IST = timezone(timedelta(hours=5, minutes=30))

# Window: 06:30 – 18:30 IST (both Saturday and Sunday)

WINDOW_START_H = 6
WINDOW_START_M = 30
WINDOW_END_H   = 18
WINDOW_END_M   = 30


# ==========================================
# TEST MODE
# Set to False on manager's device for live run.
# ==========================================

TEST_MODE = False


# ==========================================
# CURRENT IST TIME HELPERS
# ==========================================

def ist_now():

    return datetime.now(IST)


def total_minutes_ist():

    now = ist_now()

    return now.hour * 60 + now.minute


# ==========================================
# WEEKEND WINDOW CHECK
# Both Saturday (weekday=5) and Sunday (weekday=6)
# are active between 06:30 and 18:30 IST.
# ==========================================

def within_weekend_window():

    if TEST_MODE:
        return True

    now    = ist_now()
    day    = now.weekday()
    mins   = total_minutes_ist()

    start  = WINDOW_START_H * 60 + WINDOW_START_M   # 390
    end    = WINDOW_END_H   * 60 + WINDOW_END_M      # 1110

    if day in (5, 6):
        return start <= mins <= end

    return False


# ==========================================
# MAIL RECEIVED-TIME FILTER
# Outlook shows received time in the row text.
# We parse it and skip mails received outside
# the 06:30–18:30 IST window so old/off-hours
# mails don't get processed.
# ==========================================

# Patterns Outlook uses in row previews:
#   "12:07 PM"   → today's mail
#   "10:21 AM"
#   "Yesterday"  → skip (not today)
#   "Mon"/"Tue"  → skip (older)
#   "20/05/26"   → skip (dated, not today)

TIME_RE = re.compile(
    r'\b(\d{1,2}):(\d{2})\s*(AM|PM)\b',
    re.IGNORECASE
)


def mail_received_in_window(row_text):

    if TEST_MODE:
        return True

    # If row shows "Yesterday", a weekday name, or a
    # date like "20/05/26" it's not from today → skip.

    lower = row_text.lower()

    stale_signals = [
        "yesterday",
        "monday", "tuesday", "wednesday",
        "thursday", "friday",
        # dated rows contain "/" between digits
    ]

    for signal in stale_signals:
        if signal in lower:
            return False

    # Dated rows like "20/05/26" — has digit/digit
    if re.search(r'\b\d{2}/\d{2}/\d{2}\b', row_text):
        return False

    match = TIME_RE.search(row_text)

    if not match:
        # No parseable time → allow through
        # (body check will validate content anyway)
        return True

    hour   = int(match.group(1))
    minute = int(match.group(2))
    ampm   = match.group(3).upper()

    # Convert to 24h

    if ampm == "PM" and hour != 12:
        hour += 12

    if ampm == "AM" and hour == 12:
        hour = 0

    mail_mins = hour * 60 + minute

    start = WINDOW_START_H * 60 + WINDOW_START_M
    end   = WINDOW_END_H   * 60 + WINDOW_END_M

    return start <= mail_mins <= end


# ==========================================
# BODY EXTRACTION
# ==========================================

def extract_body(page):

    for selector in BODY_SELECTORS:

        try:

            locator = page.locator(selector)

            if locator.count() > 0:

                text = locator.first.inner_text(
                    timeout=5000
                )

                if (
                    text
                    and len(text.strip()) > 20
                    and "no preview is available"
                    not in text.lower()
                ):

                    return text.strip()

        except:
            pass

    return ""


# ==========================================
# SUBJECT EXTRACTION FROM ROW
# ==========================================

def extract_subject_from_row(row_text):

    if not row_text:
        return ""

    lines = [
        line.strip()
        for line in row_text.split("\n")
        if line.strip()
    ]

    junk_patterns = [
        "do not reply",
        "sent:",
        "utc",
        "@",
        "<",
        ">",
        "from:",
        "to:",
        "cc:"
    ]

    for line in lines:

        lower = line.lower()

        has_keyword = any([
            "dispatch"  in lower,
            "handover"  in lower,
            "[ho]"      in lower,
            "ho:"       in lower,
            "[ho-mw]"   in lower,
            "ho-mw"     in lower,
            "ho created" in lower,
            "-ho"       in lower
        ])

        # Also accept Re: lines that carry a case# and P1/P2
        # (handover reply chains that lost the [HO] tag)

        has_re_case = (
            lower.startswith("re:")
            and bool(re.search(r'\b\d{4}-\d{3,5}-\d{4,}\b', line))
            and bool(re.search(r'\bp[12]\b', lower))
        )

        is_junk = any(junk in lower for junk in junk_patterns)

        if (has_keyword or has_re_case) and not is_junk:

            return re.sub(r'\s+', ' ', line).strip()

    return ""


# ==========================================
# BUILD STABLE UNIQUE ID
# ==========================================

def build_unique_id(details, subject):

    case_ref = details["Case#"]

    if not case_ref:

        case_ref = re.sub(
            r'\s+', ' ',
            subject.strip().lower()
        )

    delivery = details["Case Delivery Type"].lower().strip()

    return f"{case_ref}_{delivery}"


# ==========================================
# PROCESS VISIBLE MAILS
# ==========================================

def process_visible_mails(page):

    mails = page.locator("div[role='option']")

    count = mails.count()

    print(f"\nVisible mails: {count}")

    if count == 0:

        print("Inbox appears empty -> reloading")

        page.reload()

        page.wait_for_timeout(8000)

        return

    for i in range(count):

        try:

            mail = page.locator(
                "div[role='option']"
            ).nth(i)

            try:
                mail.scroll_into_view_if_needed()
            except:
                pass

            try:
                row_text = mail.inner_text(timeout=5000)
            except:
                continue

            if not row_text:
                continue

            # ================================
            # TIME WINDOW FILTER
            # Skip mails received outside
            # 06:30–18:30 IST on Sat/Sun.
            # ================================

            if not mail_received_in_window(row_text):

                print(
                    "Mail outside time window -> skipped"
                )

                continue

            lower_row = row_text.lower()

            # Primary: known dispatch/handover keyword

            has_type_keyword = any(
                keyword in lower_row
                for keyword in KEYWORDS
            )

            # Secondary: Re: with case# + P1/P2
            # (handover reply chains, no HO tag in subject)

            has_case_in_row = bool(
                re.search(
                    r'\b\d{4}-\d{3,5}-\d{4,}\b',
                    lower_row
                )
            )

            has_p1_p2_in_row = bool(
                re.search(r'\bp[12]\b', lower_row)
            )

            is_reply_row = lower_row.lstrip().startswith("re:")

            is_candidate = (
                has_type_keyword
                or (
                    is_reply_row
                    and has_case_in_row
                    and has_p1_p2_in_row
                )
            )

            if not is_candidate:
                continue

            subject = extract_subject_from_row(row_text)

            if not subject:

                print("\n===== SUBJECT EXTRACTION FAILED =====")
                print(row_text)
                print("====================================")

                continue

            print("\n====================================")
            print("Detected Subject:")
            print(subject)
            print("====================================")

            if should_skip_mail(subject):

                print("Skipped reply mail")

                continue

            # ================================
            # INITIAL EXTRACTION (subject only)
            # Used for P3/P4 and dedup checks.
            # ================================

            details = extract_case_details(

                subject,

                timestamp=ist_now().strftime("%d-%b-%y")

            )

            if not details:

                print("Skipped P3/P4 mail")

                mark_processed(
                    "skip_" + subject.lower().strip()
                )

                continue

            unique_id = build_unique_id(details, subject)

            if already_processed(unique_id):

                print("Already processed")

                continue

            print("\nOpening mail...")

            try:

                mail.click(timeout=5000)

            except:

                print("Retrying click...")

                page.wait_for_timeout(2000)

                mail.click(timeout=5000)

            page.wait_for_timeout(2500)

            body = extract_body(page)

            # ================================
            # FINAL EXTRACTION (with body)
            # ================================

            details = extract_case_details(

                subject,

                body,

                timestamp=ist_now().strftime("%d-%b-%y")

            )

            if not details:

                print("Skipped P3/P4 mail")

                mark_processed(
                    "skip_" + subject.lower().strip()
                )

                continue

            # Handovers may arrive without a case# yet —
            # only enforce case# for Dispatch mails.

            if (
                not details["Case#"]
                and details["Case Delivery Type"] != "Handover"
            ):

                print("Case number missing -> skipped")

                mark_processed(
                    "skip_" + subject.lower().strip()
                )

                continue

            if not details["Case Delivery Type"]:

                print("Delivery type missing -> skipped")

                mark_processed(
                    "skip_" + subject.lower().strip()
                )

                continue

            print("\nExtracted Details:\n")

            for key, value in details.items():
                print(f"{key}: {value}")

            append_to_excel(details)

            final_unique_id = build_unique_id(details, subject)

            mark_processed(final_unique_id)

            logger.info(f"Processed: {subject}")

            print("\nSaved Successfully")

        except Exception as e:

            logger.error(str(e))

            print(f"\nError processing mail {i+1}")

            print(e)


# ==========================================
# MAIN LOOP
# ==========================================

def run_mail_reader():

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False)

        context = browser.new_context(
            storage_state=str(AUTH_FILE)
        )

        page = context.new_page()

        print("\nOpening Outlook...")

        page.goto(
            "https://outlook.office.com/mail",
            wait_until="domcontentloaded"
        )

        page.wait_for_timeout(10000)

        print("\nOutlook loaded successfully")

        loop_count = 0

        while True:

            try:

                # ============================
                # AUTO REFRESH EVERY 10 MIN
                # ============================

                if loop_count % 20 == 0:

                    print("\nRefreshing Outlook...")

                    page.reload()

                    page.wait_for_timeout(8000)

                if within_weekend_window():

                    now_ist = ist_now()

                    print(
                        f"\nChecking mailbox... "
                        f"[IST {now_ist.strftime('%H:%M')} "
                        f"{'Saturday' if now_ist.weekday()==5 else 'Sunday'}]"
                    )

                    process_visible_mails(page)

                    page.mouse.wheel(0, 4000)

                    page.wait_for_timeout(3000)

                else:

                    now_ist = ist_now()

                    print(
                        f"Outside window "
                        f"[IST {now_ist.strftime('%a %H:%M')}]"
                    )

                loop_count += 1

                time.sleep(5)

            except Exception as e:

                logger.error(str(e))

                print("\nMain Loop Error:")

                print(e)

                try:
                    page.reload()
                    page.wait_for_timeout(10000)
                except:
                    pass

                time.sleep(5)