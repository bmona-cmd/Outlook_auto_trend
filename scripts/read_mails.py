from playwright.sync_api import sync_playwright

from parser import (
    extract_case_details,
    should_skip_mail
)

from excel_writer import append_to_excel

from tracker import (
    already_processed,
    mark_processed
)

from logger import logger

from pathlib import Path

from datetime import datetime

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

    "ho:"
]


BODY_SELECTORS = [

    "div[role='document']",

    "div[aria-label='Message body']",

    "div[data-app-section='MailReadCompose']"
]


# ==========================================
# TEST MODE
# ==========================================

TEST_MODE = True


# ==========================================
# WEEKEND CHECK
# ==========================================

def within_weekend_window():

    # ======================================
    # TEST MODE
    # ======================================

    if TEST_MODE:
        return True

    # ======================================
    # ORIGINAL WEEKEND LOGIC
    # ======================================

    now = datetime.now()

    weekday = now.weekday()

    hour = now.hour
    minute = now.minute

    total_minutes = (
        hour * 60 + minute
    )

    saturday_start = (
        6 * 60 + 30
    )

    sunday_end = (
        18 * 60 + 30
    )

    if weekday == 5:

        return (
            total_minutes
            >=
            saturday_start
        )

    if weekday == 6:

        return (
            total_minutes
            <=
            sunday_end
        )

    return False


# ==========================================
# BODY EXTRACTION
# ==========================================

def extract_body(page):

    for selector in BODY_SELECTORS:

        try:

            locator = page.locator(
                selector
            )

            if locator.count() > 0:

                text = locator.first.inner_text(
                    timeout=5000
                )

                if (

                    text

                    and

                    len(text.strip()) > 20

                    and

                    "no preview is available"
                    not in text.lower()
                ):

                    return text.strip()

        except:
            pass

    return ""


# ==========================================
# SUBJECT EXTRACTION
# ==========================================

def extract_subject_from_row(row_text):

    if not row_text:
        return ""

    lines = [

        line.strip()

        for line in row_text.split("\n")

        if line.strip()
    ]

    candidate_lines = []

    for line in lines:

        lower = line.lower()

        if (
            "dispatch" in lower
            or "handover" in lower
            or "[ho]" in lower
            or "ho:" in lower
            or re.search(
                r'\b\d{4}-\d{3,5}-\d{4,}\b',
                line
            )
        ):

            candidate_lines.append(line)

    if not candidate_lines:
        return ""

    # PREFER LINES WITH CASE NUMBER

    for line in candidate_lines:

        if re.search(
            r'\b\d{4}-\d{3,5}-\d{4,}\b',
            line
        ):

            return re.sub(
                r'\s+',
                ' ',
                line
            ).strip()

    # OTHERWISE LONGEST LINE

    candidate_lines.sort(
        key=lambda x: len(x),
        reverse=True
    )

    return re.sub(
        r'\s+',
        ' ',
        candidate_lines[0]
    ).strip()


# ==========================================
# PROCESS MAILS
# ==========================================

def process_visible_mails(page):

    mails = page.locator(
        "div[role='option']"
    )

    count = mails.count()

    print(f"\nVisible mails: {count}")

    if count == 0:

        print(
            "Inbox appears empty -> reloading"
        )

        page.reload()

        page.wait_for_timeout(8000)

        return

    for i in range(min(count, 30)):

        try:

            # RE-QUERY EACH LOOP
            # PREVENT STALE ELEMENTS

            mail = page.locator(
                "div[role='option']"
            ).nth(i)

            try:

                mail.scroll_into_view_if_needed()

            except:
                pass

            try:

                row_text = mail.inner_text(
                    timeout=5000
                )

            except:

                continue

            if not row_text:
                continue

            lower_row = row_text.lower()

            if not any(

                keyword in lower_row

                for keyword in KEYWORDS
            ):

                continue

            subject = extract_subject_from_row(
                row_text
            )

            if not subject:
                continue

            print("\n====================================")
            print("Detected Subject:")
            print(subject)
            print("====================================")

            if should_skip_mail(subject):

                print(
                    "Skipped reply mail"
                )

                continue

            details = (
                extract_case_details(
                    subject,
                    timestamp=datetime.now()
                    .strftime("%d-%b-%y")
                )
            )

            unique_id = (

                details["Case#"]
                + "_"
                + details["Case Delivery Type"]

            ).lower().strip()

            if already_processed(
                unique_id
            ):

                print(
                    "Already processed"
                )

                continue

            print(
                "\nOpening mail..."
            )

            try:

                mail.click(
                    timeout=5000
                )

            except:

                print(
                    "Retrying click..."
                )

                page.wait_for_timeout(2000)

                mail.click(
                    timeout=5000
                )

            page.wait_for_timeout(2500)

            body = extract_body(page)

            details = (
                extract_case_details(
                    subject,
                    body,
                    timestamp=datetime.now()
                    .strftime("%d-%b-%y")
                )
            )

            if not details["Case#"]:

                print(
                    "Case number missing -> skipped"
                )

                continue

            if not details[
                "Case Delivery Type"
            ]:

                print(
                    "Delivery type missing -> skipped"
                )

                continue

            print(
                "\nExtracted Details:\n"
            )

            for key, value in details.items():

                print(
                    f"{key}: {value}"
                )

            append_to_excel(details)

            mark_processed(unique_id)

            logger.info(
                f"Processed: {subject}"
            )

            print(
                "\nSaved Successfully"
            )

        except Exception as e:

            logger.error(str(e))

            print(
                f"\nError processing mail {i+1}"
            )

            print(e)


# ==========================================
# MAIN LOOP
# ==========================================

def run_mail_reader():

    with sync_playwright() as p:

        browser = p.chromium.launch(

            executable_path=CHROME_PATH,

            headless=False
        )

        context = browser.new_context(

            storage_state=str(AUTH_FILE)
        )

        page = context.new_page()

        print(
            "\nOpening Outlook..."
        )

        page.goto(

            "https://outlook.office.com/mail",

            wait_until="domcontentloaded"
        )

        page.wait_for_timeout(10000)

        print(
            "\nOutlook loaded successfully"
        )

        loop_count = 0

        while True:

            try:

                # ==================================
                # AUTO REFRESH EVERY 10 MINUTES
                # ==================================

                if loop_count % 20 == 0:

                    print(
                        "\nRefreshing Outlook..."
                    )

                    page.reload()

                    page.wait_for_timeout(8000)

                if within_weekend_window():

                    print(
                        "\nChecking mailbox..."
                    )

                    process_visible_mails(
                        page
                    )

                    page.mouse.wheel(
                        0,
                        4000
                    )

                    page.wait_for_timeout(3000)

                else:

                    print(
                        "Outside weekend window"
                    )

                loop_count += 1

                time.sleep(30)

            except Exception as e:

                logger.error(str(e))

                print(
                    "\nMain Loop Error:"
                )

                print(e)

                try:

                    page.reload()

                    page.wait_for_timeout(
                        10000
                    )

                except:
                    pass

                time.sleep(30)