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
# WEEKEND CHECK
# ==========================================

def within_weekend_window():

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
# BETTER SUBJECT EXTRACTION
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

    candidate_lines.sort(
        key=lambda x: len(x),
        reverse=True
    )

    subject = candidate_lines[0]

    subject = subject.replace("\r", " ")
    subject = subject.replace("\n", " ")

    subject = re.sub(r'\s+', ' ', subject)

    return subject.strip()


# ==========================================
# PROCESS MAILS
# ==========================================

def process_visible_mails(page):

    mails = page.locator(
        "div[role='option']"
    )

    count = mails.count()

    print(f"Visible mails: {count}")

    for i in range(min(count, 30)):

        try:

            mail = mails.nth(i)

            try:

                row_text = mail.inner_text(
                    timeout=3000
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
                    "Skipped reply"
                )

                continue

            unique_id = (
                subject.lower().strip()
            )

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

            mail.click()

            page.wait_for_timeout(2500)

            body = extract_body(page)

            timestamp = (
                datetime.now()
                .strftime("%d-%b-%y")
            )

            details = (
                extract_case_details(
                    subject,
                    body,
                    timestamp=timestamp
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

        page.goto(

            "https://outlook.office.com/mail",

            wait_until="domcontentloaded"
        )

        page.wait_for_timeout(8000)

        print("Outlook loaded")

        while True:

            try:

                if within_weekend_window():

                    process_visible_mails(page)

                    page.mouse.wheel(
                        0,
                        4000
                    )

                    page.wait_for_timeout(3000)

                else:

                    print(
                        "Outside weekend window"
                    )

                time.sleep(30)

            except Exception as e:

                logger.error(str(e))

                print(e)

                time.sleep(30)