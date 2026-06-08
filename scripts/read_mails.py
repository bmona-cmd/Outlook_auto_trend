from playwright.sync_api import sync_playwright

from scripts.parser import (
    extract_case_details,
    should_skip_mail
)

from scripts.excel_writer import (
    append_to_excel,
    technology_missing_for_case
)

from scripts.tracker import (
    already_processed,
    mark_processed,
    reset_if_new_day
)

from scripts.logger import logger

from scripts.email_report import send_report

from pathlib import Path

from datetime import datetime, timezone, timedelta

import time
import re


BASE_DIR  = Path(__file__).resolve().parent.parent
AUTH_FILE = BASE_DIR / "auth" / "auth.json"

KEYWORDS = [
    "dispatch", "handover", "[ho]", "ho:",
    "ho created", "-ho", "[ho-mw]", "ho-mw",
    "| ho |", "|ho|", "case created"
]

BODY_SELECTORS = [
    "div[role='document']",
    "div[aria-label='Message body']",
    "div[data-app-section='MailReadCompose']"
]

JUNK = [
    "do not reply", "sent:", "utc",
    "@", "<", ">", "from:", "to:", "cc:"
]

# ──────────────────────────────────────────
# CONFIG
# TEST_MODE = True  → any day / any time
# TEST_MODE = False → Fri-Sun weekend tracking:
#   - Saturday 03:00 → Saturday 24:00
#   - Sunday   03:00 → Sunday   24:00
# ──────────────────────────────────────────
TEST_MODE       = False
IST             = timezone(timedelta(hours=5, minutes=30))
HANDOVER_START  = (3,  0)
HANDOVER_END    = (24, 0)   # midnight (end of day)
DISPATCH_START  = (6,  30)
DISPATCH_END    = (18, 30)
SLEEP_SECS      = 300       # 5 min between scans
RUNNING         = False

TIME_RE = re.compile(
    r'\b(\d{1,2}):(\d{2})\s*(AM|PM)\b', re.IGNORECASE
)
STALE = [
    "yesterday",
    "monday", "tuesday", "wednesday", "thursday", "friday"
]


def get_stale_words():
    """
    Dynamic stale-word list based on current day.
    On Sunday  -> also stale: saturday, sat
    On Saturday -> nothing extra (saturday is today)
    """
    words = list(STALE)
    if ist_now().weekday() == 6:   # Sunday
        words += ["saturday", "sat"]
    return words


def is_row_from_today(el) -> bool:
    """
    Read the aria-label on the Outlook row element.
    Outlook sets aria-label to e.g.:
      'Received Saturday June 7, ...'   (yesterday, on Sunday)
      'Received Sunday June 8, ...'     (today, on Sunday)
    We compare the date in the label to today's IST date.
    Falls back to True (don't block) if aria-label is unavailable.
    """
    try:
        aria = el.get_attribute("aria-label") or ""
        if not aria:
            return True  # can't tell -- don't block

        import re as _re
        m = _re.search(
            r'(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|'
            r'Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|'
            r'Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2})',
            aria, _re.IGNORECASE
        )
        if not m:
            return True  # can't parse -- don't block

        month_str = m.group(1)[:3].lower()
        day_num   = int(m.group(2))

        month_map = {
            "jan": 1,  "feb": 2,  "mar": 3,  "apr": 4,
            "may": 5,  "jun": 6,  "jul": 7,  "aug": 8,
            "sep": 9,  "oct": 10, "nov": 11, "dec": 12
        }
        month_num = month_map.get(month_str)
        if not month_num:
            return True

        now = ist_now()
        return now.month == month_num and now.day == day_num

    except Exception:
        return True  # on any error -- don't block


# ──────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────

def ist_now():
    return datetime.now(IST)


def _mins(h, m):
    return h * 60 + m


def handover_start_mins():
    return _mins(*HANDOVER_START)


def handover_end_mins():
    return _mins(*HANDOVER_END)   # 1440 = midnight


def dispatch_start_mins():
    return _mins(*DISPATCH_START)


def dispatch_end_mins():
    return _mins(*DISPATCH_END)


def scan_start_mins(now=None):
    return handover_start_mins()


def scan_end_mins():
    return handover_end_mins()


def is_handover_time(row_mins, now=None):
    return handover_start_mins() <= row_mins <= handover_end_mins()


def within_window():
    if TEST_MODE:
        return True
    n = ist_now()
    # Active only on Saturday and Sunday
    if n.weekday() not in (5, 6):
        return False
    now_mins = _mins(n.hour, n.minute)
    # Window open from 03:00 through midnight
    return now_mins >= handover_start_mins()


def sleep_while_running(seconds):
    for _ in range(seconds):
        if not RUNNING:
            return
        time.sleep(1)


# ──────────────────────────────────────────
# ROW TIMESTAMP CHECK
#   True      → inside matching delivery-type window
#   False     → outside matching window, but keep scanning
#   "stop"    → before all report windows or stale (old date)
#   None      → no readable timestamp
# ──────────────────────────────────────────

def delivery_type_hints(text):
    low   = text.lower()
    hints = set()

    if "dispatch" in low:
        hints.add("dispatch")

    if any(sig in low for sig in (
        "handover", "[ho]", "ho:", "ho created", "-ho",
        "[ho-mw]", "ho-mw", "| ho |", "|ho|", "case created"
    )):
        hints.add("handover")

    return hints


def _row_time_mins(text):
    low = text.lower()

    for sig in get_stale_words():
        if sig in low:
            return "stale"

    if re.search(r'\b\d{2}/\d{2}/\d{2}\b', text):
        return "stale"

    m = TIME_RE.search(text)
    if not m:
        return None

    h  = int(m.group(1))
    mn = int(m.group(2))
    ap = m.group(3).upper()

    if ap == "PM" and h != 12:
        h += 12
    if ap == "AM" and h == 12:
        h = 0

    return _mins(h, mn)


def row_in_window(text):
    if TEST_MODE:
        return True

    row_mins = _row_time_mins(text)

    if row_mins is None:
        return None

    if row_mins == "stale" or row_mins < handover_start_mins():
        return "stop"

    hints = delivery_type_hints(text)

    if "handover" in hints and is_handover_time(row_mins):
        return True

    if "dispatch" in hints and dispatch_start_mins() <= row_mins <= dispatch_end_mins():
        return True

    # Replies or ambiguous rows may not carry a clear delivery type until opened.
    # Keep them eligible when they fall into either valid reporting window.
    if not hints:
        return (
            is_handover_time(row_mins)
            or dispatch_start_mins() <= row_mins <= dispatch_end_mins()
        )

    return False


# ──────────────────────────────────────────
# SUBJECT EXTRACTION
# ──────────────────────────────────────────

def get_subject(row_text):
    for line in [l.strip() for l in row_text.split("\n") if l.strip()]:
        low = line.lower()
        has_kw = any([
            "dispatch"     in low, "handover"  in low,
            "[ho]"         in low, "ho:"       in low,
            "[ho-mw]"      in low, "ho-mw"     in low,
            "ho created"   in low, "-ho"        in low,
            "| ho |"       in low, "|ho|"       in low,
            "case created" in low,
        ])
        has_re = (
            low.startswith("re:")
            and bool(re.search(r'\b\d{4}-\d{3,5}-\d{4,}\b', line))
            and bool(re.search(r'\bp[12]\b', low))
        )
        # Remove priority-change arrows before junk check
        # so "P1 > P2" doesn't trigger the ">" junk filter
        line_for_junk = re.sub(
            r'\bP[1-5]\s*[-=]?>\s*P[1-5]\b',
            'PRICHANGE',
            line,
            flags=re.IGNORECASE
        )
        is_junk = any(j in line_for_junk.lower() for j in JUNK)
        if (has_kw or has_re) and not is_junk:
            return re.sub(r'\s+', ' ', line).strip()
    return ""


# ──────────────────────────────────────────
# DEDUP ID
# ──────────────────────────────────────────

def _norm(s):
    s = re.sub(r'^(re|fw|fwd)\s*:\s*', '', s.strip().lower())
    return re.sub(r'\s+', ' ', s).strip()


def _row_timestamp(row_text):
    """Extract the time string from the row for use in ID."""
    m = TIME_RE.search(row_text)
    if m:
        return m.group(0).strip().lower().replace(" ", "")
    return "notime"


def make_id(details, subject):
    case = details.get("Case#", "").strip()
    typ  = details.get("Case Delivery Type", "").lower().strip()
    if case and typ:
        return f"{case}_{typ}"
    if case:
        return f"{case}_unknown"
    return f"subj_{_norm(subject)}"


def make_scan_sid(subject, row_text):
    """
    Dedup key used WITHIN a scan run.
    Combines normalised subject + received time so
    a new mail with the same subject arriving later
    is not blocked by a previously processed one.
    Case# is the most stable anchor — use it when
    available so re-forwards of same case are deduped.
    """
    case_match = re.search(
        r'\b(\d{4}-\d{3,5}-\d{4,})\b', subject
    )
    if case_match:
        return f"case_{case_match.group(1)}"
    return f"subj_{_norm(subject)}_{_row_timestamp(row_text)}"


# ──────────────────────────────────────────
# BODY EXTRACTION
# ──────────────────────────────────────────

def extract_body(page):
    for sel in BODY_SELECTORS:
        try:
            loc = page.locator(sel)
            if loc.count() > 0:
                text = loc.first.inner_text(timeout=5000)
                if text and len(text.strip()) > 20 \
                   and "no preview is available" not in text.lower():
                    return text.strip()
        except Exception:
            pass
    return ""


# ──────────────────────────────────────────
# IS PINNED
# ──────────────────────────────────────────

def is_pinned_row(el, text):
    try:
        aria = el.get_attribute("aria-label") or ""
        attr = el.get_attribute("data-is-pinned") or ""
        return (
            attr == "true"
            or "pinned" in aria.lower()
            or text.strip().lower() == "pinned"
        )
    except Exception:
        return False


# ──────────────────────────────────────────
# PROCESS ONE MAIL
# ──────────────────────────────────────────

def process_mail(page, el, row_text, idx):
    """
    Returns:
      "saved"   – extracted and written to Excel
      "already" – already in processed_mails.json
      "skipped" – not relevant / P3-P5 / no case
      "stop"    – row is before all report windows (caller stops scan)
    """

    # 1. Time check
    ts = row_in_window(row_text)
    if ts == "stop":
        return "stop"
    if ts is False:
        return "skipped"
    if ts is None:
        return "skipped"

    low = row_text.lower()

    # 2. Candidate keyword check
    has_kw = any(kw in low for kw in KEYWORDS)
    has_re = (
        low.lstrip().startswith("re:")
        and bool(re.search(r'\b\d{4}-\d{3,5}-\d{4,}\b', low))
        and bool(re.search(r'\bp[12]\b', low))
    )
    if not has_kw and not has_re:
        return "skipped"

    # 3. P1/P2 only
    # Check for priority-change pattern first (P1>P2, P1->P2)
    # and use the RIGHT side as the current priority.
    change_pm = re.search(
        r'\bP[1-5]\s*[-=]?>\s*(P([1-5]))\b',
        row_text,
        re.IGNORECASE
    )
    if change_pm:
        level = int(change_pm.group(2))
        if level not in (1, 2):
            print(f"    [{idx}] P{level} (after change) — skipped")
            return "skipped"
    else:
        pm = re.search(r'\bP([1-5])\b', row_text, re.IGNORECASE)
        if pm and int(pm.group(1)) not in (1, 2):
            print(f"    [{idx}] P{pm.group(1)} — skipped")
            return "skipped"

    # 4. Subject
    subject = get_subject(row_text)
    if not subject:
        print(f"    [{idx}] subject extraction failed")
        return "skipped"

    print(f"\n  [{idx}] {subject}")

    if should_skip_mail(subject):
        print("       → ack/reply — skipped")
        return "skipped"

    # 5. Dedup against processed_mails.json
    case_match = re.search(
        r'\b(\d{4}-\d{3,5}-\d{4,})\b', subject
    )
    case_num = case_match.group(1) if case_match else ""
    needs_technology_update = technology_missing_for_case(case_num)

    # Check if this exact case was already fully saved
    if (
        case_num
        and already_processed(f"{case_num}_handover")
        and not needs_technology_update
    ):
        print("       → already processed (handover)")
        return "already"
    if (
        case_num
        and already_processed(f"{case_num}_dispatch p1")
        and not needs_technology_update
    ):
        print("       → already processed (dispatch p1)")
        return "already"
    if (
        case_num
        and already_processed(f"{case_num}_dispatch p2")
        and not needs_technology_update
    ):
        print("       → already processed (dispatch p2)")
        return "already"

    # For no-case# mails use subject+time key
    sid = make_scan_sid(subject, row_text)
    if already_processed(sid) and not needs_technology_update:
        print("       → already processed")
        return "already"

    # 6. Quick P3/P4 check without body
    details = extract_case_details(
        subject, timestamp=ist_now().strftime("%d-%b-%y")
    )
    if not details:
        print("       → P3/P4 — skipped")
        mark_processed(sid)
        return "skipped"

    # 7. Click to open, get body
    print("       → opening...")
    try:
        el.scroll_into_view_if_needed()
        el.click(timeout=6000)
    except Exception:
        try:
            page.wait_for_timeout(2000)
            el.click(timeout=6000)
        except Exception as e:
            print(f"       → click failed: {e}")
            return "skipped"

    page.wait_for_timeout(3000)
    body = extract_body(page)

    # 8. Full extraction with body
    details = extract_case_details(
        subject, body, timestamp=ist_now().strftime("%d-%b-%y")
    )
    if not details:
        print("       → P3/P4 (body) — skipped")
        mark_processed(sid)
        return "skipped"

    if not details["Case#"] and details["Case Delivery Type"] != "Handover":
        print("       → no Case# — skipped")
        mark_processed(sid)
        return "skipped"

    if not details["Case Delivery Type"]:
        print("       → no delivery type — skipped")
        mark_processed(sid)
        return "skipped"

    # 9. Save
    print("       → saving:")
    for k, v in details.items():
        print(f"          {k}: {v}")

    append_to_excel(details)
    fid = make_id(details, subject)
    mark_processed(sid)
    mark_processed(fid)
    logger.info(f"Processed: {subject}")
    print("       → SAVED ✓")
    return "saved"


# ──────────────────────────────────────────
# FULL SCAN
# ──────────────────────────────────────────

def run_one_scan(page):

    print("\nStarting full inbox scan...")

    saved = skipped = already = 0
    processed_ids     = set()
    consecutive_no_new = 0

    while True:

        rows  = page.locator("div[role='option']")
        count = rows.count()

        if count == 0:
            print("  No rows visible — inbox may be empty")
            break

        found_new = False

        for i in range(count):

            try:
                el   = rows.nth(i)
                text = el.inner_text(timeout=3000)
            except Exception:
                continue

            if not text or not text.strip():
                continue

            if is_pinned_row(el, text):
                continue

            fp_case = re.search(r'\b(\d{4}-\d{3,5}-\d{4,})\b', text)
            fp_time = TIME_RE.search(text)
            if fp_case and fp_time:
                fp = f"{fp_case.group(1)}_{fp_time.group(0).replace(' ', '').lower()}"
            elif fp_case:
                fp = fp_case.group(1)
            else:
                fp = _norm(text[:120])

            if fp in processed_ids:
                continue

            processed_ids.add(fp)
            found_new = True

            if not TEST_MODE and not is_row_from_today(el):
                print(f"    [{i}] not from today (aria-label date mismatch) — skipped")
                skipped += 1
                continue

            result = process_mail(page, el, text, i)

            if result == "stop":
                print(
                    "\n  Pre-03:00 mail reached — "
                    "scan complete up to window boundary"
                )
                print(
                    f"  Scan totals: "
                    f"saved={saved} already={already} skipped={skipped}"
                )
                return True

            elif result == "saved":
                saved += 1
            elif result == "already":
                already += 1
            else:
                skipped += 1

        if not found_new:
            consecutive_no_new += 1
            if consecutive_no_new >= 3:
                print("  Reached inbox bottom")
                break
        else:
            consecutive_no_new = 0

        try:
            panel = page.locator("div[role='list']").first
            panel.evaluate("el => el.scrollBy(0, 300)")
        except Exception:
            page.keyboard.press("ArrowDown")

        page.wait_for_timeout(800)

    print(
        f"\n  Scan done: "
        f"saved={saved} already={already} skipped={skipped}"
    )
    return True


# ──────────────────────────────────────────
# MAIN LOOP
# ──────────────────────────────────────────

_report_sent_today = None

CHROME_PATH = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"


def run_mail_reader():
    global RUNNING
    global _report_sent_today

    RUNNING = True

    with sync_playwright() as p:

        browser = p.chromium.launch(
            executable_path=CHROME_PATH,
            headless=False
        )
        context = browser.new_context(storage_state=str(AUTH_FILE))
        page    = context.new_page()

        print("\nOpening Outlook...")
        page.goto(
            "https://outlook.office.com/mail",
            wait_until="domcontentloaded"
        )
        page.wait_for_timeout(10000)
        print("Outlook loaded ✓")

        while RUNNING:

            try:

                # ── Reset tracker at the start of each new day ──
                reset_if_new_day()

                if not within_window():

                    n         = ist_now()
                    today     = n.date()
                    yesterday = today - timedelta(days=1)

                    # Send report once after window closes
                    if (
                        not TEST_MODE
                        and _report_sent_today != today
                        and n.weekday() in (5, 6)
                        and _mins(n.hour, n.minute) >= scan_end_mins()
                    ):
                        print(
                            f"\n{'='*55}\n"
                            f"WINDOW CLOSED [{n.strftime('%H:%M')} IST]"
                            f" — Sending report email...\n"
                            f"{'='*55}"
                        )
                        send_report(target_date=today)
                        _report_sent_today = today

                    elif (
                        not TEST_MODE
                        and _report_sent_today != yesterday
                        and n.weekday() == 0
                        and yesterday.weekday() == 6
                        and _mins(n.hour, n.minute) < handover_start_mins()
                    ):
                        print(
                            f"\n{'='*55}\n"
                            f"WINDOW CLOSED [Sunday missed] [{n.strftime('%H:%M')} IST]"
                            f" — Sending Sunday report email...\n"
                            f"{'='*55}"
                        )
                        send_report(target_date=yesterday)
                        _report_sent_today = yesterday

                    print(
                        f"Outside window "
                        f"[{n.strftime('%a %H:%M')} IST] — sleeping 60s"
                    )
                    sleep_while_running(60)
                    continue

                n   = ist_now()
                day = (
                    "Saturday" if n.weekday() == 5
                    else "Sunday" if n.weekday() == 6
                    else "TEST"
                )

                print(
                    f"\n{'='*55}\n"
                    f"SCAN START [{n.strftime('%H:%M')} IST — {day}]\n"
                    f"{'='*55}"
                )

                print("Reloading Outlook...")
                page.reload()
                page.wait_for_timeout(8000)

                # Scroll to top of inbox before scanning
                scrolled = False
                for sel in [
                    "div[role='list']",
                    "div[aria-label='Message list']",
                    "div.customScrollBar",
                    "div[data-testid='MailList']",
                ]:
                    try:
                        loc = page.locator(sel).first
                        if loc.count() > 0:
                            loc.evaluate("el => el.scrollTo(0, 0)")
                            page.wait_for_timeout(600)
                            scrolled = True
                            break
                    except Exception:
                        continue

                try:
                    first_row = page.locator("div[role='option']").first
                    if first_row.count() > 0:
                        first_row.click(timeout=3000)
                        page.wait_for_timeout(300)
                except Exception:
                    pass

                try:
                    page.keyboard.press("Home")
                    page.wait_for_timeout(600)
                except Exception:
                    pass

                if not scrolled:
                    print("  Warning: could not confirm scroll-to-top")
                page.wait_for_timeout(800)

                run_one_scan(page)

                n    = ist_now()
                wake = n + timedelta(seconds=SLEEP_SECS)
                print(
                    f"\nSleeping {SLEEP_SECS // 60} min "
                    f"[now {n.strftime('%H:%M')} — "
                    f"next ~{wake.strftime('%H:%M')} IST]"
                )
                sleep_while_running(SLEEP_SECS)

            except Exception as e:
                logger.error(str(e))
                print(f"\nMain loop error: {e}")
                try:
                    page.reload()
                    page.wait_for_timeout(10000)
                except Exception:
                    pass
                sleep_while_running(30)

    RUNNING = False