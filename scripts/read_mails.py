from playwright.sync_api import sync_playwright
RUNNING = False

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
# TEST_MODE = False → Sat+Sun 06:30–18:30 IST
# ──────────────────────────────────────────
TEST_MODE  = True
IST        = timezone(timedelta(hours=5, minutes=30))
WIN_START  = (6,  30)
WIN_END    = (18, 30)
SLEEP_SECS = 300          # 5 min between scans

TIME_RE = re.compile(
    r'\b(\d{1,2}):(\d{2})\s*(AM|PM)\b', re.IGNORECASE
)
STALE = [
    "yesterday",
    "monday","tuesday","wednesday","thursday","friday"
]

# ──────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────

def ist_now():
    return datetime.now(IST)

def _mins(h, m):
    return h * 60 + m

def win_start_mins():
    return _mins(*WIN_START)

def win_end_mins():
    return _mins(*WIN_END)

def within_window():
    if TEST_MODE:
        return True
    n = ist_now()
    if n.weekday() not in (5, 6):
        return False
    return win_start_mins() <= _mins(n.hour, n.minute) <= win_end_mins()


# ──────────────────────────────────────────
# ROW TIMESTAMP CHECK
#   True  → inside 06:30–18:30
#   False → before window or stale (old date)
#   None  → no readable timestamp
# ──────────────────────────────────────────

def row_in_window(text):
    if TEST_MODE:
        return True

    low = text.lower()

    for sig in STALE:
        if sig in low:
            return False

    if re.search(r'\b\d{2}/\d{2}/\d{2}\b', text):
        return False

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

    return win_start_mins() <= _mins(h, mn) <= win_end_mins()


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
        is_junk = any(j in low for j in JUNK)
        if (has_kw or has_re) and not is_junk:
            return re.sub(r'\s+', ' ', line).strip()
    return ""


# ──────────────────────────────────────────
# DEDUP ID
# ──────────────────────────────────────────

def _norm(s):
    s = re.sub(r'^(re|fw|fwd)\s*:\s*', '', s.strip().lower())
    return re.sub(r'\s+', ' ', s).strip()

def make_id(details, subject):
    case = details.get("Case#", "").strip()
    typ  = details.get("Case Delivery Type", "").lower().strip()
    if case and typ:  return f"{case}_{typ}"
    if case:          return f"{case}_unknown"
    return f"subj_{_norm(subject)}"


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
        except:
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
    except:
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
      "stop"    – row is before 06:30 (caller stops scan)
    """

    # 1. Time check
    ts = row_in_window(row_text)
    if ts is False:
        return "stop"
    if ts is None:
        return "skipped"          # no timestamp → skip quietly

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

    # 5. Dedup
    sid = f"subj_{_norm(subject)}"
    if already_processed(sid):
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
    except:
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
# Outlook virtual-scrolls: only ~6 rows exist
# in the DOM at any time. Old rows are removed
# as new ones appear. We CANNOT pre-load all
# rows. Instead we work with what's visible,
# process each mail on the spot, then scroll
# down to reveal the next batch — like reading
# a newspaper page by page.
#
# Algorithm:
#   while True:
#     for each visible row (top → bottom):
#       if pre-window → stop, return False
#       process it
#     scroll down one row-height
#     if no new rows appeared → inbox bottom reached
#     loop
# ──────────────────────────────────────────

def run_one_scan(page):

    print("\nStarting full inbox scan...")

    saved = skipped = already = 0
    processed_ids = set()     # track within this scan run
    consecutive_no_new = 0    # detect true bottom
    last_seen_count = 0

    while True:

        # --- grab currently visible rows ---
        rows = page.locator("div[role='option']")
        count = rows.count()

        if count == 0:
            print("  No rows visible — inbox may be empty")
            break

        found_new = False

        for i in range(count):

            try:
                el   = rows.nth(i)
                text = el.inner_text(timeout=3000)
            except:
                continue

            if not text or not text.strip():
                continue

            # Skip pinned
            if is_pinned_row(el, text):
                continue

            # Build a short fingerprint to know if we've
            # seen this exact row in this scan already
            fp = _norm(text[:120])
            if fp in processed_ids:
                continue

            processed_ids.add(fp)
            found_new = True

            result = process_mail(page, el, text, i)

            if result == "stop":
                print(
                    "\n  Pre-06:30 mail reached — "
                    "scan complete up to window boundary"
                )
                print(
                    f"  Scan totals: "
                    f"saved={saved} already={already} skipped={skipped}"
                )
                return True    # normal completion, sleep after

            elif result == "saved":
                saved += 1
            elif result == "already":
                already += 1
            else:
                skipped += 1

        # --- scroll down to reveal next batch ---
        if not found_new:
            consecutive_no_new += 1
            if consecutive_no_new >= 3:
                # Scrolled 3 times with zero new rows
                # → we're at the real bottom of the inbox
                print("  Reached inbox bottom")
                break
        else:
            consecutive_no_new = 0

        # Scroll the list panel down
        try:
            panel = page.locator("div[role='list']").first
            panel.evaluate("el => el.scrollBy(0, 300)")
        except:
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
# Every cycle:
#   1. Reload Outlook (fresh inbox state)
#   2. Scroll through ENTIRE inbox row by row
#      processing every mail as we go,
#      stopping when we hit a pre-06:30 mail
#      or the inbox bottom
#   3. Sleep 5 minutes
#   4. Repeat until 18:30 IST
# ──────────────────────────────────────────

def run_mail_reader():

    with sync_playwright() as p:

        CHROME_PATH = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
               
    
        browser = p.chromium.launch(
        executable_path=CHROME_PATH,
        headless=False
    )
        
        browser = p.chromium.launch(headless=False)
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

                if not within_window():
                    n = ist_now()
                    print(
                        f"Outside window "
                        f"[{n.strftime('%a %H:%M')} IST] — sleeping 60s"
                    )
                    time.sleep(60)
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

                # Fresh reload so new mails are visible
                print("Reloading Outlook...")
                page.reload()
                page.wait_for_timeout(8000)

                # Scroll to top of inbox before scanning
                try:
                    panel = page.locator("div[role='list']").first
                    panel.evaluate("el => el.scrollTo(0, 0)")
                    page.wait_for_timeout(1000)
                except:
                    pass

                # Run full scan
                run_one_scan(page)

                # Sleep 5 minutes
                n    = ist_now()
                wake = n + timedelta(seconds=SLEEP_SECS)
                print(
                    f"\nSleeping {SLEEP_SECS // 60} min "
                    f"[now {n.strftime('%H:%M')} — "
                    f"next ~{wake.strftime('%H:%M')} IST]"
                )
                time.sleep(SLEEP_SECS)

            except Exception as e:
                logger.error(str(e))
                print(f"\nMain loop error: {e}")
                try:
                    page.reload()
                    page.wait_for_timeout(10000)
                except:
                    pass
                time.sleep(30)