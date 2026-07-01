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
    mark_processed
)

from scripts.logger import logger

from scripts.email_report import send_report          # ← NEW

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
# TEST_MODE = False → Sat+Sun, delivery-type windows below
# ──────────────────────────────────────────
TEST_MODE  = False
IST        = timezone(timedelta(hours=5, minutes=30))
HANDOVER_START  = (0,  0)    # midnight — all day
HANDOVER_END    = (23, 59)   # end of day
DISPATCH_START  = (6,  30)
DISPATCH_END    = (18, 30)
SLEEP_SECS = 60           # 1 min between scans
RUNNING    = False

# Folder config — set at start time via run_mail_reader()
_dispatch_folder = "inbox"
_handover_folder = "inbox"
_em_name         = ""     # EM selected on Email tab, saved per row in Excel

TIME_RE = re.compile(
    r'\b(\d{1,2}):(\d{2})\s*(AM|PM)\b', re.IGNORECASE
)

# Day names used by Outlook as section-header labels.
# Any row whose text starts with one of these is from a previous day.
_ALL_DAYS = [
    "monday", "tuesday", "wednesday", "thursday", "friday",
    "saturday", "sunday",
    "mon", "tue", "wed", "thu", "fri", "sat", "sun",
]

def _stale_day_labels():
    """
    Return the set of day-name labels that represent *previous* days.
    On any given day we want to stop when we hit a row labelled with
    another day-of-week (Outlook uses these as section dividers:
    "Yesterday", "Monday", "Tuesday", …).
    We always include "yesterday" and every day name except today's,
    so the list works correctly on weekdays as well as weekends.
    """
    today_name = ist_now().strftime("%A").lower()          # e.g. "monday"
    today_short = ist_now().strftime("%a").lower()         # e.g. "mon"
    stale = {"yesterday", "this week", "last week", "older"}
    for d in _ALL_DAYS:
        if d != today_name and d != today_short:
            stale.add(d)
    return stale

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
    return _mins(*HANDOVER_END)

def dispatch_start_mins():
    return _mins(*DISPATCH_START)

def dispatch_end_mins():
    return _mins(*DISPATCH_END)

def scan_start_mins():
    return min(handover_start_mins(), dispatch_start_mins())

def scan_end_mins():
    return max(handover_end_mins(), dispatch_end_mins())

def within_window():
    if TEST_MODE:
        return True
    n = ist_now()
    # Run every day — same behaviour on weekdays for testing
    return True


def report_due_now(now):
    if TEST_MODE:
        return False
    return (
        now.weekday() in (5, 6)
        and _report_sent_today != now.date()
        and _mins(now.hour, now.minute) >= dispatch_end_mins()
    )


def sleep_while_running(seconds):
    _wake_event.clear()
    for _ in range(seconds):
        if not RUNNING:
            return
        if _wake_event.wait(timeout=1):
            _wake_event.clear()
            return


# ──────────────────────────────────────────
# ROW TIMESTAMP CHECK
# ──────────────────────────────────────────

def delivery_type_hints(text):
    low = text.lower()
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
    stale_labels = _stale_day_labels()

    # ── Outlook mail-row timestamp logic ──────────────────────────────
    # Outlook shows each mail row with a timestamp on the right:
    #   TODAY's mails  → "6:44 AM"  (a real HH:MM AM/PM time)
    #   OLDER mails    → "Yesterday", "Monday", "6/27" etc.
    #
    # IMPORTANT: stale-label checks run FIRST, before checking for a
    # real time, because a subject can contain digit patterns that
    # accidentally match TIME_RE (e.g. "P1 case 6:30 update").
    # "Yesterday" in any position means the row is not from today.

    # 1. "Yesterday" anywhere in the row → always stale.
    if "yesterday" in low:
        return "stale"

    # 2. Any other stale day/period label present in the row text.
    #    Outlook places these as the timestamp on older mail rows, so
    #    they appear somewhere within the full inner_text of the row.
    for sig in stale_labels:
        # Use word-boundary style match: sig surrounded by
        # start/end of string, whitespace, or newline.
        if re.search(r'(?:^|\s)' + re.escape(sig) + r'(?:\s|$)', low):
            return "stale"

    # 3. Short date stamps like "6/27" or "06/27/25" shown instead of a
    #    time on mails older than today (no time → not today).
    if re.search(r'\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b', text) and not TIME_RE.search(text):
        return "stale"

    # 4. Real time present → today's mail, parse it.
    m = TIME_RE.search(text)
    if not m:
        # No time, no stale label — safest to treat as stale.
        return "stale"
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
    # stale or no recognisable timestamp → stop the scan immediately
    if row_mins is None or row_mins == "stale":
        return "stop"
    hints = delivery_type_hints(text)
    if "handover" in hints:
        return True if handover_start_mins() <= row_mins <= handover_end_mins() else False
    if "dispatch" in hints:
        return True if dispatch_start_mins() <= row_mins <= dispatch_end_mins() else False
    if not hints:
        return (
            handover_start_mins() <= row_mins <= handover_end_mins()
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
        # Also match FW:/FWD: forwards (not just RE:) that carry a case# + P1/P2
        has_re = (
            bool(re.match(r'^(re|fw|fwd)\s*:', low))
            and bool(re.search(r'\b\d{4}-\d{3,5}-\d{4,}\b', line))
            and bool(re.search(r'\bp[12]\b', low))
        )
        # Remove priority-change arrows before junk check
        # so "P1 > P2" doesn't trigger the ">" junk filter.
        # Also strip "| " pipe separators used in case subjects like
        # "[HANDOVER] | P1-->P2 || TIGO..." so "|" doesn't hit junk filter.
        line_for_junk = re.sub(
            r'\bP[1-5]\s*[-=]*>\s*P[1-5]\b',
            'PRICHANGE',
            line,
            flags=re.IGNORECASE
        )
        # Remove pipes that are separators between case fields (not email headers)
        # A pipe is a separator if it's surrounded by spaces or at start/end,
        # rather than being part of an email address or header like "From: | To:"
        if has_kw or has_re:
            line_for_junk = re.sub(r'\s*\|\|?\s*', ' ', line_for_junk)
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
    if case and typ:  return f"{case}_{typ}"
    if case:          return f"{case}_unknown"
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
        except:
            pass
    return ""


# ──────────────────────────────────────────
# IS PINNED
# ──────────────────────────────────────────

def is_pinned_row(el, text):
    """
    Returns True if this DOM row is a section header (Pinned, Today, Yesterday…)
    or an actual pinned mail — both should be skipped without stopping the scan.
    """
    try:
        aria = el.get_attribute("aria-label") or ""
        attr = el.get_attribute("data-is-pinned") or ""
        text_lower = text.strip().lower()
        # Pinned mail attribute (OWA)
        if attr == "true" or "pinned" in aria.lower():
            return True
        # Section-header rows: their entire text is just a label like
        # "Pinned", "Today", "Yesterday", "This Week", etc.
        # They have no case number and no time — short single-line text.
        lines = [l.strip() for l in text_lower.splitlines() if l.strip()]
        if len(lines) <= 2 and not re.search(r'\d{4}-\d{3,5}-\d{4,}', text):
            section_labels = (
                {"pinned", "today", "this month", "earlier this month", "focused", "other"}
                | _stale_day_labels()
            )
            if any(text_lower.startswith(label) for label in section_labels):
                return True
        return False
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
      "stop"    – row is before all report windows (caller stops scan)
    """

    # 1. Time check
    ts = row_in_window(row_text)
    if ts == "stop":
        return "stop"
    if ts is False:
        return "skipped"
    # ts is True → row has a recognised today-time and is within window

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
        r'\bP[1-5]\s*[-=]*>\s*(P([1-5]))\b',
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
    #    Use case#+delivery_type as key when available
    #    (most specific). For new mails the delivery
    #    type isn't known yet so we check case# alone
    #    first — if same case+type already saved, skip.
    case_match = re.search(
        r'\b(\d{4}-\d{3,5}-\d{4,})\b', subject
    )
    case_num = case_match.group(1) if case_match else ""
    needs_technology_update = technology_missing_for_case(
        case_num
    )

    # Check if this exact case was already fully saved
    # (any delivery type) — avoid re-saving same case twice
    if case_num and not needs_technology_update:
        for suffix in ("handover", "dispatch p1", "dispatch p2", "unknown", ""):
            key = f"{case_num}_{suffix}" if suffix else case_num
            if already_processed(key):
                print(f"       → already processed ({suffix or 'any'})")
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
    details["EM"] = _em_name   # inject selected EM for this session
    for k, v in details.items():
        print(f"          {k}: {v}")

    append_to_excel(details)
    fid = make_id(details, subject)
    mark_processed(sid)
    mark_processed(fid)
    # Also mark bare case# so any re-forward of the same case is caught
    if case_num:
        mark_processed(case_num)
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
    processed_ids      = set()
    consecutive_no_new = 0
    max_no_new         = 20   # increased: Outlook re-renders same rows often
    last_count         = 0

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
            except:
                continue

            if not text or not text.strip():
                continue

            if is_pinned_row(el, text):
                continue

            fp_case = re.search(r'\b(\d{4}-\d{3,5}-\d{4,})\b', text)
            fp_time = TIME_RE.search(text)
            # Include a delivery-type hint so the same case# appearing as
            # both a Dispatch and a Handover email are treated as distinct rows
            low_text = text.lower()
            if "handover" in low_text or "[ho]" in low_text or "ho:" in low_text:
                fp_type = "ho"
            elif "dispatch" in low_text or "case created" in low_text:
                fp_type = "dp"
            else:
                fp_type = "xx"
            if fp_case and fp_time:
                fp = f"{fp_case.group(1)}_{fp_type}_{fp_time.group(0).replace(' ','').lower()}"
            elif fp_case:
                fp = f"{fp_case.group(1)}_{fp_type}"
            else:
                fp = _norm(text[:120])

            if fp in processed_ids:
                continue

            processed_ids.add(fp)
            found_new = True


            result = process_mail(page, el, text, i)


            # After clicking a mail the DOM is rebuilt — rows locator is stale.
            # Break inner loop and re-fetch from the outer while True.
            if result in ("saved", "already"):
                if result == "saved":
                    saved += 1
                else:
                    already += 1
                # Scroll back to top so we re-scan from newest mail
                _scroll_to_top(page)
                page.wait_for_timeout(1500)
                consecutive_no_new = 0
                break   # re-enter while True to get fresh rows locator

            if result == "stop":
                print(f"\n  Stale mail reached — folder scan complete")
                print(f"  Folder totals: saved={saved} already={already} skipped={skipped}")
                return True
            elif result == "skipped":
                skipped += 1

        if not found_new:
            consecutive_no_new += 1
            if consecutive_no_new >= max_no_new:
                print("  Reached inbox bottom")
                break
        else:
            consecutive_no_new = 0

        # Scroll to load more
        try:
            panel = page.locator("div[role='list']").first
            if panel.count() > 0:
                panel.evaluate("el => el.scrollBy(0, 1200)")
                page.wait_for_timeout(600)
                panel.evaluate("el => el.scrollBy(0, 1200)")
                page.wait_for_timeout(600)
            else:
                raise Exception("no panel")
        except:
            try:
                page.locator("div[role='option']").last.scroll_into_view_if_needed()
                page.wait_for_timeout(600)
            except:
                page.keyboard.press("End")
                page.wait_for_timeout(400)

        page.wait_for_timeout(1200)

        new_count = page.locator("div[role='option']").count()
        if new_count == last_count:
            try:
                page.keyboard.press("PageDown")
                page.wait_for_timeout(800)
            except:
                pass
        last_count = new_count

    print(f"\n  Folder scan done: saved={saved} already={already} skipped={skipped}")
    return True


# ──────────────────────────────────────────
# MAIN LOOP
# ──────────────────────────────────────────
# Every cycle:
#   1. Reload Outlook (fresh inbox state)
#   2. Scroll through ENTIRE inbox row by row
#      processing every mail as we go,
#      stopping when we hit a pre-03:00 mail
#      or the inbox bottom
#   3. Sleep 5 minutes
#   4. Repeat until 18:30 IST
#   5. After 18:30 — send Excel report by email
# ──────────────────────────────────────────

# Track if report has been sent this session
# so it doesn't send multiple times
_report_sent_today = None
CHROME_PATH = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

# ── Manual send-report flags (set by app.py via /api/send_report) ──
_send_requested = False   # app.py sets True → loop sends report then clears
_paused         = False   # True while report is being sent
_sleeping       = False   # True while between scans (safe window for send)
_wake_event     = __import__("threading").Event()


def _scroll_to_top(page):
    """
    Wait for the mail list to be rendered, then scroll to top
    so the scan always starts from the newest (topmost) mail.
    """
    LIST_SELS = ["div[role='list']", "div[aria-label='Message list']"]

    # 1. Wait up to 10 s for the mail list panel to appear after a
    #    reload or folder navigation — it may not exist yet immediately.
    list_loc = None
    for _ in range(20):                       # 20 × 500 ms = 10 s max
        for sel in LIST_SELS:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible(timeout=500):
                    list_loc = loc
                    break
            except Exception:
                pass
        if list_loc:
            break
        page.wait_for_timeout(500)

    # 2. Also wait for at least one mail row to be visible —
    #    the list container can appear before rows are rendered.
    for _ in range(10):                       # 10 × 800 ms = 8 s max
        try:
            if page.locator("div[role='option']").count() > 0:
                break
        except Exception:
            pass
        page.wait_for_timeout(800)

    # 3. Scroll the list panel to the very top.
    if list_loc:
        try:
            list_loc.evaluate("el => el.scrollTo(0, 0)")
            page.wait_for_timeout(400)
            list_loc.evaluate("el => el.scrollTo(0, 0)")  # double-tap
            page.wait_for_timeout(400)
        except Exception:
            pass

    # 4. Keyboard Home as a belt-and-braces fallback.
    try:
        page.keyboard.press("Home")
        page.wait_for_timeout(400)
    except Exception:
        pass

    page.wait_for_timeout(600)


def _navigate_to_folder(page, folder_name):
    """
    Navigate to a named folder in Outlook.
    - 'inbox' → clicks the Inbox item in the sidebar
    - anything else → expands/scans the folder tree for a text match
    """
    target = (folder_name or "inbox").strip()

    print(f"  → Navigating to folder: '{target}'")

    def _click_if_visible(selectors, timeout=2000):
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible(timeout=timeout):
                    loc.click(timeout=timeout)
                    page.wait_for_timeout(800)
                    return True
            except Exception:
                continue
        return False

    def _try_click_folder():
        selectors = [
            f"[aria-label='{target}']",
            f"[title='{target}']",
            f"div[role='treeitem']:has-text('{target}')",
            f"a[role='treeitem']:has-text('{target}')",
            f"button:has-text('{target}')",
            f"span:text-is('{target}')",
            f"div[role='option']:has-text('{target}')",
        ]

        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible(timeout=1500):
                    loc.scroll_into_view_if_needed(timeout=2000)
                    loc.click(timeout=4000)
                    page.wait_for_timeout(3500)
                    print(f"  → Folder '{target}' found and clicked ✓")
                    return True
            except Exception:
                continue

        try:
            loc = page.get_by_text(target, exact=True).first
            if loc.count() > 0 and loc.is_visible(timeout=1500):
                loc.scroll_into_view_if_needed(timeout=2000)
                loc.click(timeout=4000)
                page.wait_for_timeout(3500)
                print(f"  → Folder '{target}' found by exact text ✓")
                return True
        except Exception:
            pass

        return False

    try:
        if target.lower() == "inbox":
            try:
                page.goto("https://outlook.office.com/mail/inbox", wait_until="domcontentloaded")
                page.wait_for_timeout(5000)
                print("  → Inbox opened ✓")
                return
            except Exception:
                pass

        # Expand/show the folder tree if Outlook has hidden custom folders.
        _click_if_visible([
            "button[aria-label='Expand folders']",
            "button[title='Expand folders']",
            "[aria-label='Show folder list']",
            "button[aria-label='Show navigation pane']",
            "button[title='Show navigation pane']",
        ])
        _click_if_visible([
            "button:has-text('More')",
            "span:has-text('More')",
            "button:has-text('Folders')",
            "div[role='button']:has-text('Folders')",
            "[aria-label='More folders']",
            "[title='More folders']",
        ])

        if _try_click_folder():
            return

        # Custom folders can be below the visible part of Outlook's folder tree.
        tree = None
        for tree_sel in [
            "div[role='tree']",
            "nav[aria-label*='folder']",
            "div[aria-label*='folder']",
            "div[aria-label*='Folder']",
        ]:
            try:
                loc = page.locator(tree_sel).first
                if loc.count() > 0 and loc.is_visible(timeout=1000):
                    tree = loc
                    break
            except Exception:
                continue

        if tree:
            for _ in range(12):
                if _try_click_folder():
                    return
                try:
                    tree.evaluate("el => el.scrollBy(0, 350)")
                    page.wait_for_timeout(500)
                except Exception:
                    break

        # Last resort: iterate loaded tree items and match visible text.
        try:
            items = page.locator("div[role='treeitem'], a[role='treeitem']")
            count = items.count()
            for i in range(count):
                try:
                    item = items.nth(i)
                    txt  = item.inner_text(timeout=1000).strip()
                    if txt.lower() == target.lower():
                        item.click(timeout=4000)
                        page.wait_for_timeout(3000)
                        print(f"  → Folder '{target}' matched by text scan ✓")
                        return
                except Exception:
                    continue
        except Exception:
            pass

        print(f"  → Warning: folder '{target}' not found in sidebar — scanning current view")

    except Exception as e:
        print(f"  → Navigation error: {e} — scanning current view")


def run_mail_reader(dispatch_folder="inbox", handover_folder="inbox", em_name=""):
    global RUNNING, _report_sent_today, _dispatch_folder, _handover_folder, _live_page, _em_name
    global _send_requested, _paused, _sleeping

    RUNNING          = True
    _send_requested  = False
    _paused          = False
    _sleeping        = False
    _wake_event.clear()
    _dispatch_folder = (dispatch_folder or "inbox").strip()
    _handover_folder = (handover_folder or "inbox").strip()
    _em_name         = (em_name or "").strip()

    print(f"\nFolder config:")
    print(f"  Dispatch → '{_dispatch_folder}'")
    print(f"  Handover → '{_handover_folder}'")
    print(f"  EM       → '{_em_name}'")

    browser = None
    try:
        with sync_playwright() as p:

            browser = p.chromium.launch(
                executable_path=CHROME_PATH,
                headless=False
            )
            context = browser.new_context(storage_state=str(AUTH_FILE))
            page    = context.new_page()
            _live_page = page

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
                        n     = ist_now()
                        today = n.date()
                        if (
                            not TEST_MODE
                            and _report_sent_today != today
                            and n.weekday() in (5, 6)
                            and _mins(n.hour, n.minute) > scan_end_mins()
                        ):
                            print(
                                f"\n{'='*55}\n"
                                f"WINDOW CLOSED [{n.strftime('%H:%M')} IST]"
                                f" — Sending report email...\n"
                                f"{'='*55}"
                            )
                            _paused = True
                            try:
                                send_report(page=page, signer_name=_em_name)
                            except Exception as _e:
                                print(f"  Report send error: {_e}")
                            finally:
                                _paused = False
                            _report_sent_today = today
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
                        else n.strftime("%A")
                    )

                    print(
                        f"\n{'='*55}\n"
                        f"SCAN START [{n.strftime('%H:%M')} IST — {day}]\n"
                        f"{'='*55}"
                    )

                    print("Reloading Outlook...")
                    page.reload()
                    page.wait_for_timeout(8000)

                    if not RUNNING:
                        break

                    same_folder = _dispatch_folder.lower() == _handover_folder.lower()

                    if same_folder:
                        print(f"\n[Folder] Both in '{_dispatch_folder}' — single scan")
                        _navigate_to_folder(page, _dispatch_folder)
                        _scroll_to_top(page)
                        if RUNNING:
                            run_one_scan(page)
                    else:
                        print(f"\n[Folder] Dispatch → '{_dispatch_folder}'")
                        _navigate_to_folder(page, _dispatch_folder)
                        _scroll_to_top(page)
                        if RUNNING:
                            run_one_scan(page)

                        if RUNNING:
                            print(f"\n[Folder] Handover → '{_handover_folder}'")
                            _navigate_to_folder(page, _handover_folder)
                            _scroll_to_top(page)
                            run_one_scan(page)

                    if not RUNNING:
                        break

                    # ── Emit cycle-complete after BOTH folders are done ──
                    # app.py watches for "scan done:" to fire the UI notification.
                    # Emitting it here (not inside run_one_scan) ensures it only
                    # fires once per full cycle, after dispatch + handover are both
                    # processed — not mid-cycle after just the first folder.
                    print(f"Scan done: both folders processed.")

                    n = ist_now()
                    if report_due_now(n):
                        print(
                            f"\n{'='*55}\n"
                            f"REPORT TIME [{n.strftime('%H:%M')} IST]"
                            f" — Sending report email...\n"
                            f"{'='*55}"
                        )
                        _paused = True
                        try:
                            send_report(page=page, signer_name=_em_name)
                        except Exception as _e:
                            print(f"  Report send error: {_e}")
                        finally:
                            _paused = False
                        _report_sent_today = n.date()
                        print("  Automation resumed.\n")

                    n    = ist_now()
                    wake = n + timedelta(seconds=SLEEP_SECS)
                    print(
                        f"\nSleeping {SLEEP_SECS // 60} min "
                        f"[now {n.strftime('%H:%M')} — "
                        f"next ~{wake.strftime('%H:%M')} IST]"
                    )
                    _sleeping = True
                    sleep_while_running(SLEEP_SECS)
                    _sleeping = False

                    # Handle send request that arrived during sleep
                    if _send_requested:
                        _send_requested = False
                        _paused = True
                        print("\n" + "="*55 + "\nSENDING REPORT (requested during sleep)\n" + "="*55)
                        try:
                            send_report(page=page, signer_name=_em_name)
                        except Exception as _e:
                            print(f"  Report send error: {_e}")
                        finally:
                            _paused = False
                        print("  Automation resumed.\n")

                except Exception as e:
                    logger.error(str(e))
                    print(f"\nMain loop error: {e}")
                    if not RUNNING:
                        break
                    try:
                        page.reload()
                        page.wait_for_timeout(10000)
                    except:
                        pass
                    sleep_while_running(30)

    except Exception as e:
        print(f"\nBrowser error: {e}")
    finally:
        # Always close browser and reset state on exit
        _live_page = None
        RUNNING    = False
        try:
            if browser:
                browser.close()
                print("\nBrowser closed ✓")
        except Exception as e:
            print(f"\nBrowser close error: {e}")
        print("\nAutomation stopped.")
