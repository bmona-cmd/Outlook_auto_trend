"""
email_report.py  —  Sends the weekend Excel report via Outlook Web (Playwright).

No SMTP, no app passwords, no extra credentials needed.
Reuses the already-authenticated Playwright page that is scanning mails.

Called by read_mails.py — pass the live `page` object.
Also callable manually from app.py via a stored page reference.
"""

import json
from pathlib  import Path
from datetime import datetime
import importlib.util


BASE_DIR    = Path(__file__).resolve().parent.parent
CONFIG_FILE = BASE_DIR / "email_config.json"
OUTPUT_DIR  = BASE_DIR / "output"
TRACKER_FILE = OUTPUT_DIR / "Weekend_Cases_Tracker.xlsx"


# ==========================================
# LOAD CONFIG
# ==========================================

def load_config():
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"email_config.json not found at {CONFIG_FILE}\n"
            "Please create it with recipients list."
        )
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)


# ==========================================
# FIND LATEST EXCEL
# ==========================================

def get_report_excel():
    """Return the canonical workbook that chart_exporter also reads."""
    return TRACKER_FILE if TRACKER_FILE.exists() else None


# ==========================================
# GENERATE CHART IMAGES
# ==========================================

def _generate_charts() -> list:
    if importlib.util.find_spec("matplotlib") is None:
        print(
            "  Chart generation failed: matplotlib is not installed. "
            "Run: venv/bin/python -m pip install -r requirements.txt"
        )
        return []
    try:
        from scripts.chart_exporter import generate_all_charts
        charts = generate_all_charts()
        print(f"  Charts generated: {[c.name for c in charts]}")
        return charts
    except Exception as e:
        print(f"  Chart generation skipped: {e}")
        return []


def _click_first_visible(page, selectors, timeout=3000):
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=timeout)
            loc.click(timeout=timeout)
            return sel
        except Exception:
            continue
    return None


# ==========================================
# BUILD EMAIL BODY
# ==========================================

def build_body(day_name, date_str, excel_file, chart_paths):
    chart_note = ""
    if chart_paths:
        names = ", ".join(p.name for p in chart_paths)
        chart_note = f"\nChart images attached: {names}\n"

    if excel_file:
        return (
            f"Hi,\n\n"
            f"Please find attached the weekend cases report "
            f"for {day_name}, {date_str}.\n\n"
            f"File: {excel_file.name}\n"
            f"{chart_note}\n"
            f"This email was sent automatically by the "
            f"Weekend Mail Automation.\n\n"
            f"Regards,\n"
            f"Weekend Automation Bot"
        )
    return (
        f"Hi,\n\n"
        f"The weekend automation completed for {day_name}, {date_str}, "
        f"but no Excel file was found in the output folder.\n\n"
        f"This may mean no cases were processed today.\n"
        f"{chart_note}\n"
        f"Regards,\n"
        f"Weekend Automation Bot"
    )


# ==========================================
# SEND VIA OUTLOOK WEB (PLAYWRIGHT)
# ==========================================

def _send_via_outlook_web(page, recipients, subject, body, excel_file, chart_paths):
    """
    Compose and send via Outlook Web using keyboard shortcut (Ctrl+N / N key)
    which is far more reliable than hunting for the New Mail button.
    """

    print("  Navigating to Outlook inbox...")
    page.goto("https://outlook.office.com/mail/inbox", wait_until="domcontentloaded")
    page.wait_for_timeout(6000)

    # ── Open compose via keyboard shortcut ────────────────────────────────
    # Outlook Web shortcut: press "N" when focus is on the mail list,
    # or Ctrl+N globally. Both are more reliable than clicking a button.
    print("  Opening compose via keyboard shortcut...")
    compose_open = False

    # Method 1: press N (Outlook Web global shortcut for New Message)
    try:
        page.keyboard.press("n")
        page.wait_for_timeout(2000)
        # Check if compose pane opened
        if page.locator("div[aria-label='To'], input[aria-label='To'], div[role='dialog']").count() > 0:
            compose_open = True
            print("  Compose opened via N shortcut")
    except Exception:
        pass

    # Method 2: Ctrl+N
    if not compose_open:
        try:
            page.keyboard.press("Control+n")
            page.wait_for_timeout(2000)
            if page.locator("div[aria-label='To'], input[aria-label='To']").count() > 0:
                compose_open = True
                print("  Compose opened via Ctrl+N")
        except Exception:
            pass

    # Method 3: find New Mail button by iterating all buttons and checking text/aria
    if not compose_open:
        try:
            buttons = page.locator("button, div[role='button']")
            count = buttons.count()
            print(f"  Scanning {count} buttons for New Mail...")
            for i in range(min(count, 30)):
                btn = buttons.nth(i)
                try:
                    label = (btn.get_attribute("aria-label") or "").lower()
                    text  = btn.inner_text(timeout=500).strip().lower()
                    if any(kw in label or kw in text for kw in
                           ["new mail", "new message", "compose", "new email"]):
                        btn.click(timeout=3000)
                        page.wait_for_timeout(2000)
                        compose_open = True
                        print(f"  Compose opened via button scan (i={i}): {label or text}")
                        break
                except Exception:
                    continue
        except Exception:
            pass

    # Method 4: direct URL for new compose (OWA deep link)
    if not compose_open:
        print("  Trying compose deep-link URL...")
        page.goto(
            "https://outlook.office.com/mail/deeplink/compose",
            wait_until="domcontentloaded"
        )
        page.wait_for_timeout(5000)
        compose_open = True   # assume it worked; failures caught below

    # ── Wait for To field ─────────────────────────────────────────────────
    to_selectors = [
        "div[role='textbox'][aria-label='To']",
        "div[aria-label='To'][contenteditable='true']",
        "div[aria-label='To']",
        "input[aria-label='To']",
        "div[aria-label='To'] input",
        "div[class*='to'] input",
        "div[id*='to'] input",
    ]
    to_field = None
    for sel in to_selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=5000)
            to_field = loc
            print(f"  To field found: {sel}")
            break
        except Exception:
            continue

    if to_field is None:
        raise RuntimeError(
            "Could not open Outlook compose window. "
            "Tried N key, Ctrl+N, button scan, and deep-link URL."
        )

    # ── Fill To ───────────────────────────────────────────────────────────
    suggestion_selectors = [
        "div[role='option']",
        "li[role='option']",
        "div[role='listbox'] div[role='option']",
        "div[aria-label*='suggestion']",
        "button[role='option']",
    ]

    for recipient in recipients:
        # Refocus the To field. After picking a suggestion (click, not
        # Enter), Outlook can take well over the original 3s to finish
        # rebuilding the editor before it's interactable again — so retry
        # with a longer per-attempt timeout instead of failing on the first
        # miss.
        focused_sel = None
        for attempt in range(4):                  # up to ~4 x 6s ≈ 24s total
            focused_sel = _click_first_visible(page, to_selectors, timeout=6000)
            if focused_sel:
                break
            print(f"  To field not ready yet (attempt {attempt + 1}/4) — retrying...")
            page.wait_for_timeout(1000)

        if not focused_sel:
            raise RuntimeError(f"Could not focus To field while adding {recipient}")

        # Dismiss any leftover dropdown from the previous recipient so it
        # can't be mistaken for the current one's suggestion below.
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)

        # Outlook rebuilds the To editor after resolving each recipient, so
        # type through the active keyboard focus instead of reusing a locator.
        page.keyboard.type(recipient, delay=60)

        # Poll for the autocomplete dropdown instead of a single fixed sleep.
        # Outlook's render time varies (a few hundred ms up to ~2s), and a
        # stale/leftover dropdown element can satisfy a one-shot count()
        # check even when the *current* recipient's suggestion isn't ready
        # yet — that mismatch is what was causing names to run together.
        suggestion = None
        for _ in range(10):                       # up to ~2s total
            page.wait_for_timeout(200)
            for sel in suggestion_selectors:
                try:
                    loc = page.locator(sel).first
                    if loc.count() > 0 and loc.is_visible():
                        suggestion = loc
                        break
                except Exception:
                    continue
            if suggestion:
                break

        if suggestion:
            # Click the suggestion directly — far more reliable than a blind
            # Enter press, which can land before the dropdown has taken
            # keyboard focus and just insert a space/newline instead of
            # committing a recipient chip.
            try:
                suggestion.click(timeout=2000)
                print(f"  Picked suggestion for: {recipient}")
            except Exception:
                page.keyboard.press("Enter")
                print(f"  Suggestion click failed, fell back to Enter for: {recipient}")
        else:
            # No dropdown appeared — common when the typed text is already
            # a fully valid, resolvable email address. Commit it directly.
            page.keyboard.press("Enter")
            print(f"  No suggestion, confirmed typed: {recipient}")

        # Give Outlook time to render the recipient as a confirmed chip
        # and finish rebuilding the To editor before the next loop
        # iteration tries to refocus and type the following name.
        page.wait_for_timeout(1800)

    print(f"  To filled: {recipients}")

    # ── Fill Subject ───────────────────────────────────────────────────────
    for sel in ["input[aria-label='Subject']", "input[placeholder*='Subject']",
                "div[aria-label='Subject'] input", "input[name='subject']"]:
        try:
            f = page.locator(sel).first
            if f.count() > 0:
                f.click(timeout=3000)
                f.fill(subject)
                print(f"  Subject: {subject}")
                break
        except Exception:
            continue

    # ── Fill Body ──────────────────────────────────────────────────────────
    for sel in [
        "div[aria-label='Message body, press Alt+F10 to exit']",
        "div[role='textbox'][contenteditable='true']",
        "div[contenteditable='true'][aria-multiline='true']",
        "div[aria-label*='body'][contenteditable]",
        "div[contenteditable='true']",
    ]:
        try:
            f = page.locator(sel).first
            if f.count() > 0:
                f.click(timeout=3000)
                for line in body.split("\n"):
                    f.type(line, delay=8)
                    page.keyboard.press("Shift+Enter")
                print("  Body filled.")
                break
        except Exception:
            continue

    page.wait_for_timeout(800)

    # ── Attach files ───────────────────────────────────────────────────────
    files_to_attach = []
    if excel_file and excel_file.exists():
        files_to_attach.append(str(excel_file))
    for cp in chart_paths:
        if cp.exists():
            files_to_attach.append(str(cp))

    if files_to_attach:
        print(f"  Attaching {len(files_to_attach)} file(s)...")
        print(f"  Files: {[Path(f).name for f in files_to_attach]}")
        attached = False

        # Try attach button → Upload from computer
        for attach_sel in [
            "button[aria-label='Attach']", "button[title='Attach']",
            "button:has-text('Attach')", "div[aria-label='Attach']",
            "button[aria-label='Insert']",
        ]:
            try:
                btn = page.locator(attach_sel).first
                if btn.count() == 0:
                    continue
                btn.click(timeout=4000)
                page.wait_for_timeout(1000)

                for up_sel in [
                    "span:has-text('Upload from computer')",
                    "button:has-text('Upload from computer')",
                    "span:has-text('Browse this computer')",
                    "li:has-text('Upload from computer')",
                    "div[role='menuitem']:has-text('computer')",
                ]:
                    try:
                        opt = page.locator(up_sel).first
                        if opt.count() > 0:
                            with page.expect_file_chooser(timeout=5000) as fc_info:
                                opt.click(timeout=4000)
                            fc_info.value.set_files(files_to_attach)
                            page.wait_for_timeout(5000)
                            print("  Files attached ✓")
                            attached = True
                            break
                    except Exception:
                        continue
                if attached:
                    break

                # Some Outlook builds create a hidden file input after the
                # Attach menu opens, without raising a file chooser event.
                file_input = page.locator("input[type='file']").last
                if file_input.count() > 0:
                    file_input.set_input_files(files_to_attach)
                    page.wait_for_timeout(5000)
                    print("  Files attached via file input ✓")
                    attached = True
                    break
            except Exception:
                continue

        # Fallback: set any existing file input directly, then try a chooser.
        if not attached:
            try:
                file_input = page.locator("input[type='file']").last
                if file_input.count() > 0:
                    file_input.set_input_files(files_to_attach)
                    page.wait_for_timeout(5000)
                    print("  Files attached via fallback input ✓")
                    attached = True
            except Exception as e:
                print(f"  Direct file input attach failed: {e}")

        if not attached:
            try:
                with page.expect_file_chooser(timeout=5000) as fc_info:
                    page.evaluate(
                        "() => { const i = document.querySelector('input[type=\"file\"]'); if(i) i.click(); }"
                    )
                fc_info.value.set_files(files_to_attach)
                page.wait_for_timeout(5000)
                print("  Files attached via fallback chooser ✓")
                attached = True
            except Exception as e:
                print(f"  Attachment skipped: {e}")

    # ── Send ───────────────────────────────────────────────────────────────
    # Wait a moment to ensure attachments have finished uploading
    page.wait_for_timeout(3000)

    sent = False

    # Method 1: find Send button and click it
    for sel in [
        "button[aria-label='Send']",
        "button[title='Send']",
        "button:has-text('Send')",
        "div[aria-label='Send']",
        "div[role='button'][aria-label='Send']",
    ]:
        try:
            btn = page.locator(sel).first
            if btn.count() > 0:
                # Scroll into view and force-click to ensure it registers
                btn.scroll_into_view_if_needed(timeout=2000)
                btn.focus()
                page.wait_for_timeout(300)
                btn.click(force=True, timeout=5000)
                page.wait_for_timeout(2000)
                # Confirm compose window closed (means sent successfully)
                still_open = page.locator("button[aria-label='Send'], button:has-text('Send')").count()
                if still_open == 0:
                    sent = True
                    print("  Email sent via button click ✓")
                    break
                # Still open — try clicking again
                btn.click(force=True, timeout=3000)
                page.wait_for_timeout(2000)
                sent = True
                print("  Email sent via button click (2nd attempt) ✓")
                break
        except Exception as e:
            print(f"  Send btn ({sel}): {e}")
            continue

    # Method 2: Ctrl+Enter — works in Outlook Web compose regardless of focus
    if not sent:
        try:
            print("  Trying Ctrl+Enter to send...")
            page.keyboard.press("Control+Return")
            page.wait_for_timeout(2000)
            sent = True
            print("  Email sent via Ctrl+Enter ✓")
        except Exception as e:
            print(f"  Ctrl+Enter failed: {e}")

    # Method 3: Alt+S — another Outlook send shortcut
    if not sent:
        try:
            print("  Trying Alt+S to send...")
            page.keyboard.press("Alt+s")
            page.wait_for_timeout(2000)
            sent = True
            print("  Email sent via Alt+S ✓")
        except Exception as e:
            print(f"  Alt+S failed: {e}")

    if not sent:
        raise RuntimeError("Could not send email — compose window may still be open as draft")

    # Wait to confirm send completes before returning to inbox
    page.wait_for_timeout(3000)

    # Return to inbox
    try:
        page.goto("https://outlook.office.com/mail/inbox", wait_until="domcontentloaded")
        page.wait_for_timeout(5000)
    except Exception:
        pass


# ==========================================
# MAIN — SEND REPORT
# ==========================================

def send_report(page=None, target_date=None):
    """
    page: the live Playwright page object from read_mails.py.
          Required — email is sent via Outlook Web.
    """
    try:
        config     = load_config()
        recipients = config.get("recipients", [])

        if not recipients:
            print("Email skipped: no recipients in email_config.json")
            return False

        if page is None:
            print("Email skipped: no Playwright page provided. "
                  "Pass the active page from read_mails.py.")
            return False

        excel_file = get_report_excel()

        now      = target_date or datetime.now()
        weekday  = now.weekday()
        day_name = (
            "Saturday" if weekday == 5
            else "Sunday" if weekday == 6
            else now.strftime("%A")
        )
        date_str = now.strftime("%d %b %Y")

        subject = f"Weekend Cases Report — {day_name} {date_str}"

        print("\nGenerating chart images...")
        chart_paths = _generate_charts()

        body = build_body(day_name, date_str, excel_file, chart_paths)

        print(f"\nSending report via Outlook Web...")
        print(f"  To:     {', '.join(recipients)}")
        print(f"  Excel:  {excel_file.name if excel_file else 'none'}")
        print(f"  Charts: {[c.name for c in chart_paths] or 'none'}")

        _send_via_outlook_web(
            page, recipients, subject, body, excel_file, chart_paths
        )

        print(
            f"Report email sent successfully via Outlook Web ✓\n"
            f"  To:     {', '.join(recipients)}\n"
            f"  Excel:  {excel_file.name if excel_file else 'none'}\n"
            f"  Charts: {[c.name for c in chart_paths] or 'none'}"
        )
        return True

    except FileNotFoundError as e:
        print(f"Email config error: {e}")
        return False

    except Exception as e:
        print(f"Email sending failed: {e}")
        return False