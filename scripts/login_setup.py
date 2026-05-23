from playwright.sync_api import sync_playwright
from pathlib import Path

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


with sync_playwright() as p:

    browser = p.chromium.launch(
        executable_path=CHROME_PATH,
        headless=False
    )

    context = browser.new_context()

    page = context.new_page()

    page.goto(
        "https://outlook.office.com"
    )

    input(
        "Login manually and press ENTER..."
    )

    AUTH_FILE.parent.mkdir(exist_ok=True)

    context.storage_state(
        path=str(AUTH_FILE)
    )

    print("Session saved")

    browser.close()
