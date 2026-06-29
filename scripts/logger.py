import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

LOG_FILE = LOG_DIR / "automation.log"

# Rotating handler: caps automation.log at ~2 MB, keeps 3 old copies
# (automation.log.1, .2, .3) before deleting the oldest. Prevents
# unbounded growth even if something unexpectedly logs a lot.
_handler = RotatingFileHandler(
    filename=str(LOG_FILE),
    maxBytes=2 * 1024 * 1024,  # 2 MB
    backupCount=3,
    encoding="utf-8",
)
_handler.setFormatter(
    logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
)

logging.basicConfig(
    level=logging.INFO,
    handlers=[_handler],
)

# Stop Flask/werkzeug's per-request access logs (e.g. every "GET /api/status"
# poll from the frontend) from flooding this file. Only warnings/errors from
# werkzeug will now be logged; routine 200 OK request logs are suppressed.
logging.getLogger("werkzeug").setLevel(logging.WARNING)

logger = logging.getLogger(__name__)