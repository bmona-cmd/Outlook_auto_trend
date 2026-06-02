import re

from scripts.vertical_lookup import (
    get_vertical,
    vertical_map
)

# ==========================================
# TECHNOLOGY MAP
# ==========================================

DEVICE_TECH_MAP = {

    "mx": "Routing",
    "ptx": "Routing",
    "acx": "Routing",

    "srx": "Security",
    "ssg": "Security",

    "qfx": "Switching",
    "ex": "Switching",

    "mist": "Wireless",

    "128t": "SDWAN",

    "software": "Software"
}

# ==========================================
# STOP WORDS
# ==========================================

STOP_WORDS = [

    "interface",
    "packet",
    "drops",
    "drop",
    "flap",
    "flapping",
    "error",
    "errors",
    "issue",
    "failure",
    "problem",
    "queue",
    "dispatch",
    "handover",
    "latency",
    "link",
    "down",
    "unstable",
    "alarm",
    "alerts",
    "investigation",
    "crc",
    "peer",
    "notice",
    "monitoring",
    "case",
    "reminder"
]

# ==========================================
# SKIP REPLY MAILS
# ==========================================

# These keywords indicate a casual reply with no
# case content. However they must NOT skip a mail
# that also contains a case number or P1/P2 priority
# — that check is done inside should_skip_mail().

SKIP_KEYWORDS = [

    "acknowledged",
    "thanks",
    "thank you",
    "working on",
    "noted",
    "fyi"
]

# ==========================================
# CLEAN SUBJECT
# ==========================================

def clean_subject(subject):

    if not subject:
        return ""

    subject = str(subject)

    subject = subject.replace("\n", " ")
    subject = subject.replace("\r", " ")

    subject = re.sub(
        r'\s+',
        ' ',
        subject
    )

    return subject.strip()

# ==========================================
# NORMALIZE SUBJECT
# ==========================================

def normalize_subject(subject):

    subject = clean_subject(subject)

    subject = subject.lower()

    subject = subject.replace(".", " ")

    subject = re.sub(
        r'[\|\:\,\-\/\[\]\(\)_]+',
        ' ',
        subject
    )

    subject = re.sub(
        r'\s+',
        ' ',
        subject
    )

    return subject.strip()

# ==========================================
# SKIP MAIL CHECK
# ==========================================

def should_skip_mail(subject):

    lower_subject = clean_subject(
        subject
    ).lower().strip()

    # "Re:" only skips if the mail has NO case number
    # AND NO P1/P2 priority in the subject.
    # Rationale: Re: mails like
    #   "Re: [HANDOVER] Case# 2026-... P1->P2 GOOGLE LLC"
    #   "Re: P2 | 2026-... | QFX5120 | TIGO | ..."
    # are legitimate handovers/cases forwarded to the
    # inbox and must be processed.
    # Only skip Re: mails that are plain acknowledgement
    # replies with no case reference.

    is_reply = bool(
        re.match(r'^re\s*:', lower_subject)
    )

    has_case_number = bool(
        re.search(
            r'\b\d{4}-\d{3,5}-\d{4,}\b',
            lower_subject
        )
    )

    has_priority = bool(
        re.search(
            r'\bp[12]\b',
            lower_subject
        )
    )

    if is_reply and not has_case_number and not has_priority:
        return True

    # For other skip keywords, also protect mails that
    # have a case number or P1/P2 — e.g. a subject like
    # "Thanks | 2026-0430-697636 | P1 | ..." should still
    # be processed. Only skip keyword-matched mails that
    # have no case reference at all.

    for keyword in SKIP_KEYWORDS:

        if keyword in lower_subject:

            if not has_case_number and not has_priority:
                return True

    return False

# ==========================================
# CASE NUMBER
# ==========================================

def extract_case_number(subject):

    subject = clean_subject(subject)

    patterns = [

        r'\b\d{4}-\d{4}-\d{4,}\b',

        r'\b\d{4}-\d{3,5}-\d{4,}\b'
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            subject
        )

        if match:
            return match.group()

    return ""

# ==========================================
# PRIORITY
# ==========================================

def extract_priority(subject):

    match = re.search(
        r'\bP[1-5]\b',
        subject,
        re.IGNORECASE
    )

    if match:
        return match.group().upper()

    return ""

# ==========================================
# MAIL TYPE
# ==========================================

def extract_mail_type(subject, body=""):

    # Check subject first

    lower_subject = subject.lower()

    if (
        "handover" in lower_subject
        or "[ho]" in lower_subject
        or "ho:" in lower_subject
        or "[ho-mw]" in lower_subject
        or "ho-mw" in lower_subject
        or "ho created" in lower_subject
        or "-ho" in lower_subject
        or "| ho |" in lower_subject
        or "|ho|" in lower_subject
    ):
        return "Handover"

    if (
        "dispatch" in lower_subject
        or "case created" in lower_subject
    ):
        return "Dispatch"

    # Fallback: scan body for handover/dispatch signals.
    # Needed for Re: chains where the subject lost the
    # [HO] / [HANDOVER] tag (e.g. "Re: P2 | 2026-...").

    if body:

        lower_body = body[:2000].lower()

        handover_body_signals = [
            "please accept the ho",
            "please accept the handover",
            "handover accepted",
            "accepting the handover",
            "accepting the ho",
            "accept the ho",
            "accept the case",
            "continue monitoring",
            "hi bnc cfts",
            "hi cfts",
            "bng cfts",
            "[handover]",
            "handover note",
            "ho note",
            "warm handover",
            "warm hand",
            "i will call you for warm handover",
        ]

        for signal in handover_body_signals:
            if signal in lower_body:
                return "Handover"

        dispatch_body_signals = [
            "dispatch notification",
            "please take appropriate action",
            "in queue: vonage queue",
            "dispatch notice",
            "case created",
            "has been created",
            "priority of p2",
            "priority of p1",
        ]

        for signal in dispatch_body_signals:
            if signal in lower_body:
                return "Dispatch"

    return ""

# ==========================================
# TECHNOLOGY EXTRACTION
# ==========================================

def _get_device_pattern(keyword):

    # FIX BUG 2: "ex" alone as \bex\d*\b matches
    # common words like "example", "existing",
    # "expected" which breaks customer extraction.
    # Require at least one digit after "ex" so only
    # real device names like ex2300, ex4300 match.

    if keyword == "ex":
        return r'\bex\d+\b'

    return rf'\b{re.escape(keyword)}\d*\b'


def extract_technology(subject, body=""):

    combined = (
        subject
        + " "
        + body
    ).lower()

    for keyword, tech in DEVICE_TECH_MAP.items():

        pattern = _get_device_pattern(keyword)

        if re.search(
            pattern,
            combined
        ):

            return tech

    return ""

# ==========================================
# REMOVE KNOWN METADATA
# ==========================================

def remove_known_patterns(text):

    text = normalize_subject(text)

    # REMOVE CASE NUMBERS

    text = re.sub(
        r'\b\d{4}-\d{3,5}-\d{4,}\b',
        ' ',
        text
    )

    # REMOVE PRIORITY

    text = re.sub(
        r'\bp[1-5]\b',
        ' ',
        text
    )

    # REMOVE MAIL TYPE WORDS

    text = re.sub(
        r'\bdispatch\b',
        ' ',
        text
    )

    text = re.sub(
        r'\bhandover\b',
        ' ',
        text
    )

    text = re.sub(
        r'\bho\b',
        ' ',
        text
    )

    # REMOVE DEVICE NAMES
    # FIX BUG 2: use same stricter pattern for "ex"
    # here so customer names containing "ex" aren't
    # mangled (e.g. "Apex", "Vertex", "Vexcel").

    for keyword in DEVICE_TECH_MAP.keys():

        text = re.sub(
            _get_device_pattern(keyword),
            ' ',
            text
        )

    # REMOVE EXTRA SPACES

    text = re.sub(
        r'\s+',
        ' ',
        text
    )

    return text.strip()

# ==========================================
# CUSTOMER EXTRACTION
# ==========================================

def extract_customer(text):

    if not text:
        return ""

    cleaned = remove_known_patterns(
        text
    )

    cleaned = cleaned.lower()

    best_match = ""
    best_score = 0

    for excel_customer in vertical_map.keys():

        customer_clean = (
            excel_customer
            .lower()
            .strip()
        )

        customer_words = (
            customer_clean.split()
        )

        matched_words = 0

        for word in customer_words:

            if word in cleaned:

                matched_words += 1

        if not customer_words:
            continue

        score = (
            matched_words
            /
            len(customer_words)
        )

        # ==================================
        # BEST MATCH
        # ==================================

        if score > best_score:

            best_score = score

            best_match = excel_customer

    # ======================================
    # MINIMUM CONFIDENCE
    # ======================================

    if best_score >= 0.6:

        return best_match.title()

    return ""

# ==========================================
# VERTICAL EXTRACTION
# ==========================================

def extract_vertical(customer):

    if not customer:
        return ""

    lower_customer = (
        customer.lower()
        .strip()
    )

    for company in vertical_map.keys():

        if company in lower_customer:

            return get_vertical(company)

    return ""

# ==========================================
# MAIN EXTRACTION
# ==========================================

def extract_case_details(

    subject,
    body="",
    cc_text="",
    timestamp=""
):

    subject = clean_subject(
        subject
    )

    # ======================================
    # CUSTOMER EXTRACTION
    # ======================================

    customer = extract_customer(
        subject
    )

    # FALLBACK TO BODY

    if not customer and body:

        customer = extract_customer(
            body
        )

    # ======================================
    # VERTICAL
    # ======================================

    vertical = extract_vertical(
        customer
    )

    # ======================================
    # PRIORITY
    # ======================================

    priority = extract_priority(
        subject
    )

    # ======================================
    # MAIL TYPE
    # ======================================

    mail_type = extract_mail_type(
        subject,
        body
    )

    # ======================================
    # SKIP P3/P4
    # ======================================

    if priority in ["P3", "P4"]:

        return None

    # ======================================
    # DELIVERY TYPE
    # ======================================

    delivery_type = ""

    if mail_type == "Dispatch":

        delivery_type = (
            f"Dispatch {priority}"
        ).strip()

    elif mail_type == "Handover":

        delivery_type = "Handover"

    # ======================================
    # RETURN FINAL DATA
    # ======================================

    return {

        "Date": timestamp,

        "Case#": extract_case_number(
            subject
        ),

        "Customer": customer,

        "Vertical": vertical,

        "Technology": extract_technology(
            subject,
            body
        ),

        "Case Delivery Type": delivery_type,

        "Comments": ""
    }