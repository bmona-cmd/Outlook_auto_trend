import re

from scripts.vertical_lookup import (
    get_vertical,
    get_original_company_name,
    first_word_vertical_map,
    normalize_customer_name,
    normalized_vertical_map,
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
    "vsrx": "Security",
    "ssg": "Security",

    "qfx": "Switching",
    "ex": "Switching",

    "mist": "Wireless",

    "128t": "SDWAN",

    "software": "Software"
}


TEAM_TECH_MAP = {

    "routing": "Routing",
    "switching": "Switching",
    "security": "Security",
    "software": "Software",
    "wireless": "Wireless",
    "sdwan": "SDWAN"
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
        re.search(r'\bp[12]\b', lower_subject)
    ) or bool(
        # P1>P2 or P1->P2 style — current priority is P1 or P2
        re.search(r'\bp[1-5]\s*[-=]?>\s*p[12]\b', lower_subject)
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

    # Handle priority change patterns first:
    #   P1 > P2, P1 -> P2, P1>P2, P1->P2
    # The RIGHT side is the current priority.

    change = re.search(
        r'\bP[1-5]\s*[-=]?>\s*(P[1-5])\b',
        subject,
        re.IGNORECASE
    )

    if change:
        return change.group(1).upper()

    # Standard single priority
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


def extract_technology(subject, body="", cc_text=""):

    combined = (
        subject
        + " "
        + body
        + " "
        + cc_text
    ).lower()

    for keyword, tech in DEVICE_TECH_MAP.items():

        pattern = _get_device_pattern(keyword)

        if re.search(
            pattern,
            combined
        ):

            return tech

    recipient_text = (
        body
        + " "
        + cc_text
    ).lower()

    tech_scores = {}

    for keyword, tech in TEAM_TECH_MAP.items():

        patterns = [

            rf'\b[a-z0-9._%+\-]*{keyword}[a-z0-9._%+\-]*@',

            rf'\barc[-\w]*{keyword}[-\w]*\b',

            rf'\b{keyword}\s*team\b',

            rf'\b{keyword}\s*@'
        ]

        score = 0

        for pattern in patterns:

            score += len(
                re.findall(
                    pattern,
                    recipient_text,
                    re.IGNORECASE
                )
            )

        if score:

            tech_scores[tech] = score

    if tech_scores:

        return max(
            tech_scores,
            key=tech_scores.get
        )

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


def _format_customer_name(customer):

    if not customer:
        return ""

    known_name = get_original_company_name(
        customer
    )

    if known_name != customer:
        return known_name

    words = []

    for word in str(customer).strip().split():

        if word.isupper() or any(char.isdigit() for char in word):
            words.append(word)
        else:
            words.append(word.capitalize())

    return " ".join(words)


def _first_word_has_vertical(customer):

    normalized_customer = normalize_customer_name(
        customer
    )

    if not normalized_customer:
        return False

    first_word = normalized_customer.split()[0]

    return first_word in first_word_vertical_map


def _trim_customer_words(words, start_index):

    customer_words = []

    for word in words[start_index:]:

        normalized_word = normalize_customer_name(
            word
        )

        if not normalized_word:
            continue

        if (
            normalized_word in ["case", "sr", "pr", "p1", "p2", "p3", "p4", "p5"]
            or
            re.match(r'^\d{4}$', normalized_word)
            or
            re.match(r'^\d{4}\s+\d{3,5}\s+\d{4,}$', normalized_word)
        ):
            break

        if normalized_word in DEVICE_TECH_MAP:
            break

        customer_words.append(
            word
        )

        if len(customer_words) >= 4:
            break

    return " ".join(
        customer_words
    )


def _find_first_word_customer(words):

    for index, word in enumerate(words):

        normalized_word = normalize_customer_name(
            word
        )

        if normalized_word in first_word_vertical_map:

            return _trim_customer_words(
                words,
                index
            )

    return ""


def _extract_customer_candidate(text):

    text = clean_subject(text)

    if not text:
        return ""

    dispatch_match = re.search(
        r'in\s+queue\s*:\s*.*?\s+-\s+(.+)$',
        text,
        re.IGNORECASE
    )

    if dispatch_match:

        return dispatch_match.group(1).strip()

    parts = [

        part.strip()

        for part in re.split(
            r'\|\||\||:',
            text
        )

        if part.strip()
    ]

    for part in parts:

        cleaned_part = remove_known_patterns(
            part
        )

        normalized_part = normalize_customer_name(
            cleaned_part
        )

        if not normalized_part:
            continue

        if _first_word_has_vertical(
            normalized_part
        ):

            return cleaned_part

        part_candidate = _find_first_word_customer(
            cleaned_part.split()
        )

        if part_candidate:

            return part_candidate

    cleaned = remove_known_patterns(
        text
    )

    words = cleaned.split()

    return _find_first_word_customer(
        words
    )

# ==========================================
# CUSTOMER EXTRACTION
# ==========================================

def extract_customer(text):

    if not text:
        return ""

    cleaned = remove_known_patterns(
        text
    )

    cleaned = normalize_customer_name(
        cleaned
    )

    cleaned_words = set(
        cleaned.split()
    )

    best_match = ""
    best_score = 0
    best_word_count = 0
    exact_match = ""
    exact_word_count = 0

    for excel_customer in normalized_vertical_map.keys():

        customer_clean = (
            excel_customer
            .lower()
            .strip()
        )

        customer_words = [

            word

            for word in customer_clean.split()

            if len(word) > 1
        ]

        matched_words = 0

        for word in customer_words:

            if word in cleaned_words:

                matched_words += 1

        if not customer_words:
            continue

        if (
            customer_clean in cleaned
            and
            len(customer_words) > exact_word_count
        ):

            exact_match = excel_customer

            exact_word_count = len(customer_words)

        score = (
            matched_words
            /
            len(customer_words)
        )

        # ==================================
        # BEST MATCH
        # ==================================

        if (
            score > best_score
            or
            (
                score == best_score
                and
                len(customer_words) > best_word_count
            )
        ):

            best_score = score

            best_match = excel_customer

            best_word_count = len(customer_words)

    # ======================================
    # MINIMUM CONFIDENCE
    # ======================================

    if exact_match:

        return _format_customer_name(
            exact_match
        )

    customer_candidate = _extract_customer_candidate(
        text
    )

    if _first_word_has_vertical(
        customer_candidate
    ):

        return _format_customer_name(
            customer_candidate
        )

    if best_score >= 0.85:

        return _format_customer_name(
            best_match
        )

    return ""

# ==========================================
# VERTICAL EXTRACTION
# ==========================================

def extract_vertical(customer):

    if not customer:
        return ""

    vertical = get_vertical(
        customer
    )

    if vertical:
        return vertical

    lower_customer = (
        normalize_customer_name(
            customer
        )
    )

    for company in normalized_vertical_map.keys():

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
            body,
            cc_text
        ),

        "Case Delivery Type": delivery_type,

        "Comments": ""
    }
