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
    "routing",
    "switching",
    "security",
    "wireless",
    "vpn",
    "bgp",
    "ospf",
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

SKIP_KEYWORDS = [

    "acknowledged",
    "thanks",
    "thank you",
    "working on",
    "noted",
    "re:",
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
    ).lower()

    for keyword in SKIP_KEYWORDS:

        if keyword in lower_subject:
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

def extract_mail_type(subject):

    lower_subject = subject.lower()

    if (
        "handover" in lower_subject
        or "[ho]" in lower_subject
        or "ho:" in lower_subject
    ):

        return "Handover"

    if "dispatch" in lower_subject:

        return "Dispatch"

    return ""


# ==========================================
# TECHNOLOGY EXTRACTION
# ==========================================

def extract_technology(subject, body=""):

    combined = (
        subject
        + " "
        + body
    ).lower()

    for keyword, tech in DEVICE_TECH_MAP.items():

        pattern = (
            rf'\b{re.escape(keyword)}\d*\b'
        )

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

    # REMOVE DELIVERY WORDS

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

    for keyword in DEVICE_TECH_MAP.keys():

        text = re.sub(
            rf'\b{keyword}\d*\b',
            ' ',
            text
        )

    text = re.sub(
        r'\s+',
        ' ',
        text
    )

    return text.strip()


# ==========================================
# CUSTOMER EXTRACTION
# ==========================================

def extract_customer(subject):

    if not subject:
        return ""

    cleaned = remove_known_patterns(
        subject
    )

    tokens = cleaned.split()

    best_customer = ""
    best_match_length = 0

    for excel_customer in vertical_map.keys():

        excel_customer_clean = (
            excel_customer
            .lower()
            .strip()
        )

        customer_words = (
            excel_customer_clean.split()
        )

        # ==================================
        # DIRECT CONTAINMENT
        # ==================================

        if excel_customer_clean in cleaned:

            start_index = cleaned.find(
                excel_customer_clean
            )

            remaining = cleaned[
                start_index:
            ]

            extracted_words = []

            for word in remaining.split():

                if word.lower() in STOP_WORDS:
                    break

                extracted_words.append(word)

            extracted_customer = (
                " ".join(extracted_words)
            ).strip()

            if (
                len(extracted_customer)
                >
                best_match_length
            ):

                best_customer = (
                    extracted_customer
                )

                best_match_length = (
                    len(extracted_customer)
                )

        # ==================================
        # TOKEN OVERLAP
        # ==================================

        else:

            common_words = (

                set(tokens)
                &
                set(customer_words)
            )

            overlap_ratio = 0

            if customer_words:

                overlap_ratio = (

                    len(common_words)
                    /
                    len(customer_words)
                )

            if overlap_ratio >= 0.6:

                extracted_words = []

                started = False

                for token in tokens:

                    if (
                        token in common_words
                        or started
                    ):

                        started = True

                        if token.lower() in STOP_WORDS:
                            break

                        extracted_words.append(
                            token
                        )

                extracted_customer = (
                    " ".join(extracted_words)
                ).strip()

                if (
                    len(extracted_customer)
                    >
                    best_match_length
                ):

                    best_customer = (
                        extracted_customer
                    )

                    best_match_length = (
                        len(extracted_customer)
                    )

    best_customer = re.sub(
        r'\s+',
        ' ',
        best_customer
    ).strip()

    return best_customer.title()


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
# COMMENTS
# ==========================================

def extract_comments():

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

    customer = extract_customer(
        subject
    )

    vertical = extract_vertical(
        customer
    )

    priority = extract_priority(
        subject
    )

    mail_type = extract_mail_type(
        subject
    )

    delivery_type = (
        f"{mail_type} {priority}"
    ).strip()

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