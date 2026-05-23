import re
from difflib import SequenceMatcher

from vertical_lookup import (
    get_vertical,
    vertical_map
)


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

    subject = subject.replace(
        "\n",
        " "
    )

    subject = subject.replace(
        "\r",
        " "
    )

    subject = re.sub(
        r'\s+',
        ' ',
        subject
    )

    return subject.strip()


# ==========================================
# SKIP REPLY MAILS
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

    subject = clean_subject(
        subject
    )

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
# DELIVERY TYPE
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
# TECHNOLOGY
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
# FUZZY SIMILARITY
# ==========================================

def similarity(a, b):

    return SequenceMatcher(
        None,
        a.lower(),
        b.lower()
    ).ratio()



# ==========================================
# CUSTOMER EXTRACTION
# RETURNS CUSTOMER TEXT FROM SUBJECT
# ==========================================

def extract_customer(subject):

    subject = clean_subject(
        subject
    )

    if not subject:
        return ""

    lower_subject = subject.lower()

    normalized_subject = re.sub(
        r'[\|\:\,\-\/]+',
        ' ',
        lower_subject
    )

    normalized_subject = re.sub(
        r'\s+',
        ' ',
        normalized_subject
    ).strip()

    best_match = ""
    best_score = 0

    extracted_customer = ""

    subject_tokens = normalized_subject.split()

    for excel_customer in vertical_map.keys():

        excel_customer_clean = (
            excel_customer
            .lower()
            .strip()
        )

        customer_words = (
            excel_customer_clean.split()
        )

        # =====================================
        # DIRECT CONTAINMENT
        # =====================================

        if excel_customer_clean in normalized_subject:

            if len(excel_customer_clean) > len(best_match):

                best_match = excel_customer_clean
                best_score = 1.0

                start_index = normalized_subject.find(
                    excel_customer_clean
                )

                remaining = normalized_subject[
                    start_index:
                ]

                extracted_customer = remaining.split(
                    "|"
                )[0].strip()

        # =====================================
        # TOKEN OVERLAP
        # =====================================

        common_words = (

            set(subject_tokens)
            &
            set(customer_words)
        )

        if common_words:

            overlap_ratio = (
                len(common_words)
                /
                len(customer_words)
            )

            if overlap_ratio >= 0.6:

                if overlap_ratio > best_score:

                    best_match = excel_customer_clean
                    best_score = overlap_ratio

                    matched_words = []

                    for token in subject_tokens:

                        if (
                            token in common_words
                            or
                            len(matched_words) > 0
                        ):

                            matched_words.append(
                                token
                            )

                            if len(matched_words) >= 4:
                                break

                    extracted_customer = (
                        " ".join(matched_words)
                    )

        # =====================================
        # FUZZY MATCH
        # =====================================

        ratio = similarity(
            normalized_subject,
            excel_customer_clean
        )

        if ratio >= 0.82:

            if ratio > best_score:

                best_match = excel_customer_clean
                best_score = ratio

                extracted_customer = (
                    excel_customer_clean
                )

    if extracted_customer:

        extracted_customer = re.sub(
            r'\s+',
            ' ',
            extracted_customer
        ).strip()

        extracted_customer = extracted_customer.title()

        # REMOVE TRAILING GARBAGE

        garbage = [

            "queue",
            "dispatch",
            "handover",
            "interface",
            "packet",
            "flap",
            "errors",
            "drops",
            "failure",
            "notice"
        ]

        cleaned_words = []

        for word in extracted_customer.split():

            if word.lower() in garbage:
                break

            cleaned_words.append(word)

        extracted_customer = (
            " ".join(cleaned_words)
        ).strip()

        return extracted_customer

    return ""



# ==========================================
# COMMENTS
# ==========================================

def extract_comments(subject, customer=""):

    subject = clean_subject(
        subject
    )

    if not subject:
        return ""

    working = subject

    # REMOVE CASE NUMBER

    case_number = extract_case_number(
        working
    )

    if case_number:

        working = working.replace(
            case_number,
            " "
        )

    # REMOVE PRIORITY

    priority = extract_priority(
        working
    )

    if priority:

        working = re.sub(
            rf'\b{priority}\b',
            ' ',
            working,
            flags=re.IGNORECASE
        )

    # REMOVE DELIVERY TYPE

    mail_type = extract_mail_type(
        working
    )

    if mail_type:

        working = re.sub(
            mail_type,
            ' ',
            working,
            flags=re.IGNORECASE
        )

    # REMOVE CUSTOMER

    if customer:

        working = re.sub(
            re.escape(customer),
            ' ',
            working,
            flags=re.IGNORECASE
        )

    # REMOVE DEVICE NAMES

    for keyword in DEVICE_TECH_MAP.keys():

        working = re.sub(
            rf'\b{keyword}\d*\b',
            ' ',
            working,
            flags=re.IGNORECASE
        )

    # REMOVE GARBAGE WORDS

    garbage_words = [

        "queue",
        "dispatch",
        "handover",
        "monitoring",
        "vonage",
        "munich",
        "monich",
        "rtp",
        "emea",
        "apac",
        "tac",
        "notice",
        "medium",
        "high",
        "low",
        "urgent",
        "case",
        "in",
        "to",
        "from",
        "reminder"
    ]

    for word in garbage_words:

        working = re.sub(
            rf'\b{word}\b',
            ' ',
            working,
            flags=re.IGNORECASE
        )

    # REMOVE ALL KNOWN CUSTOMERS

    for company in vertical_map.keys():

        working = re.sub(
            re.escape(company),
            ' ',
            working,
            flags=re.IGNORECASE
        )

    separators = [
        "||",
        "|",
        ":",
        "-"
    ]

    for sep in separators:

        working = working.replace(
            sep,
            " "
        )

    working = re.sub(
        r'\s+',
        ' ',
        working
    ).strip()

    if not working:
        return ""

    words = working.split()

    if len(words) <= 1:
        return ""

    return working



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

    vertical = ""

    # ======================================
    # FIND MATCHING EXCEL COMPANY
    # ======================================

    for company in vertical_map.keys():

        if company.lower() in customer.lower():

            vertical = get_vertical(
                company
            )

            break

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

        "Comments": extract_comments(
            subject,
            customer
        )
    }