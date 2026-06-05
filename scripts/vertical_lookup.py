import pandas as pd

from pathlib import Path

from collections import Counter


BASE_DIR = Path(__file__).resolve().parent.parent


MAPPING_FILE = (
    BASE_DIR
    / "customer_vertical_mapping.xlsx"
)


vertical_map = {}

original_company_names = {}

normalized_vertical_map = {}

normalized_original_company_names = {}

first_word_vertical_map = {}


def normalize_customer_name(customer_name):

    if not customer_name:
        return ""

    customer_name = str(customer_name).lower().strip()

    customer_name = customer_name.replace(".", " ")

    import re

    customer_name = re.sub(
        r'[\|\:\,\-\/\[\]\(\)_]+',
        ' ',
        customer_name
    )

    customer_name = re.sub(
        r'\s+',
        ' ',
        customer_name
    )

    return customer_name.strip()


def load_vertical_mapping():

    global vertical_map
    global original_company_names
    global normalized_vertical_map
    global normalized_original_company_names
    global first_word_vertical_map

    vertical_map = {}
    original_company_names = {}
    normalized_vertical_map = {}
    normalized_original_company_names = {}
    first_word_vertical_map = {}
    first_word_candidates = {}
    first_word_counts = {}

    if not MAPPING_FILE.exists():

        print(
            f"Mapping file not found: {MAPPING_FILE}"
        )

        return

    try:

        df = pd.read_excel(
            MAPPING_FILE,
            engine="openpyxl"
        )

        # CLEAN COLUMN NAMES

        df.columns = [

            str(col)
            .strip()
            .lower()

            for col in df.columns
        ]

        print(
            "Excel Columns Found:"
        )

        print(
            df.columns.tolist()
        )

        # FLEXIBLE COLUMN DETECTION

        customer_col = None
        vertical_col = None

        for col in df.columns:

            if (
                "customer" in col
                or
                "company" in col
            ):

                customer_col = col

            if "vertical" in col:

                vertical_col = col

        if not customer_col:

            print(
                "Customer column not found"
            )

            return

        if not vertical_col:

            print(
                "Vertical column not found"
            )

            return

        for _, row in df.iterrows():

            customer = str(
                row[customer_col]
            ).strip()

            vertical = str(
                row[vertical_col]
            ).strip()

            if not customer:
                continue

            customer_lower = (
                customer.lower()
                .strip()
            )

            vertical_map[
                customer_lower
            ] = vertical

            original_company_names[
                customer_lower
            ] = customer

            normalized_customer = normalize_customer_name(
                customer
            )

            if normalized_customer:

                normalized_vertical_map[
                    normalized_customer
                ] = vertical

                normalized_original_company_names[
                    normalized_customer
                ] = customer

                first_word = normalized_customer.split()[0]

                first_word_candidates.setdefault(
                    first_word,
                    set()
                ).add(vertical)

                first_word_counts.setdefault(
                    first_word,
                    Counter()
                )[vertical] += 1

        for first_word, verticals in first_word_candidates.items():

            if len(verticals) == 1:

                first_word_vertical_map[
                    first_word
                ] = next(iter(verticals))

                continue

            most_common_vertical, count = first_word_counts[
                first_word
            ].most_common(1)[0]

            total_count = sum(
                first_word_counts[first_word].values()
            )

            if count / total_count >= 0.8:

                first_word_vertical_map[
                    first_word
                ] = most_common_vertical

        print(
            f"Loaded {len(vertical_map)} customers"
        )

    except Exception as e:

        print(
            "Vertical mapping load failed"
        )

        print(e)


load_vertical_mapping()


def get_vertical(customer_name):

    if not customer_name:
        return ""

    customer_key = customer_name.lower().strip()

    vertical = vertical_map.get(
        customer_key,
        ""
    )

    if vertical:
        return vertical

    normalized_customer = normalize_customer_name(
        customer_name
    )

    vertical = normalized_vertical_map.get(
        normalized_customer,
        ""
    )

    if vertical:
        return vertical

    if normalized_customer:

        first_word = normalized_customer.split()[0]

        return first_word_vertical_map.get(
            first_word,
            ""
        )

    return ""


def get_original_company_name(customer_name):

    if not customer_name:
        return ""

    customer_key = customer_name.lower().strip()

    original_name = original_company_names.get(
        customer_key,
        ""
    )

    if original_name:
        return original_name

    normalized_customer = normalize_customer_name(
        customer_name
    )

    return normalized_original_company_names.get(
        normalized_customer,
        customer_name
    )
