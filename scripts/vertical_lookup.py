import pandas as pd

from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent


MAPPING_FILE = (
    BASE_DIR
    / "customer_vertical_mapping.xlsx"
)


vertical_map = {}

original_company_names = {}


def load_vertical_mapping():

    global vertical_map
    global original_company_names

    vertical_map = {}
    original_company_names = {}

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

    return vertical_map.get(
        customer_name.lower().strip(),
        ""
    )


def get_original_company_name(customer_name):

    if not customer_name:
        return ""

    return original_company_names.get(
        customer_name.lower().strip(),
        customer_name
    )