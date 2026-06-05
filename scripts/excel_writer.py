import pandas as pd

from pathlib import Path

from openpyxl import load_workbook

from openpyxl.utils import get_column_letter

from datetime import datetime


BASE_DIR = Path(__file__).resolve().parent.parent


COLUMNS = [

    "Date",

    "Case#",

    "Customer",

    "Vertical",

    "Technology",

    "Case Delivery Type",

    "Comments"
]


# ==========================================
# DYNAMIC OUTPUT FILE — SAT / SUN
# ==========================================

def get_output_file():

    # Returns a Path like:
    #   output/saturday_cases.xlsx   (weekday == 5)
    #   output/sunday_cases.xlsx     (weekday == 6)
    # Falls back to weekend_cases.xlsx on weekdays
    # (covers TEST_MODE runs during the week).

    weekday = datetime.now().weekday()

    if weekday == 5:
        filename = "saturday_cases.xlsx"

    elif weekday == 6:
        filename = "sunday_cases.xlsx"

    else:
        filename = "weekend_cases.xlsx"

    return BASE_DIR / "output" / filename


# ==========================================
# CREATE EMPTY FILE
# ==========================================

def create_excel_file(output_file):

    output_file.parent.mkdir(exist_ok=True)

    df = pd.DataFrame(columns=COLUMNS)

    df.to_excel(
        output_file,
        index=False,
        engine="openpyxl"
    )


# ==========================================
# AUTO-FIT COLUMN WIDTHS
# ==========================================

def adjust_column_width(output_file):

    wb = load_workbook(output_file)

    ws = wb.active

    for column_cells in ws.columns:

        max_length = 0

        column = column_cells[0].column

        for cell in column_cells:

            try:

                if cell.value:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            except:
                pass

        adjusted_width = min(
            max_length + 5,
            50
        )

        ws.column_dimensions[
            get_column_letter(column)
        ].width = adjusted_width

    wb.save(output_file)


# ==========================================
# APPEND ROW
# ==========================================

def append_to_excel(data):

    output_file = get_output_file()

    if not output_file.exists():
        create_excel_file(output_file)

    try:

        df = pd.read_excel(
            output_file,
            engine="openpyxl"
        )

    except:

        create_excel_file(output_file)

        df = pd.read_excel(
            output_file,
            engine="openpyxl"
        )

    new_row = {

        "Date": data.get("Date", ""),

        "Case#": data.get("Case#", ""),

        "Customer": data.get("Customer", ""),

        "Vertical": data.get("Vertical", ""),

        "Technology": data.get("Technology", ""),

        "Case Delivery Type": data.get(
            "Case Delivery Type",
            ""
        ),

        "Comments": data.get("Comments", "")
    }

    df.loc[len(df)] = new_row

    df.to_excel(
        output_file,
        index=False,
        engine="openpyxl"
    )

    adjust_column_width(output_file)


def technology_missing_for_case(case_number):

    if not case_number:
        return False

    output_file = get_output_file()

    if not output_file.exists():
        return False

    try:

        df = pd.read_excel(
            output_file,
            engine="openpyxl"
        )

    except:

        return False

    if (
        "Case#" not in df.columns
        or
        "Technology" not in df.columns
    ):
        return False

    rows = df[
        df["Case#"].astype(str).str.strip()
        == str(case_number).strip()
    ]

    if rows.empty:
        return False

    technology = rows.iloc[-1].get(
        "Technology",
        ""
    )

    return (
        pd.isna(technology)
        or
        not str(technology).strip()
    )
