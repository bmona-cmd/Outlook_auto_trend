import pandas as pd

from pathlib import Path

from openpyxl import load_workbook

from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent.parent


OUTPUT_FILE = (

    BASE_DIR

    / "output"

    / "weekend_cases.xlsx"
)


COLUMNS = [

    "Date",

    "Case#",

    "Customer",

    "Vertical",

    "Technology",

    "Case Delivery Type",

    "Comments"
]


def create_excel_file():

    OUTPUT_FILE.parent.mkdir(exist_ok=True)

    df = pd.DataFrame(columns=COLUMNS)

    df.to_excel(
        OUTPUT_FILE,
        index=False,
        engine="openpyxl"
    )


def adjust_column_width():

    wb = load_workbook(OUTPUT_FILE)

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

    wb.save(OUTPUT_FILE)


def append_to_excel(data):

    if not OUTPUT_FILE.exists():
        create_excel_file()

    try:

        df = pd.read_excel(
            OUTPUT_FILE,
            engine="openpyxl"
        )

    except:

        create_excel_file()

        df = pd.read_excel(
            OUTPUT_FILE,
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
        OUTPUT_FILE,
        index=False,
        engine="openpyxl"
    )

    adjust_column_width()