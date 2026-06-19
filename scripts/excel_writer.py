import pandas as pd

from pathlib import Path

from openpyxl import load_workbook, Workbook

from openpyxl.utils import get_column_letter

from openpyxl.chart import BarChart, Reference

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from datetime import datetime


BASE_DIR = Path(__file__).resolve().parent.parent

# ==========================================
# ONE PERSISTENT FILE — TWO SHEETS
# ==========================================

OUTPUT_FILE = BASE_DIR / "output" / "Weekend_Cases_Tracker.xlsx"

SHEET_SATURDAY = "Sat"
SHEET_SUNDAY   = "Sun"

COLUMNS = [
    "Date",
    "Case#",
    "Customer",
    "Vertical",
    "Technology",
    "Case Delivery Type",
    "Comments"
]

CHART_COLORS = [
    "4472C4", "ED7D31", "A9D18E",
    "FF0000", "FFC000", "70AD47",
    "9E480E", "997300"
]


# ==========================================
# WHICH SHEET TO WRITE TO
# ==========================================

def get_sheet_name(dt=None):
    """Return 'Saturday' or 'Sunday' based on today (or given datetime)."""
    weekday = (dt or datetime.now()).weekday()
    if weekday == 5:
        return SHEET_SATURDAY
    elif weekday == 6:
        return SHEET_SUNDAY
    else:
        # TEST_MODE / weekday run — default to Saturday
        return SHEET_SATURDAY


def get_output_file():
    """Always returns the single persistent tracker file path."""
    return OUTPUT_FILE


# ==========================================
# CREATE / INITIALISE WORKBOOK
# ==========================================

def _ensure_workbook():
    """
    If the file does not exist, create it with both sheets.
    If it exists but is missing a sheet, add the missing sheet.
    Returns the open workbook (caller must save).
    """
    OUTPUT_FILE.parent.mkdir(exist_ok=True)

    if not OUTPUT_FILE.exists():
        wb = Workbook()
        # openpyxl creates a default 'Sheet' — rename it
        ws_sat = wb.active
        ws_sat.title = SHEET_SATURDAY
        ws_sat.append(COLUMNS)
        ws_sun = wb.create_sheet(SHEET_SUNDAY)
        ws_sun.append(COLUMNS)
        wb.save(OUTPUT_FILE)
        return wb

    wb = load_workbook(OUTPUT_FILE)

    changed = False
    for sheet_name in (SHEET_SATURDAY, SHEET_SUNDAY):
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            ws.append(COLUMNS)
            changed = True

    if changed:
        wb.save(OUTPUT_FILE)

    return wb


# ==========================================
# AUTO-FIT COLUMN WIDTHS FOR A WORKSHEET
# ==========================================

def _adjust_column_width(ws):
    for column_cells in ws.columns:
        max_length = 0
        col_idx    = column_cells[0].column
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 5, 50)


# ==========================================
# BUILD / REFRESH CHARTS SHEET
# ==========================================

def _rebuild_charts_sheet(wb):
    """
    Reads data from Saturday + Sunday sheets and writes a Charts sheet
    with bar charts for Vertical, Technology, and Case Delivery Type.
    Called automatically after every append.
    """
    # Collect data from both day sheets
    frames = []
    for sheet_name in (SHEET_SATURDAY, SHEET_SUNDAY):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        rows = list(ws.values)
        if len(rows) < 2:
            continue
        header = rows[0]
        data   = rows[1:]
        df     = pd.DataFrame(data, columns=header)
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=COLUMNS)

    # Remove / recreate Charts sheet
    if "Charts" in wb.sheetnames:
        del wb["Charts"]
    wc = wb.create_sheet("Charts")

    wc.column_dimensions["A"].width = 26
    wc.column_dimensions["B"].width = 12

    thin   = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_block(col_name, title, start_row):
        if col_name not in combined.columns or combined.empty:
            return start_row

        counts = (
            combined[col_name]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", None)
            .dropna()
            .value_counts()
            .sort_index()
        )
        if counts.empty:
            return start_row

        # Section title
        t = wc.cell(row=start_row, column=1, value=title)
        t.font = Font(bold=True, size=13, color="1F3864", name="Arial")
        start_row += 1

        # Header row
        for c, val in [(1, col_name), (2, "Count")]:
            cell = wc.cell(row=start_row, column=c, value=val)
            cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
            cell.fill      = PatternFill("solid", start_color="4472C4")
            cell.alignment = Alignment(horizontal="center")
            cell.border    = border

        # Data rows
        for i, (label, count) in enumerate(counts.items(), 1):
            r  = start_row + i
            lc = wc.cell(row=r, column=1, value=label)
            cc = wc.cell(row=r, column=2, value=int(count))
            for cell in [lc, cc]:
                cell.font      = Font(name="Arial", size=10)
                cell.border    = border
                cell.alignment = Alignment(
                    horizontal="left" if cell.column == 1 else "center"
                )
                if i % 2 == 0:
                    cell.fill = PatternFill("solid", start_color="DCE6F1")

        data_end = start_row + len(counts)

        # Bar chart
        chart          = BarChart()
        chart.type     = "col"
        chart.grouping = "clustered"
        chart.title    = title
        chart.y_axis.title = "Cases"
        chart.width    = 18
        chart.height   = 12
        chart.style    = 10
        chart.add_data(
            Reference(wc, min_col=2, max_col=2, min_row=start_row, max_row=data_end),
            titles_from_data=True
        )
        chart.set_categories(
            Reference(wc, min_col=1, min_row=start_row + 1, max_row=data_end)
        )
        chart.series[0].graphicalProperties.solidFill = CHART_COLORS[0]
        wc.add_chart(chart, f"D{start_row}")

        return data_end + 24   # leave room below chart

    row = 1
    row = write_block("Vertical",           "Cases by Vertical",           row)
    row = write_block("Technology",         "Cases by Technology",         row)
    row = write_block("Case Delivery Type", "Cases by Case Delivery Type", row)


# ==========================================
# APPEND ROW
# ==========================================

def append_to_excel(data, dt=None):
    """
    Append one case row to the correct day sheet (Saturday or Sunday).
    Auto-creates the workbook if it does not exist.
    Rebuilds the Charts sheet after every write.
    """
    sheet_name = get_sheet_name(dt)

    wb = _ensure_workbook()
    ws = wb[sheet_name]

    new_row = [
        data.get("Date", ""),
        data.get("Case#", ""),
        data.get("Customer", ""),
        data.get("Vertical", ""),
        data.get("Technology", ""),
        data.get("Case Delivery Type", ""),
        data.get("Comments", ""),
    ]
    ws.append(new_row)

    _adjust_column_width(ws)
    _rebuild_charts_sheet(wb)

    wb.save(OUTPUT_FILE)


# ==========================================
# CHECK IF TECHNOLOGY IS MISSING FOR A CASE
# ==========================================

def technology_missing_for_case(case_number):
    if not case_number or not OUTPUT_FILE.exists():
        return False

    try:
        wb = load_workbook(OUTPUT_FILE, read_only=True)
    except Exception:
        return False

    for sheet_name in (SHEET_SATURDAY, SHEET_SUNDAY):
        if sheet_name not in wb.sheetnames:
            continue
        ws     = wb[sheet_name]
        rows   = list(ws.values)
        if len(rows) < 2:
            continue
        header = list(rows[0])
        try:
            case_idx = header.index("Case#")
            tech_idx = header.index("Technology")
        except ValueError:
            continue
        for row in reversed(rows[1:]):
            if str(row[case_idx]).strip() == str(case_number).strip():
                tech = row[tech_idx]
                wb.close()
                return tech is None or not str(tech).strip()

    try:
        wb.close()
    except Exception:
        pass
    return False