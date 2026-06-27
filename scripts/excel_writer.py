import pandas as pd
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

BASE_DIR    = Path(__file__).resolve().parent.parent
OUTPUT_FILE = BASE_DIR / "output" / "Weekend_Cases_Tracker.xlsx"

SHEET_SATURDAY = "Sat"
SHEET_SUNDAY   = "Sun"

COLUMNS = [
    "Date", "Case#", "Customer", "Vertical",
    "Technology", "Case Delivery Type", "EM", "Comments"
]

CHART_COLORS = [
    "4472C4", "ED7D31", "A9D18E",
    "FF0000", "FFC000", "70AD47",
    "9E480E", "997300"
]


def get_sheet_name(dt=None):
    weekday = (dt or datetime.now()).weekday()
    if weekday == 5:
        return SHEET_SATURDAY
    elif weekday == 6:
        return SHEET_SUNDAY
    else:
        return SHEET_SATURDAY  # TEST_MODE / weekday run


def get_output_file():
    return OUTPUT_FILE


def _ensure_workbook():
    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    if not OUTPUT_FILE.exists():
        wb = Workbook()
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


def _rebuild_charts_sheet(wb):
    frames = []
    for sheet_name in (SHEET_SATURDAY, SHEET_SUNDAY):
        if sheet_name not in wb.sheetnames:
            continue
        ws   = wb[sheet_name]
        rows = list(ws.values)
        if len(rows) < 2:
            continue
        header = rows[0]
        data   = rows[1:]
        # Deduplicate column names safely
        seen = {}
        clean_header = []
        for col in header:
            col_str = str(col) if col is not None else "unnamed"
            if col_str in seen:
                seen[col_str] += 1
                clean_header.append(f"{col_str}_{seen[col_str]}")
            else:
                seen[col_str] = 0
                clean_header.append(col_str)
        df = pd.DataFrame(data, columns=clean_header).reset_index(drop=True)
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=COLUMNS)

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
            .pipe(lambda s: s[s != ""])
            .value_counts()
            .sort_index()
        )
        if counts.empty:
            return start_row
        t = wc.cell(row=start_row, column=1, value=title)
        t.font = Font(bold=True, size=13, color="1F3864", name="Arial")
        start_row += 1
        for c, val in [(1, col_name), (2, "Count")]:
            cell = wc.cell(row=start_row, column=c, value=val)
            cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
            cell.fill      = PatternFill("solid", start_color="4472C4")
            cell.alignment = Alignment(horizontal="center")
            cell.border    = border
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
        chart = BarChart()
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
        return data_end + 24

    row = 1
    row = write_block("Vertical",           "Cases by Vertical",           row)
    row = write_block("Technology",         "Cases by Technology",         row)
    row = write_block("Case Delivery Type", "Cases by Case Delivery Type", row)


def _find_last_data_row(ws):
    """
    Find the true last row that contains data.
    openpyxl's ws.max_row can be inflated by formatting/ghost rows.
    We scan from the bottom up for the first non-empty row.
    """
    for row in range(ws.max_row, 0, -1):
        for cell in ws[row]:
            if cell.value is not None and str(cell.value).strip():
                return row
    return 1  # only header row exists


def append_to_excel(data, dt=None):
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
        data.get("EM", ""),
        data.get("Comments", ""),
    ]

    # Write to the row immediately after the last real data row
    # This avoids the gap caused by openpyxl max_row inflation
    last_row    = _find_last_data_row(ws)
    target_row  = last_row + 1
    for col_idx, value in enumerate(new_row, start=1):
        ws.cell(row=target_row, column=col_idx, value=value)

    # Highlight the row red if any critical extraction field is blank.
    # Critical fields: Customer (col3), Vertical (col4), Technology (col5),
    # Case Delivery Type (col6). Comments (col8) and EM (col7) are optional.
    CRITICAL_COLS = [3, 4, 5, 6]   # 1-based indices matching new_row order
    has_blank = any(
        not str(new_row[c - 1]).strip()
        for c in CRITICAL_COLS
    )
    if has_blank:
        red_fill = PatternFill("solid", start_color="FFB3B3")   # soft red
        red_font = Font(name="Arial", color="7B0000")
        for col_idx in range(1, len(COLUMNS) + 1):
            cell = ws.cell(row=target_row, column=col_idx)
            cell.fill = red_fill
            cell.font = red_font
        print(f"       [excel] Row {target_row} flagged red (blank fields detected)")

    _adjust_column_width(ws)
    # Save row first — nothing below can block this
    wb.save(OUTPUT_FILE)
    print(f"       [excel] Row saved to sheet '{sheet_name}' row {target_row} ✓")
    # Best-effort chart rebuild
    try:
        wb2 = load_workbook(OUTPUT_FILE)
        _rebuild_charts_sheet(wb2)
        wb2.save(OUTPUT_FILE)
    except Exception as e:
        print(f"       [excel] Chart rebuild skipped: {e}")


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
                try:
                    wb.close()
                except Exception:
                    pass
                return tech is None or not str(tech).strip()
    try:
        wb.close()
    except Exception:
        pass
    return False