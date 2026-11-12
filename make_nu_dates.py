from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

# ---------- Styling helpers ----------
header_fill = PatternFill("solid", fgColor="1F4E78")
header_font = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
title_font = Font(bold=True, color="1F4E78", size=14, name="Calibri")
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin = Side(style="thin", color="9CA3AF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

def autowidth(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

# =====================================================================
# Sheet: nu-dates (only 5 columns)
# =====================================================================
ws = wb.active
ws.title = "nu-dates"
ws.sheet_view.showGridLines = False

ws["A1"] = "NU Notice Dates Summary"
ws["A1"].font = title_font
ws.merge_cells("A1:E1")
ws.row_dimensions[1].height = 22

headers = ["Program", "Year", "Notice Type", "Session", "Notice Publish Date"]
ncols = len(headers)

r = 3
for c, h in enumerate(headers, start=1):
    ws.cell(row=r, column=c, value=h)
style_header_row(ws, r, ncols)
r += 1
first_data_row = r

data = [
    ["Honours", "1st Year", "Form Fill-up", "2024-25", "2026-05-06"],
    ["Honours", "1st Year", "Routine", "2024-25", "2026-06-10"],
    ["Honours", "1st Year", "Result", "2023-24", "2025-12-23"],
    ["Honours", "2nd Year", "Form Fill-up", "2023-24", "2026-06-29"],
    ["Honours", "2nd Year", "Routine", "2023-24", "2026-08-26"],
    ["Honours", "2nd Year", "Result", "2022-23", "2026-03-07"],
    ["Honours", "3rd Year", "Form Fill-up", "2022-23", "2026-10-04"],
    ["Honours", "3rd Year", "Routine", "2022-23", "2026-10-07"],
    ["Honours", "3rd Year", "Result", "2021-22", "2026-05-11"],
    ["Honours", "4th Year", "Form Fill-up", "2021-22", "2026-11-01"],
    ["Honours", "4th Year", "Routine", "2020-21", "2026-03-11"],
    ["Honours", "4th Year", "Result", "2020-21", "2026-08-31"],
    ["Degree", "1st Year", "Form Fill-up", "2024-25", "2026-07-27"],
    ["Degree", "1st Year", "Routine", "2024-25", "2026-08-18"],
    ["Degree", "1st Year", "Result", "2023-24", "2026-02-28"],
    ["Degree", "2nd Year", "Form Fill-up", "2023-24", "2026-09-29"],
    ["Degree", "2nd Year", "Routine", "2023-24", "2026-10-12"],
    ["Degree", "2nd Year", "Result", "2022-23", "2026-07-11"],
    ["Degree", "3rd Year", "Form Fill-up", "2021-22", "2026-05-17"],
    ["Degree", "3rd Year", "Routine", "2021-22", "2026-06-11"],
    ["Degree", "3rd Year", "Result", "2021-22", "2026-09-24"],
    ["Preliminary to Masters", "Preliminary", "Form Fill-up", "2022-23", "2026-04-30"],
    ["Preliminary to Masters", "Preliminary", "Routine", "2022-23", "2026-06-03"],
    ["Preliminary to Masters", "Preliminary", "Result", "2022-23", "2026-10-06"],
    ["Masters", "Final", "Form Fill-up", "2021-22", "2025-12-23"],
    ["Masters", "Final", "Routine", "2021-22", "2026-03-14"],
    ["Masters", "Final", "Result", "2021-22", "2026-08-06"],
]

for row in data:
    for c, val in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=val)
        cell.alignment = wrap
        cell.border = border
    r += 1

for rr in range(first_data_row, r):
    for cc in range(1, ncols + 1):
        ws.cell(row=rr, column=cc).border = border

autowidth(ws, [14, 12, 16, 14, 20])
ws.freeze_panes = "A4"

out = "nu-dates.xlsx"
wb.save(out)
print("Saved:", out)
