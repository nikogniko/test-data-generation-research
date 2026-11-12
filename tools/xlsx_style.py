"""Спільне оформлення аркушів Excel: заголовок, ширини колонок, рамки, закріплений перший рядок."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HEAD = PatternFill("solid", fgColor="1F4E79")
HEAD_FONT = Font(bold=True, color="FFFFFF")
FILL_INPUT = PatternFill("solid", fgColor="FFF8E1")
BOLD = Font(bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")


def header(ws, headers, widths=None, row=1, freeze=True):
    for col, name in enumerate(headers, start=1):
        c = ws.cell(row=row, column=col, value=name)
        c.fill, c.font, c.border = HEAD, HEAD_FONT, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        if widths:
            ws.column_dimensions[c.column_letter].width = widths[col - 1]
    ws.row_dimensions[row].height = 45
    if freeze:
        ws.freeze_panes = ws.cell(row=row + 1, column=1)


def body(ws, rows, start=2):
    for r, values in enumerate(rows, start=start):
        for col, v in enumerate(values, start=1):
            c = ws.cell(row=r, column=col, value=v)
            c.border, c.alignment = BORDER, WRAP


def sheet(wb, title, headers, rows, widths=None, first=False):
    """Додає аркуш із заголовком, рядками й фільтром; first=True використовує перший (порожній) аркуш книги."""
    ws = wb.active if first else wb.create_sheet(title)
    ws.title = title
    header(ws, headers, widths)
    body(ws, rows)
    if rows:
        ws.auto_filter.ref = ws.dimensions
    return ws


def text_sheet(wb, title, lines, width=120, first=False):
    """Аркуш з текстом по рядку в комірці: перший рядок — заголовок."""
    ws = wb.active if first else wb.create_sheet(title)
    ws.title = title
    for k, line in enumerate(lines, start=1):
        c = ws.cell(row=k, column=1, value=line)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if k == 1:
            c.font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width = width
    return ws
