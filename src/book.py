"""Thin layer over openpyxl: named cells, deferred formula resolution and
banking colour conventions (inputs blue, formulas black, cross-sheet links green).

Formulas are written with tokens that are resolved after every sheet exists:
  {name}        absolute address of a named cell, e.g. Assumptions!$C$3
  {name@D}      column D of a named row (relative), e.g. Projections!D12
  {name@D:I}    range across a named row, e.g. Projections!D12:I12
Tokens pointing to the formula's own sheet are written without a sheet prefix.
"""
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BLUE, BLACK, GREEN = "0000FF", "000000", "008000"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="808080")

FMT = {
    "m": '#,##0.0;(#,##0.0);"–"',          # $ millions
    "m3": '#,##0.000;(#,##0.000);"–"',     # $ millions, 3dp (exact filing thousands)
    "pct": '0.0%;(0.0%);"–"',
    "pct2": '0.00%;(0.00%);"–"',
    "x": '0.0"x"',
    "n": '#,##0;(#,##0);"–"',
    "n3": '#,##0.000',
    "ps": '$#,##0.00',
    "beta": '0.00',
    "num4": '0.0000',
    "date": 'yyyy-mm-dd',
    "gen": 'General',
}

TOKEN = re.compile(r"\{([A-Za-z0-9_]+)(?:@([A-Z]+)(?::([A-Z]+))?)?\}")
PURE_LINK = re.compile(r"^=(?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]+\$?\d+$")


def q(sheet):
    return f"'{sheet}'" if re.search(r"[^A-Za-z0-9_]", sheet) else sheet


class Book:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.names = {}       # name -> (sheet, col or None, row)
        self.pending = []     # (ws, coord, template, fmt)

    def sheet(self, title):
        ws = self.wb.create_sheet(title)
        ws.sheet_view.showGridLines = False
        return ws

    # ---- naming ---------------------------------------------------------
    def name_cell(self, name, ws, col, row):
        assert name not in self.names, f"duplicate name {name}"
        self.names[name] = (ws.title, col, row)

    def name_row(self, name, ws, row):
        assert name not in self.names, f"duplicate name {name}"
        self.names[name] = (ws.title, None, row)

    def addr(self, name, col=None, col2=None, from_sheet=None, absolute=True):
        sheet, c, r = self.names[name]
        prefix = "" if sheet == from_sheet else q(sheet) + "!"
        if c is not None and col is None:
            return f"{prefix}${c}${r}" if absolute else f"{prefix}{c}{r}"
        if col is None:
            raise KeyError(f"row name {name} used without a column")
        if col2:
            return f"{prefix}{col}{r}:{col2}{r}"
        return f"{prefix}{col}{r}"

    # ---- writing --------------------------------------------------------
    def label(self, ws, coord, text, bold=False, italic=False, size=None, color=None, wrap=False, fill=None):
        c = ws[coord]
        c.value = text
        c.font = Font(bold=bold, italic=italic, size=size or 10, color=color or BLACK, name="Calibri")
        if wrap:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        if fill:
            c.fill = fill
        return c

    def inp(self, ws, coord, value, fmt="m", name=None):
        c = ws[coord]
        c.value = value
        c.font = Font(color=BLUE, name="Calibri", size=10)
        c.number_format = FMT.get(fmt, fmt)
        if name:
            col = re.match(r"[A-Z]+", coord).group()
            self.name_cell(name, ws, col, int(coord[len(col):]))
        return c

    def f(self, ws, coord, template, fmt="m", name=None, bold=False):
        c = ws[coord]
        c.number_format = FMT.get(fmt, fmt)
        c.font = Font(bold=bold, name="Calibri", size=10)
        self.pending.append((ws, coord, template, bold))
        if name:
            col = re.match(r"[A-Z]+", coord).group()
            self.name_cell(name, ws, col, int(coord[len(col):]))
        return c

    def resolve(self):
        for ws, coord, tpl, bold in self.pending:
            def sub(m):
                return self.addr(m.group(1), m.group(2), m.group(3), from_sheet=ws.title)
            formula = TOKEN.sub(sub, tpl)
            c = ws[coord]
            c.value = formula
            color = GREEN if PURE_LINK.match(formula) else BLACK
            c.font = Font(bold=bold, color=color, name="Calibri", size=10)

    # ---- layout helpers -------------------------------------------------
    def title(self, ws, text, subtitle=None, width=12):
        ws["A1"].value = text
        ws["A1"].font = Font(bold=True, size=14, color="FFFFFF", name="Calibri")
        for i in range(1, width + 1):
            ws.cell(row=1, column=i).fill = HEAD_FILL
        if subtitle:
            ws["A2"].value = subtitle
            ws["A2"].font = Font(italic=True, size=9, color="595959", name="Calibri")

    def section(self, ws, row, text, width=12):
        ws.cell(row=row, column=1, value=text).font = Font(bold=True, size=11, name="Calibri")
        for i in range(1, width + 1):
            ws.cell(row=row, column=i).fill = SUB_FILL

    def widths(self, ws, spec):
        for col, w in spec.items():
            ws.column_dimensions[col].width = w

    def header_row(self, ws, row, start_col, labels, bold=True):
        for i, t in enumerate(labels):
            c = ws.cell(row=row, column=start_col + i, value=t)
            c.font = Font(bold=bold, name="Calibri", size=10)
            c.alignment = Alignment(horizontal="right")
            c.border = Border(bottom=THIN)


def col(i):
    return get_column_letter(i)
