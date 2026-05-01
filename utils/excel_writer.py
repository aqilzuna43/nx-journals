"""Thin xlsxwriter wrapper used by BOM and audit journals."""

import xlsxwriter


_HEADER_BG = "#1a1a2e"
_HEADER_FG = "#ffffff"
_AMBER_BG = "#FFF3CD"


class ExcelWriter:
    """
    Manages a single workbook with multiple sheets.

    Usage::

        writer = ExcelWriter("output.xlsx")
        writer.add_header_row("BOM", ["Level", "Part Number", "Description"])
        writer.add_row("BOM", [1, "PN-001", "Bracket"])
        writer.save()
    """

    def __init__(self, filepath):
        self._filepath = filepath
        self._workbook = xlsxwriter.Workbook(filepath)
        self._sheets = {}
        self._row_counters = {}

        self._header_fmt = self._workbook.add_format({
            "bold": True,
            "font_color": _HEADER_FG,
            "bg_color": _HEADER_BG,
            "border": 1,
        })
        self._cell_fmt = self._workbook.add_format({"border": 1})
        self._amber_fmt = self._workbook.add_format({
            "bg_color": _AMBER_BG,
            "border": 1,
        })

    def _get_or_create_sheet(self, sheet_name):
        if sheet_name not in self._sheets:
            self._sheets[sheet_name] = self._workbook.add_worksheet(sheet_name)
            self._row_counters[sheet_name] = 0
        return self._sheets[sheet_name]

    def add_header_row(self, sheet_name, headers):
        """Write a styled header row (bold, white text, dark blue background)."""
        sheet = self._get_or_create_sheet(sheet_name)
        row = self._row_counters[sheet_name]
        for col, header in enumerate(headers):
            sheet.write(row, col, header, self._header_fmt)
        self._row_counters[sheet_name] += 1

    def add_row(self, sheet_name, values, highlight_amber=False):
        """Append a data row. Pass highlight_amber=True to colour the row amber."""
        sheet = self._get_or_create_sheet(sheet_name)
        row = self._row_counters[sheet_name]
        fmt = self._amber_fmt if highlight_amber else self._cell_fmt
        for col, value in enumerate(values):
            sheet.write(row, col, value, fmt)
        self._row_counters[sheet_name] += 1

    def write_cell(self, sheet_name, row, col, value, highlight_amber=False):
        """Write to an explicit (row, col) position (0-indexed, absolute)."""
        sheet = self._get_or_create_sheet(sheet_name)
        fmt = self._amber_fmt if highlight_amber else self._cell_fmt
        sheet.write(row, col, value, fmt)

    def save(self):
        """Close the workbook and flush to disk."""
        self._workbook.close()
