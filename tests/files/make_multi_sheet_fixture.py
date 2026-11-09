#!/usr/bin/env python

"""
builds the multi-sheet, multi-table xlsx fixture used by the
sheet_filter / range_filter tests

layout:
  Sheet1:
    table at A1  (ID / Name)       -> excel range A1:B3
    table at A10 (Code / Value)    -> excel range A10:B12
  Sheet2:
    table at A1  (Key / Amount)    -> excel range A1:B3
  Sheet3:
    table at A1  (Foo / Bar)       -> excel range A1:B3
"""

from pathlib import Path

from openpyxl import Workbook


def build():
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Sheet1"
    rows1 = [
        ["ID", "Name"],
        [1, "Alpha"],
        [2, "Beta"],
    ]
    for r, row in enumerate(rows1, start=1):
        for c, value in enumerate(row, start=1):
            ws1.cell(row=r, column=c, value=value)

    rows2 = [
        ["Code", "Value"],
        ["X1", 100],
        ["X2", 200],
    ]
    for r, row in enumerate(rows2, start=10):
        for c, value in enumerate(row, start=1):
            ws1.cell(row=r, column=c, value=value)

    ws2 = wb.create_sheet("Sheet2")
    rows3 = [
        ["Key", "Amount"],
        ["K1", 10],
        ["K2", 20],
    ]
    for r, row in enumerate(rows3, start=1):
        for c, value in enumerate(row, start=1):
            ws2.cell(row=r, column=c, value=value)

    ws3 = wb.create_sheet("Sheet3")
    rows4 = [
        ["Foo", "Bar"],
        ["F1", "B1"],
        ["F2", "B2"],
    ]
    for r, row in enumerate(rows4, start=1):
        for c, value in enumerate(row, start=1):
            ws3.cell(row=r, column=c, value=value)

    out = Path(__file__).parent / "multi_sheet_tables.xlsx"
    wb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    build()
