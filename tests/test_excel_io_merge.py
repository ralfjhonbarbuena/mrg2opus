from __future__ import annotations

import openpyxl
import pytest

from mrg2opus.excel_io.merge import DuplicateSheetError, merge_workbooks


def _wb_with_sheet(name: str, values: dict[tuple[int, int], object]) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(name)
    for (row, col), value in values.items():
        ws.cell(row=row, column=col, value=value)
    return wb


def test_single_workbook_passthrough():
    wb = _wb_with_sheet("A", {(1, 1): "x"})
    assert merge_workbooks([wb]) is wb


def test_merges_sheets_from_multiple_workbooks():
    wb1 = _wb_with_sheet("CSE", {(1, 1): "main"})
    wb2 = _wb_with_sheet("CSE VE", {(1, 1): "ve"})

    merged = merge_workbooks([wb1, wb2])

    assert set(merged.sheetnames) == {"CSE", "CSE VE"}
    assert merged["CSE"].cell(row=1, column=1).value == "main"
    assert merged["CSE VE"].cell(row=1, column=1).value == "ve"


def test_preserves_merged_cell_ranges():
    wb1 = openpyxl.Workbook()
    wb1.remove(wb1.active)
    ws1 = wb1.create_sheet("Sheet1")
    ws1["A1"] = "spans"
    ws1.merge_cells("A1:C1")
    wb2 = _wb_with_sheet("Sheet2", {(1, 1): "y"})

    merged = merge_workbooks([wb1, wb2])

    assert "A1:C1" in [str(r) for r in merged["Sheet1"].merged_cells.ranges]


def test_preserves_font_and_fill_for_exclusion_detection():
    wb1 = openpyxl.Workbook()
    wb1.remove(wb1.active)
    ws1 = wb1.create_sheet("Sheet1")
    from openpyxl.styles import Font

    cell = ws1.cell(row=1, column=1, value=100)
    cell.font = Font(strike=True)
    wb2 = _wb_with_sheet("Sheet2", {(1, 1): "y"})

    merged = merge_workbooks([wb1, wb2])

    assert merged["Sheet1"].cell(row=1, column=1).font.strike is True


def test_duplicate_sheet_name_raises():
    wb1 = _wb_with_sheet("CSE", {(1, 1): "main"})
    wb2 = _wb_with_sheet("CSE", {(1, 1): "other"})

    with pytest.raises(DuplicateSheetError):
        merge_workbooks([wb1, wb2])
