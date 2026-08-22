from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from mrg2opus.excel_io.style_utils import is_blacked_out, is_excluded, is_struck_through


def test_strikethrough_cell_is_detected():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "cancelled rate"
    ws["A1"].font = Font(strike=True)
    ws["A2"] = "normal rate"

    assert is_struck_through(ws["A1"])
    assert not is_struck_through(ws["A2"])
    assert is_excluded(ws["A1"])
    assert not is_excluded(ws["A2"])


def test_blacked_out_fill_is_detected():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "hidden"
    ws["A1"].fill = PatternFill(start_color="FF000000", end_color="FF000000", fill_type="solid")
    ws["A2"] = "visible"
    ws["A2"].fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

    assert is_blacked_out(ws["A1"])
    assert not is_blacked_out(ws["A2"])
    assert is_excluded(ws["A1"])
