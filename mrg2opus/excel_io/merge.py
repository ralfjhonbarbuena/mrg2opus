"""Merge multiple uploaded MRG workbooks into one combined workbook before
classification/parsing.

Real-world lanes sometimes arrive as more than one file - e.g. CSE ships
as a main "Tier 1" file plus a separate "...for VELAG and VEPBL" file
covering Venezuela (VELAG/VEPBL are Venezuelan port codes - La Guaira and
Puerto Cabello). The paired sample workbooks already bundle multiple
real-world files' sheets into one file for convenience (see cse.py's "CSE
VE" sheet docstring note, and eaf.py's TZDAR/KEMBA note) - merging
uploads the same way lets the existing single-workbook parser pipeline
handle them completely unchanged, no parser-side code needed.
"""
from __future__ import annotations

from copy import copy

import openpyxl
from openpyxl.workbook import Workbook


class DuplicateSheetError(Exception):
    """Two uploaded files both contain a sheet with the same name - can't
    tell which one should win, so this is surfaced rather than silently
    picking one (or worse, silently overwriting)."""


def merge_workbooks(workbooks: list[Workbook]) -> Workbook:
    if not workbooks:
        raise ValueError("merge_workbooks() needs at least one workbook")
    if len(workbooks) == 1:
        return workbooks[0]

    merged = openpyxl.Workbook()
    merged.remove(merged.active)
    seen: set[str] = set()
    for wb in workbooks:
        for sheet_name in wb.sheetnames:
            if sheet_name in seen:
                raise DuplicateSheetError(
                    f"Sheet {sheet_name!r} appears in more than one uploaded file - "
                    "can't tell which one should win. Remove the duplicate and try again."
                )
            seen.add(sheet_name)
            _copy_sheet(wb[sheet_name], merged, sheet_name)
    return merged


def _copy_sheet(src_ws, dst_wb: Workbook, sheet_name: str) -> None:
    dst_ws = dst_wb.create_sheet(sheet_name)
    # Touch every cell in the source's used range (not just non-blank
    # ones) so the destination sheet's max_row/max_column match the
    # source exactly - some parsers (e.g. lawc.py's single-column sheet
    # parser) iterate up to ws.max_row/max_column directly, and a
    # shrunken destination range would silently truncate them.
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))
