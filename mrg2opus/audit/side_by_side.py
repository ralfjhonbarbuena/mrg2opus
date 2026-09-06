"""Write both drafts into one workbook, ready for the auditor's own
CONCATENATE-and-XLOOKUP reconciliation.

The diff on screen says what the tool thinks differs. This is the other
half: the two sheets themselves, side by side in one file, each carrying
the match key already built and an XLOOKUP already pointing at the other
sheet - so the reconciliation can be done and re-checked by hand, in the
form the audit already uses, without trusting our diff to have asked the
right question.

Sheet pairs come out as "<name> (ours)" and "<name> (ref)".
"""
from __future__ import annotations

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from mrg2opus.audit.compare import AUDIT_CONCAT_FIELDS, audit_concat
from mrg2opus.schema import opus_columns as cols

# Excel's own cap; a scoped port-port pair ("RATES-OEW PORT-PORT (ref)")
# runs to 25, so this only ever bites on an unusually long sub-lane name.
_MAX_SHEET_NAME = 31

_KEY_HEADER = "MATCH KEY"
_STATUS_HEADER = "IN THE OTHER DRAFT?"

_HEADER_FILL = PatternFill("solid", start_color=cols.HEADER_FILL_RGB, end_color=cols.HEADER_FILL_RGB)
_HEADER_FONT = Font(name=cols.HEADER_FONT_NAME, size=cols.HEADER_FONT_SIZE, color=cols.HEADER_FONT_RGB)
_KEY_FILL = PatternFill("solid", start_color="FFF2E8F0", end_color="FFF2E8F0")


def _readable_headers() -> list[str]:
    """One header row rather than the OPUS two, because a lookup range is
    easier to write against a single row. Group and field labels are
    joined so it still reads like the sheet it came from."""
    headers = []
    for group, field in zip(cols.RATES_HEADER_GROUP, cols.RATES_HEADER_FIELD):
        group_label = (group or "").replace("\n", " ").strip()
        field_label = (field or "").replace("\n", " ").strip()
        headers.append(f"{group_label} {field_label}".strip() if field_label else group_label)
    return headers


def _sheet_name(base: str, side: str) -> str:
    name = f"{base} ({side})"
    return name if len(name) <= _MAX_SHEET_NAME else f"{base[: _MAX_SHEET_NAME - len(side) - 3]} ({side})"


def _write_side(wb: Workbook, title: str, rows: list[dict], other_title: str, other_label: str) -> None:
    ws = wb.create_sheet(title[:_MAX_SHEET_NAME])
    ws.append([_KEY_HEADER, _STATUS_HEADER, *_readable_headers()])
    for cell in ws[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = cols.HEADER_ROW_HEIGHT

    quoted = other_title.replace("'", "''")
    for i, row in enumerate(rows, start=2):
        # The lookup is written out rather than the answer, so it stays
        # live: edit either sheet in Excel and the column re-evaluates.
        lookup = (
            f"=XLOOKUP($A{i},'{quoted}'!$A:$A,'{quoted}'!$A:$A,\"NOT IN {other_label}\")"
        )
        ws.append([
            audit_concat(row), lookup,
            *[row.get(field) for field in cols.RATES_ROW_FIELDS],
        ])
        ws.cell(row=i, column=1).fill = _KEY_FILL

    ws.freeze_panes = "C2"
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 22
    for idx in range(3, len(cols.RATES_ROW_FIELDS) + 3):
        ws.column_dimensions[get_column_letter(idx)].width = 14
    ws.auto_filter.ref = ws.dimensions


def build_side_by_side_workbook(pairs: list[tuple[str, list[dict], list[dict]]]) -> bytes:
    """pairs: (sheet label, our rows, the reference's rows)."""
    wb = Workbook()
    wb.remove(wb.active)

    notes = wb.create_sheet("HOW TO USE")
    for line in (
        ["Both drafts, one workbook - reconcile them the way the audit already does."],
        [],
        ["MATCH KEY (column A) is the CONCATENATE, already built:"],
        ["  " + " | ".join(AUDIT_CONCAT_FIELDS)],
        [],
        ["IN THE OTHER DRAFT? (column B) is a live XLOOKUP against the paired sheet."],
        ["  It returns the key when the row is found, or NOT IN ... when it isn't."],
        ["  Both columns re-evaluate if you edit either sheet."],
        [],
        ["A row present on both sheets but with a different rate WILL match on the key:"],
        ["  the key deliberately leaves the rate values out, so a rate error shows up as"],
        ["  a matched row whose rate columns differ, not as a row that went missing."],
        ["  Compare the rate columns directly, or use the on-screen substance list."],
    ):
        notes.append(line)
    notes.column_dimensions["A"].width = 100
    notes["A1"].font = Font(bold=True)

    for label, ours, theirs in pairs:
        ours_title = _sheet_name(label, "ours")
        ref_title = _sheet_name(label, "ref")
        _write_side(wb, ours_title, ours, ref_title, "REFERENCE")
        _write_side(wb, ref_title, theirs, ours_title, "YOUR DRAFT")

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
