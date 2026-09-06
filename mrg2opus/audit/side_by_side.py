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
from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from mrg2opus.audit.compare import AUDIT_CONCAT_FIELDS
from mrg2opus.schema import opus_columns as cols

# Excel's own cap; a scoped port-port pair ("RATES-OEW PORT-PORT (ref)")
# runs to 25, so this only ever bites on an unusually long sub-lane name.
_MAX_SHEET_NAME = 31

# Column A is the key, B the lookup, so the data starts at C.
_DATA_START_COLUMN = 3

_KEY_HEADER = "MATCH KEY"
_STATUS_HEADER = "IN THE OTHER DRAFT?"

_HEADER_FILL = PatternFill("solid", start_color=cols.HEADER_FILL_RGB, end_color=cols.HEADER_FILL_RGB)
_HEADER_FONT = Font(name=cols.HEADER_FONT_NAME, size=cols.HEADER_FONT_SIZE, color=cols.HEADER_FONT_RGB)
_KEY_FILL = PatternFill("solid", start_color="FFF2E8F0", end_color="FFF2E8F0")

# Columns nobody's draft can be judged on: OPUS assigns them, or we write
# them blank on purpose. Leaving them out of a key is what stops two
# identical rows failing to match over a sequence number neither side
# controls.
_ASSIGNED_FIELDS = frozenset({
    "type", "cmdt_seq", "route_seq", "header_seq", "note_seq", "charge_seq",
}) | cols.FREETIME_UNFILED_FIELDS


@dataclass(frozen=True)
class SheetSpec:
    """How one sheet type is laid out, and what identifies a row on it."""

    row_fields: tuple[str, ...]
    headers: tuple[str, ...]
    key_fields: tuple[str, ...]


def _joined_headers(group: list[str], field: list[str | None]) -> tuple[str, ...]:
    """A single header row from OPUS's two, because a lookup range is
    easier to write against one row. The two labels are joined so it
    still reads like the sheet it came from."""
    out = []
    for group_label, field_label in zip(group, field):
        top = (group_label or "").replace("\n", " ").strip()
        bottom = (field_label or "").replace("\n", " ").strip()
        out.append(f"{top} {bottom}".strip() if bottom else top)
    return tuple(out)


def _flat_headers(header: list[str]) -> tuple[str, ...]:
    return tuple((h or "").replace("\n", " ").strip() for h in header)


def _key_fields(row_fields: tuple[str, ...]) -> tuple[str, ...]:
    """Everything on the row that either draft is actually responsible
    for - the default identity for sheets the audit has no stated
    convention for."""
    return tuple(f for f in row_fields if f not in _ASSIGNED_FIELDS)


# RATES uses the audit's own stated concat rather than the default: every
# location code, the terms, the transmodes, the whole Rate section, the
# Route Note. The rest fall back to "everything either side controls".
SHEET_SPECS: dict[str, SheetSpec] = {
    "RATES": SheetSpec(
        tuple(cols.RATES_ROW_FIELDS),
        _joined_headers(cols.RATES_HEADER_GROUP, cols.RATES_HEADER_FIELD),
        tuple(AUDIT_CONCAT_FIELDS),
    ),
    "RATES PORT-PORT": SheetSpec(
        tuple(cols.RATES_ROW_FIELDS),
        _joined_headers(cols.RATES_HEADER_GROUP, cols.RATES_HEADER_FIELD),
        tuple(AUDIT_CONCAT_FIELDS),
    ),
    "VERTICAL RATES": SheetSpec(
        tuple(cols.VERTICAL_RATES_ROW_FIELDS),
        _joined_headers(cols.VERTICAL_RATES_HEADER_GROUP, cols.VERTICAL_RATES_HEADER_FIELD),
        _key_fields(tuple(cols.VERTICAL_RATES_ROW_FIELDS)),
    ),
    "ARBS": SheetSpec(
        tuple(cols.ARBS_ROW_FIELDS),
        _flat_headers(cols.ARBS_HEADER),
        _key_fields(tuple(cols.ARBS_ROW_FIELDS)),
    ),
    "CMDT NOTE": SheetSpec(
        tuple(cols.CMDT_NOTE_ROW_FIELDS),
        _flat_headers(cols.CMDT_NOTE_HEADER),
        _key_fields(tuple(cols.CMDT_NOTE_ROW_FIELDS)),
    ),
    "SPECIAL NOTE": SheetSpec(
        tuple(cols.SPECIAL_NOTE_ROW_FIELDS),
        _flat_headers(cols.SPECIAL_NOTE_HEADER),
        _key_fields(tuple(cols.SPECIAL_NOTE_ROW_FIELDS)),
    ),
    "ROUTE NOTE": SheetSpec(
        tuple(cols.RN_ROW_FIELDS),
        _flat_headers(cols.RN_HEADER),
        _key_fields(tuple(cols.RN_ROW_FIELDS)),
    ),
    "FREETIME": SheetSpec(
        tuple(cols.FREETIME_ROW_FIELDS),
        _joined_headers(cols.FREETIME_HEADER_GROUP, cols.FREETIME_HEADER_FIELD),
        _key_fields(tuple(cols.FREETIME_ROW_FIELDS)),
    ),
}

# Sheets whose rows are not self-contained, so their keys legitimately
# repeat and a lookup lands on the first match. Called out in HOW TO USE
# rather than left for someone to discover.
_CONTINUATION_SHEETS = {"VERTICAL RATES", "CMDT NOTE", "SPECIAL NOTE"}


def _concat_formula(spec: SheetSpec, excel_row: int) -> str:
    """The match key as a live CONCATENATE over the row's own cells.

    Built with "&" rather than TEXTJOIN on purpose. TEXTJOIN postdates
    the xlsx format, so Excel only recognises it in a file when it is
    written as _xlfn.TEXTJOIN - a plain one evaluates to #NAME?. "&" has
    no such problem in any version.

    A blank cell concatenates as empty, so the separators still line up
    across rows. It also settles a difference building the string in
    Python would create: a rate held as 5800 on one sheet and 5800.0 on
    the other renders "5800" both times here, where str() gives "5800"
    and "5800.0" and the two rows never find each other.
    """
    refs = [
        f"{get_column_letter(_DATA_START_COLUMN + spec.row_fields.index(field))}{excel_row}"
        for field in spec.key_fields
    ]
    return "=" + '&"|"&'.join(refs)


def _sheet_name(base: str, side: str) -> str:
    name = f"{base} ({side})"
    return name if len(name) <= _MAX_SHEET_NAME else f"{base[: _MAX_SHEET_NAME - len(side) - 3]} ({side})"


def _write_side(
    wb: Workbook, spec: SheetSpec, title: str, rows: list[dict], other_title: str, other_label: str
) -> None:
    ws = wb.create_sheet(title[:_MAX_SHEET_NAME])
    ws.append([_KEY_HEADER, _STATUS_HEADER, *spec.headers])
    for cell in ws[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = cols.HEADER_ROW_HEIGHT

    quoted = other_title.replace("'", "''")
    for i, row in enumerate(rows, start=2):
        # _xlfn. is required, not decoration. XLOOKUP postdates the xlsx
        # format, and Excel only recognises a post-spec function in a FILE
        # when it carries that prefix - written plainly it evaluates to
        # #NAME?. Excel hides the prefix in the formula bar and re-adds it
        # on save, so a cell retyped by hand looks identical and works,
        # which is what made this read as a calculation problem rather
        # than a spelling one.
        lookup = (
            f"=_xlfn.XLOOKUP($A{i},'{quoted}'!$A:$A,'{quoted}'!$A:$A,\"NOT IN {other_label}\")"
        )
        ws.append([
            _concat_formula(spec, i), lookup,
            *[row.get(field) for field in spec.row_fields],
        ])
        ws.cell(row=i, column=1).fill = _KEY_FILL

    ws.freeze_panes = "C2"
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 22
    for idx in range(_DATA_START_COLUMN, len(spec.row_fields) + _DATA_START_COLUMN):
        ws.column_dimensions[get_column_letter(idx)].width = 14
    ws.auto_filter.ref = ws.dimensions


def build_side_by_side_workbook(pairs: list[tuple[str, str, list[dict], list[dict]]]) -> bytes:
    """pairs: (sheet type, sheet label, our rows, the reference's rows).

    Sheet type selects the layout and the match key - see SHEET_SPECS. An
    unknown type is skipped rather than written with the wrong columns.
    """
    wb = Workbook()
    wb.remove(wb.active)
    written = [p for p in pairs if p[0] in SHEET_SPECS]

    notes = wb.create_sheet("HOW TO USE")
    lines = [
        ["Both drafts, one workbook - reconcile them the way the audit already does."],
        [],
        ["MATCH KEY (column A) is the CONCATENATE, over this row's own cells."],
        ["IN THE OTHER DRAFT? (column B) is an XLOOKUP against the paired sheet:"],
        ["  it returns the key when the row is found, or NOT IN ... when it isn't."],
        [],
        ["Both are live formulas, not saved answers: edit a cell on either sheet and"],
        ["the key rebuilds and the lookup re-runs."],
        [],
        ["Sequence numbers are left OUT of every key - OPUS assigns them and neither"],
        ["draft controls them, so including them would fail every row."],
        [],
        ["On RATES the key is the audit's own concat:"],
        ["  " + " | ".join(AUDIT_CONCAT_FIELDS)],
        ["It INCLUDES the rate section, so a row whose rate differs does not match:"],
        ["  it reads NOT IN ... on both sheets, once as your row and once as theirs."],
        [],
        ["The on-screen comparison deliberately does the opposite - it matches WITHOUT"],
        ["the rates, so it can name the route and say which rate column differs. Use"],
        ["this workbook to find that a row disagrees; use the on-screen substance list"],
        ["to see what disagrees."],
    ]
    if any(p[0] in _CONTINUATION_SHEETS for p in written):
        lines += [
            [],
            ["One caveat on " + ", ".join(sorted(_CONTINUATION_SHEETS)) + ":"],
            ["  their rows are continuations - a VERTICAL RATES row carries only a container"],
            ["  size and a rate, a note's child rows carry only a charge code - so identical"],
            ["  continuation rows share a key and the lookup lands on the first one. Read"],
            ["  those sheets by block, not row by row."],
        ]
    for line in lines:
        notes.append(line)
    notes.column_dimensions["A"].width = 100
    notes["A1"].font = Font(bold=True)

    for sheet_type, label, ours, theirs in written:
        spec = SHEET_SPECS[sheet_type]
        ours_title = _sheet_name(label, "ours")
        ref_title = _sheet_name(label, "ref")
        _write_side(wb, spec, ours_title, ours, ref_title, "REFERENCE")
        _write_side(wb, spec, ref_title, theirs, ours_title, "YOUR DRAFT")

    # Every cell we write is a formula with no cached result, and Excel
    # trusts a saved workbook's stored results by default - so it opens
    # showing nothing until each cell is entered by hand. This is the flag
    # that tells it to calculate the whole book on load.
    wb.calculation.fullCalcOnLoad = True

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
