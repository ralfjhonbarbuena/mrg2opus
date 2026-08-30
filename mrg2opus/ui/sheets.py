"""What sheets an export will actually contain - the one definition the
whole wizard asks.

Step 2's row counts, Step 3's skip checkboxes, Step 4's summary and the
skip-to-field translation each used to carry their own hardcoded list.
All four listed the same 5 of the 8 sheet types, so ROUTE NOTE, VERTICAL
RATES and FREETIME were invisible everywhere and impossible to skip -
57,904 rows across the reference lanes that the UI never mentioned, and
for several lanes (WAF, LAWC, the TAD trades) more rows hidden than
shown. They also labelled sheets "OPUS RATES"/"OPUS ARBS", names that
appear in no output workbook.

Both problems came from the list being copied rather than derived, so
this derives it: field order from OpusRowSet, sheet names from the
writer's own resolver, counts from the row sets in hand.
"""
from __future__ import annotations

from dataclasses import dataclass

from mrg2opus.excel_io.writer import resolve_sheet_names
from mrg2opus.schema.opus_rows import OpusRowSet

# Every OpusRowSet field the writer can emit a sheet for, in the order the
# workbook lays them out.
OUTPUT_SHEET_FIELDS = [
    "rates",
    "rates_port_port",
    "arbs",
    "cmdt_notes",
    "special_notes",
    "route_notes",
    "vertical_rates",
    "freetime",
]


@dataclass(frozen=True)
class OutputSheet:
    scope: str          # run_multi() sub-lane key ("" for a single-output lane)
    field: str          # OpusRowSet attribute
    name: str           # the sheet name the workbook will actually carry
    rows: int

    @property
    def scope_label(self) -> str:
        return self.scope or "(default)"


def output_sheets(row_sets: dict[str, OpusRowSet], parser_cls=None) -> list[OutputSheet]:
    """Every sheet the current parse would write, in workbook order.

    Empty sheets are left out because the writer skips them too - what's
    listed here is exactly what lands in the file.
    """
    overrides = getattr(parser_cls, "SHEET_NAME_OVERRIDES", None)
    scoped = getattr(parser_cls, "SCOPED_SHEET_NAME_OVERRIDES", None)

    sheets: list[OutputSheet] = []
    for scope, row_set in row_sets.items():
        names = resolve_sheet_names(scope, overrides, scoped)
        for field in OUTPUT_SHEET_FIELDS:
            rows = getattr(row_set, field, None) or []
            if rows:
                sheets.append(OutputSheet(scope=scope, field=field, name=names[field], rows=len(rows)))
    return sheets


def apply_skips(row_sets: dict[str, OpusRowSet], skip_output_sheets: dict[str, bool], parser_cls=None):
    """Blank out whichever OpusRowSet fields the user unticked. Keyed by
    the real sheet name, the same string output_sheets() showed them."""
    if not skip_output_sheets or not any(skip_output_sheets.values()):
        return row_sets

    skipped_fields: dict[str, set[str]] = {}
    for sheet in output_sheets(row_sets, parser_cls):
        if skip_output_sheets.get(sheet.name):
            skipped_fields.setdefault(sheet.scope, set()).add(sheet.field)

    if not skipped_fields:
        return row_sets

    filtered: dict[str, OpusRowSet] = {}
    for scope, row_set in row_sets.items():
        fields = skipped_fields.get(scope)
        if not fields:
            filtered[scope] = row_set
            continue
        data = row_set.model_dump()
        for field in fields:
            data[field] = []
        filtered[scope] = OpusRowSet.model_validate(data)
    return filtered
