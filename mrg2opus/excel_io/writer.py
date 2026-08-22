"""Write an OpusRowSet out to an OPUS-format Excel workbook, reproducing the
exact 2-row header on RATES sheets and the fill-down parent/child semantics
on CMDT NOTE / SPECIAL NOTE sheets.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from mrg2opus.schema import opus_columns as cols
from mrg2opus.schema.opus_rows import ArbsRow, CmdtNoteRow, OpusRowSet, RatesRow, SpecialNoteRow


def _write_rates_sheet(wb: Workbook, sheet_name: str, rows: list[RatesRow]) -> None:
    ws: Worksheet = wb.create_sheet(sheet_name)
    ws.append(cols.RATES_HEADER_GROUP)
    ws.append([f or "" for f in cols.RATES_HEADER_FIELD])
    for row in rows:
        data = row.model_dump()
        ws.append([data[field_name] for field_name in cols.RATES_ROW_FIELDS])


def _write_arbs_sheet(wb: Workbook, sheet_name: str, rows: list[ArbsRow]) -> None:
    ws: Worksheet = wb.create_sheet(sheet_name)
    ws.append(cols.ARBS_HEADER)
    for row in rows:
        data = row.model_dump()
        ws.append([data[field_name] for field_name in cols.ARBS_ROW_FIELDS])


def _write_cmdt_note_sheet(wb: Workbook, sheet_name: str, rows: list[CmdtNoteRow]) -> None:
    ws: Worksheet = wb.create_sheet(sheet_name)
    ws.append(cols.CMDT_NOTE_HEADER)
    for row in rows:
        data = row.model_dump()
        ws.append([data[field_name] for field_name in cols.CMDT_NOTE_ROW_FIELDS])


def _write_special_note_sheet(wb: Workbook, sheet_name: str, rows: list[SpecialNoteRow]) -> None:
    # Same field set as CMDT NOTE, different column order - see
    # schema/opus_columns.py's SPECIAL_NOTE_HEADER comment.
    ws: Worksheet = wb.create_sheet(sheet_name)
    ws.append(cols.SPECIAL_NOTE_HEADER)
    for row in rows:
        data = row.model_dump()
        ws.append([data[field_name] for field_name in cols.SPECIAL_NOTE_ROW_FIELDS])


def _sheet_names_for_suffix(suffix: str) -> dict[str, str]:
    """Naming convention confirmed against EAF.xlsx's sub-lane sheets: the
    suffix is appended directly to the base sheet name with a hyphen (e.g.
    'OPUS RATES-TZDAR', 'OPUS RATES-TZDAR PORT-PORT', 'OPUS CMDT NOTE-TZDAR'),
    or the plain base name when there's no sub-lane (suffix "")."""
    tag = f"-{suffix}" if suffix else ""
    return {
        "rates": f"OPUS RATES{tag}",
        "rates_port_port": f"OPUS RATES{tag} PORT-PORT",
        "arbs": f"OPUS ARBS{tag}",
        "cmdt_notes": f"OPUS CMDT NOTE{tag}",
        "special_notes": f"OPUS SPECIAL NOTE{tag}",
    }


def _write_row_set(wb: Workbook, row_set: OpusRowSet, names: dict[str, str]) -> None:
    if row_set.rates:
        _write_rates_sheet(wb, names["rates"], row_set.rates)
    if row_set.rates_port_port:
        _write_rates_sheet(wb, names["rates_port_port"], row_set.rates_port_port)
    if row_set.arbs:
        _write_arbs_sheet(wb, names["arbs"], row_set.arbs)
    if row_set.cmdt_notes:
        _write_cmdt_note_sheet(wb, names["cmdt_notes"], row_set.cmdt_notes)
    if row_set.special_notes:
        _write_special_note_sheet(wb, names["special_notes"], row_set.special_notes)


def write_opus_workbook(
    row_set: OpusRowSet,
    out_path: Path | str,
    sheet_names: dict[str, str] | None = None,
) -> None:
    """sheet_names lets a lane override default sheet names without changing
    writer logic. For lanes with sub-lanes sharing one workbook (e.g. EAF's
    TZDAR/KEMBA), use write_opus_workbook_multi instead."""
    names = _sheet_names_for_suffix("")
    if sheet_names:
        names.update(sheet_names)

    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet
    _write_row_set(wb, row_set, names)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def write_opus_workbook_multi(row_sets: dict[str, OpusRowSet], out_path: Path | str) -> None:
    """Write several sub-lane OpusRowSets into ONE workbook, each under its
    own suffixed sheet names (e.g. {"TZDAR": ..., "KEMBA": ...} ->
    'OPUS RATES-TZDAR', 'OPUS RATES-KEMBA', ...). Pass {"": row_set} for a
    single-output lane to reuse this same path."""
    wb = Workbook()
    wb.remove(wb.active)
    for suffix, row_set in row_sets.items():
        _write_row_set(wb, row_set, _sheet_names_for_suffix(suffix))

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
