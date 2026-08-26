"""excel_io.writer output sheet naming + RN sheet - see
project-opus-note-sheet-taxonomy memory for why these real filing names
(RATES, ORIGIN ARBS, CMDT NOTE, SPECIAL NOTE, RN) differ from
schema.opus_columns's SHEET_NAME_* constants (those match the older
bundled-sample fixtures instead, see that module's own docstring)."""
from __future__ import annotations

import tempfile
from pathlib import Path

import openpyxl

from mrg2opus.excel_io.writer import write_opus_workbook_multi
from mrg2opus.schema.opus_rows import (
    ArbsRow,
    CmdtNoteRow,
    OpusRowSet,
    RatesPortPortRow,
    RatesRow,
    RouteNoteRow,
    SpecialNoteRow,
)

_RATES_KWARGS = dict(
    commodity_group_code="G0001", commodity_group_description="FAK",
    origin_code="CNSHA", origin_description="Shanghai",
    destination_code="USLAX", destination_description="Los Angeles",
    prefix="D", cgo_type="DR",
)


def _write_and_load(row_sets: dict[str, OpusRowSet]) -> openpyxl.Workbook:
    with tempfile.TemporaryDirectory() as tmp_dir:
        out_path = Path(tmp_dir) / "out.xlsx"
        write_opus_workbook_multi(row_sets, out_path)
        return openpyxl.load_workbook(out_path)


def test_writer_uses_real_filing_sheet_names():
    row_set = OpusRowSet(
        rates=[RatesRow(**_RATES_KWARGS)],
        rates_port_port=[RatesPortPortRow(**_RATES_KWARGS)],
        arbs=[ArbsRow(point="CNSHA")],
        cmdt_notes=[CmdtNoteRow(contents="note")],
        special_notes=[SpecialNoteRow(contents="note")],
        route_notes=[RouteNoteRow(contents="Rates are applicable for Vessel Service Lane: MX2")],
    )
    wb = _write_and_load({"": row_set})
    assert set(wb.sheetnames) == {"RATES", "RATES PORT-PORT", "ORIGIN ARBS", "CMDT NOTE", "SPECIAL NOTE", "RN"}


def test_writer_suffix_tags_real_filing_sheet_names():
    row_set = OpusRowSet(rates=[RatesRow(**_RATES_KWARGS)], route_notes=[RouteNoteRow(contents="note")])
    wb = _write_and_load({"TZDAR": row_set})
    assert set(wb.sheetnames) == {"RATES-TZDAR", "RN-TZDAR"}


def test_writer_omits_route_note_sheet_when_empty():
    row_set = OpusRowSet(rates=[RatesRow(**_RATES_KWARGS)])
    wb = _write_and_load({"": row_set})
    assert "RN" not in wb.sheetnames
