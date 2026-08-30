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
    VerticalRatesRow,
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


def _vr(cmdt_seq, n):
    """n vertical rows for one commodity group, in the shape
    explode_to_vertical_rates() produces: the header fields (cmdt_seq
    among them) only on the first row, blank on the rest."""
    return [VerticalRatesRow(cmdt_seq=cmdt_seq if i == 0 else None, per="D2", rate=100) for i in range(n)]


def test_vertical_rates_written_one_numbered_sheet_per_cmdt_seq():
    """Split by commodity group so a large lane stays under the OPUS
    upload limit, numbered by the group's own CMDT Seq."""
    row_set = OpusRowSet(vertical_rates=[*_vr(1, 3), *_vr(2, 2)])
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "o.xlsx"
        write_opus_workbook_multi({"": row_set}, out)
        wb = openpyxl.load_workbook(out)
        assert [s for s in wb.sheetnames if "VERTICAL" in s] == ["VERTICAL RATES 1", "VERTICAL RATES 2"]
        assert wb["VERTICAL RATES 1"].max_row == 5  # 2 header rows + 3
        assert wb["VERTICAL RATES 2"].max_row == 4  # 2 header rows + 2


def test_vertical_rates_non_contiguous_group_stays_one_sheet():
    """Regression: LAWC files its NOR rows under the SEA group, so one
    commodity group's rows arrive in two separate runs. Grouping on
    adjacency produced two sheets with the same name, and openpyxl
    silently renamed the second - splitting the group across two
    misleadingly-named sheets."""
    row_set = OpusRowSet(vertical_rates=[*_vr(1, 2), *_vr(2, 2), *_vr(1, 3)])
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "o.xlsx"
        write_opus_workbook_multi({"": row_set}, out)
        wb = openpyxl.load_workbook(out)
        assert [s for s in wb.sheetnames if "VERTICAL" in s] == ["VERTICAL RATES 1", "VERTICAL RATES 2"]
        assert wb["VERTICAL RATES 1"].max_row == 7  # 2 header rows + 2 + 3, merged
