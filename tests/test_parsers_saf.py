from __future__ import annotations

import openpyxl

from mrg2opus.parsers.saf import DEFAULT_COMMODITY_DESCRIPTION, SAFParser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols
from tests.golden import _normalize_cmdt_value, diff_rates, read_cmdt_note_sheet, read_rates_sheet


def _run_saf(profile: MappingProfile | None = None):
    path = "Sample MRGs with OPUS FORMATS/SAF.xlsx"
    wb = openpyxl.load_workbook(path, data_only=False)
    parser = SAFParser()
    row_set = parser.run(wb, profile or MappingProfile())
    return path, row_set


def test_saf_rates_matches_ground_truth():
    path, row_set = _run_saf()
    generated = [r.model_dump() for r in row_set.rates]
    expected = read_rates_sheet(path, cols.SHEET_NAME_RATES)

    matched, missing, extra, mismatches = diff_rates(generated, expected)

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"
    assert matched == len(expected) == len(generated)


def test_saf_rates_port_port_matches_ground_truth():
    path, row_set = _run_saf()
    generated = [r.model_dump() for r in row_set.rates_port_port]
    expected = read_rates_sheet(path, cols.SHEET_NAME_RATES_PORT_PORT)

    matched, missing, extra, mismatches = diff_rates(generated, expected)

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"
    assert matched == len(expected) == len(generated)


def test_saf_cmdt_note_matches_ground_truth():
    path, row_set = _run_saf()
    generated = [r.model_dump() for r in row_set.cmdt_notes]
    expected = read_cmdt_note_sheet(path, cols.SHEET_NAME_CMDT_NOTE)

    assert len(generated) == len(expected)
    ignore = {"header_seq", "note_seq"}  # externally-assigned running sequence numbers, not derivable from this file
    for g, e in zip(generated, expected):
        for field_name in cols.CMDT_NOTE_ROW_FIELDS:
            if field_name in ignore:
                continue
            gv = _normalize_cmdt_value(g.get(field_name))
            ev = _normalize_cmdt_value(e.get(field_name))
            assert gv == ev, f"{field_name}: {gv!r} != {ev!r}"


def test_saf_skip_dg_generation_suppresses_dg_rows_only():
    _, default_row_set = _run_saf()
    default_cgo_types = {r.cgo_type for r in default_row_set.rates}
    assert "DG" in default_cgo_types  # confirms the default (unset) case still generates DG, as before

    profile = MappingProfile(skip_dg_generation={DEFAULT_COMMODITY_DESCRIPTION: True})
    _, row_set = _run_saf(profile)
    cgo_types = {r.cgo_type for r in row_set.rates}
    assert "DG" not in cgo_types
    assert "DR" in cgo_types  # the base Dry rows are untouched
    assert len(row_set.rates) < len(default_row_set.rates)
