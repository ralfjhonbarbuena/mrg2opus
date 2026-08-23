from __future__ import annotations

import openpyxl
import pytest

from mrg2opus.parsers.eaf import DEFAULT_COMMODITY_DESCRIPTION, EAFParser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols
from tests.golden import diff_rates, read_cmdt_note_sheet, read_rates_sheet, _normalize_cmdt_value

PATH = "Sample MRGs with OPUS FORMATS/EAF.xlsx"
SUBLANES = ["TZDAR", "KEMBA"]

# Known, deliberate deviation from ground truth: `type` is forced to "C" on
# every row for every lane (a user-directed business rule, not derived from
# any one sample) - EAF's own ground truth leaves it blank, so it's excluded
# from the golden-diff rather than reproduced.
TYPE_OVERRIDE_IGNORE = {"type"}


def _run_eaf():
    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = EAFParser()
    return parser.run_multi(wb, MappingProfile())


def _extract_single_sheet_workbook(sheet_name: str) -> openpyxl.Workbook:
    """Real EAF MRGs arrive as two SEPARATE files, one per sub-lane (TZDAR,
    KEMBA) - EAF.xlsx only bundles both raw sheets into one workbook as a
    sample-data convenience. This rebuilds a single-sheet workbook to
    confirm the parser behaves the same way on a genuinely standalone file
    as it does on the combined sample."""
    src_wb = openpyxl.load_workbook(PATH, data_only=True)
    src_ws = src_wb[sheet_name]
    dst_wb = openpyxl.Workbook()
    dst_wb.remove(dst_wb.active)
    dst_ws = dst_wb.create_sheet(sheet_name)
    for row in src_ws.iter_rows():
        for cell in row:
            dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
    for merged in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged))
    return dst_wb


@pytest.mark.parametrize("sublane", SUBLANES)
def test_eaf_standalone_single_lane_file(sublane):
    """A standalone single-sheet file must classify and parse the same as
    when it's one of two sheets bundled in the combined sample - and
    run_multi() must return exactly that one sub-lane, not both."""
    wb = _extract_single_sheet_workbook(f"EAF {sublane}")
    parser = EAFParser()

    row_sets = parser.run_multi(wb, MappingProfile())
    assert set(row_sets.keys()) == {sublane}

    combined_row_sets = _run_eaf()
    standalone = row_sets[sublane]
    combined = combined_row_sets[sublane]
    assert [r.model_dump() for r in standalone.rates] == [r.model_dump() for r in combined.rates]
    assert [r.model_dump() for r in standalone.cmdt_notes] == [r.model_dump() for r in combined.cmdt_notes]


@pytest.mark.parametrize("sublane", SUBLANES)
def test_eaf_rates_matches_ground_truth(sublane):
    row_sets = _run_eaf()
    generated = [r.model_dump() for r in row_sets[sublane].rates]
    expected = read_rates_sheet(PATH, f"OPUS RATES-{sublane}")

    matched, missing, extra, mismatches = diff_rates(generated, expected, ignore_fields=TYPE_OVERRIDE_IGNORE)

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"
    assert matched == len(expected) == len(generated)


@pytest.mark.parametrize("sublane", SUBLANES)
def test_eaf_rates_port_port_matches_ground_truth(sublane):
    row_sets = _run_eaf()
    generated = [r.model_dump() for r in row_sets[sublane].rates_port_port]
    expected = read_rates_sheet(PATH, f"OPUS RATES-{sublane} PORT-PORT")

    # TZDAR's ground truth fills route_seq/cmdt_seq/commodity_note on every
    # PORT-PORT row (cross-referencing the CMDT NOTE); KEMBA's PORT-PORT
    # sheet - same workbook, same lane - leaves them blank. Since the two
    # sub-lanes contradict each other, the parser deliberately doesn't
    # generate this (see eaf.py comment) - so those 3 fields are excluded
    # here rather than asserting a convention the ground truth itself
    # doesn't consistently follow.
    ignore = {"route_seq", "cmdt_seq", "commodity_note"} | TYPE_OVERRIDE_IGNORE
    matched, missing, extra, mismatches = diff_rates(generated, expected, ignore_fields=ignore)

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"
    assert matched == len(expected) == len(generated)


@pytest.mark.parametrize("sublane", SUBLANES)
def test_eaf_cmdt_note_matches_ground_truth(sublane):
    row_sets = _run_eaf()
    generated = [r.model_dump() for r in row_sets[sublane].cmdt_notes]
    expected = read_cmdt_note_sheet(PATH, f"OPUS CMDT NOTE-{sublane}")

    assert len(generated) == len(expected)
    ignore = {"header_seq", "note_seq"}  # externally-assigned running sequence numbers, not derivable from this file
    for g, e in zip(generated, expected):
        for field_name in cols.CMDT_NOTE_ROW_FIELDS:
            if field_name in ignore:
                continue
            gv = _normalize_cmdt_value(g.get(field_name))
            ev = _normalize_cmdt_value(e.get(field_name))
            assert gv == ev, f"[{sublane}] {field_name}: {gv!r} != {ev!r}"


def test_eaf_skip_dg_generation_suppresses_dg_rows_for_both_sublanes():
    """EAF's two sub-lanes share one default description ("FAK") - same as
    every other MappingProfile override for this lane - so this toggle
    necessarily affects both TZDAR and KEMBA together, not independently."""
    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = EAFParser()
    default_row_sets = parser.run_multi(wb, MappingProfile())

    profile = MappingProfile(skip_dg_generation={DEFAULT_COMMODITY_DESCRIPTION: True})
    row_sets = parser.run_multi(wb, profile)

    for sublane in SUBLANES:
        default_cgo_types = {r.cgo_type for r in default_row_sets[sublane].rates}
        assert "DG" in default_cgo_types

        cgo_types = {r.cgo_type for r in row_sets[sublane].rates}
        assert "DG" not in cgo_types
        assert "DR" in cgo_types
        assert len(row_sets[sublane].rates) < len(default_row_sets[sublane].rates)
