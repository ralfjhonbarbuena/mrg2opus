from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from mrg2opus.audit.compare import _normalize, diff_by_key, rates_row_key, read_cmdt_note_sheet, read_rates_sheet
from mrg2opus.parsers.auec import AUECParser, DEFAULT_MAIN_DESCRIPTION, DEFAULT_NZJ_DESCRIPTION
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols

REFERENCE_DIR = Path(__file__).resolve().parents[1] / "reference"

# 2 weekly pairs, both real ground truth (tracker status "Completed").
PAIRS = [
    (
        REFERENCE_DIR / "1_MRGs" / "37_AUS NEA to AUEC FAK" / "ONE AU MRG 2026_0815 to 2026_0831 - ex NEA to AUEC (07 August 2026).xlsx",
        REFERENCE_DIR / "2_OPUS" / "37_AUS NEA to AUEC FAK" / "ONE AU MRG 2026_0815 to 2026_0831 - ex NEA to AUEC (07 August 2026)_OPUS.xlsx",
    ),
    (
        REFERENCE_DIR / "1_MRGs" / "38_AUS NEA to AUEC FAK" / "ONE AU MRG 2026_0901 to 2026_0914 - ex NEA to AUEC (24 August 2026).xlsx",
        REFERENCE_DIR / "2_OPUS" / "38_AUS NEA to AUEC FAK" / "ONE AU MRG 2026_0901 to 2026_0914 - ex NEA to AUEC (24 August 2026)_opus.xlsx",
    ),
]

pytestmark = pytest.mark.skipif(
    any(not p.exists() for pair in PAIRS for p in pair),
    reason="reference/ ground-truth files not present in this checkout",
)

# Known, deliberate deviations from ground truth (same categories already
# documented for every other lane, see RATES_IGNORE_FIELDS_BY_LANE in
# audit/compare.py):
#   - type: forced to "C" on every row, a user-directed business rule.
#   - commodity_group_code/commodity_group_description/cmdt_seq: OPUS's
#     own global running sequence (G0005/G0003 in this ground truth) is
#     not reproducible from the raw MRG alone - user-customizable per
#     filing, same gap already documented for every other lane.
#   - route_seq: single running counter per commodity group across every
#     destination/container-type block - verified correct in shape (see
#     test_auec_route_seq_is_one_continuous_counter_per_group below) but
#     not pinned to the exact literal ground-truth numbers, same
#     "OPUS renumbers on import" placeholder status as every other lane.
#   - commodity_note: exact text verified separately via the CMDT NOTE
#     sheet comparison below.
RATES_IGNORE_FIELDS = {
    "type", "commodity_group_code", "commodity_group_description", "cmdt_seq", "commodity_note", "route_seq",
}


def _run_auec(raw_path: Path):
    wb = openpyxl.load_workbook(raw_path, data_only=True)
    return AUECParser().run(wb, MappingProfile())


@pytest.mark.parametrize("raw_path,opus_path", PAIRS)
def test_auec_rates_matches_ground_truth(raw_path, opus_path):
    row_set = _run_auec(raw_path)
    generated = [r.model_dump() for r in row_set.rates]

    ref_wb = openpyxl.load_workbook(opus_path, data_only=True, read_only=True)
    expected = read_rates_sheet(ref_wb, "RATES")

    result = diff_by_key(generated, expected, key_fn=rates_row_key, fields=cols.RATES_ROW_FIELDS, ignore_fields=RATES_IGNORE_FIELDS)
    assert not result.missing, f"missing {len(result.missing)} expected rows, e.g. {list(result.missing)[:5]}"
    assert not result.extra, f"{len(result.extra)} unexpected generated rows, e.g. {list(result.extra)[:5]}"
    assert not result.field_mismatches, f"{len(result.field_mismatches)} field mismatches, e.g. {result.field_mismatches[:10]}"


@pytest.mark.parametrize("raw_path,opus_path", PAIRS)
def test_auec_cmdt_note_matches_ground_truth(raw_path, opus_path):
    """Covers the lane-specific POL scoping (ISL only for Taiwan-origin
    shipments; EFS filed as two separate rows scoped to Hong Kong and
    Korea respectively) - see auec.py's INCLUDED_CHARGES module comment."""
    row_set = _run_auec(raw_path)
    generated = [r.model_dump() for r in row_set.cmdt_notes]

    ref_wb = openpyxl.load_workbook(opus_path, data_only=True, read_only=True)
    expected = read_cmdt_note_sheet(ref_wb, "SUR")

    assert len(generated) == len(expected)
    ignore = {"header_seq", "note_seq"}
    for i, (g, e) in enumerate(zip(generated, expected)):
        for field_name in cols.CMDT_NOTE_ROW_FIELDS:
            if field_name in ignore:
                continue
            if field_name in ("application_effective", "application_expires") and g.get("code") != "APP":
                continue
            gv, ev = _normalize(g.get(field_name)), _normalize(e.get(field_name))
            assert gv == ev, f"row {i} {field_name}: {gv!r} != {ev!r}"


def test_auec_route_seq_is_one_continuous_counter_per_group():
    """Confirmed against ground truth: Route Seq. is a single running
    counter across the WHOLE commodity group (every destination and
    container-type block), NOT resetting per destination or per
    cgo_type/prefix - restarting at 1 only for the next commodity group
    (main vs NZJ)."""
    raw_path, _ = PAIRS[0]
    row_set = _run_auec(raw_path)

    for description in (DEFAULT_MAIN_DESCRIPTION, DEFAULT_NZJ_DESCRIPTION):
        group_rows = [r for r in row_set.rates if r.commodity_group_description == description]
        assert [r.route_seq for r in group_rows] == list(range(1, len(group_rows) + 1))


def test_auec_excluded_charge_codes_drops_baf_end_to_end():
    """MappingProfile.excluded_charge_codes wired through: excluding EFS
    drops BOTH its POL-scoped rows (Hong Kong and Korea), not just one."""
    raw_path, _ = PAIRS[0]
    wb = openpyxl.load_workbook(raw_path, data_only=True)
    row_set = AUECParser().run(wb, MappingProfile(excluded_charge_codes=["EFS"]))

    codes = [n.code for n in row_set.cmdt_notes]
    assert "EFS" not in codes
    assert "ISL" in codes  # untouched - only EFS was excluded


def test_auec_skip_dg_generation_suppresses_dg_rows_for_main_group_only():
    """skip_dg_generation is keyed per commodity group's default
    description - suppressing the main group's DG rows must not affect
    the separate NZJ group."""
    raw_path, _ = PAIRS[0]
    default_row_set = _run_auec(raw_path)
    wb = openpyxl.load_workbook(raw_path, data_only=True)
    row_set = AUECParser().run(wb, MappingProfile(skip_dg_generation={DEFAULT_MAIN_DESCRIPTION: True}))

    default_main_cgo = {r.cgo_type for r in default_row_set.rates if r.commodity_group_description == DEFAULT_MAIN_DESCRIPTION}
    assert "DG" in default_main_cgo

    main_cgo = {r.cgo_type for r in row_set.rates if r.commodity_group_description == DEFAULT_MAIN_DESCRIPTION}
    assert "DG" not in main_cgo
    assert "DR" in main_cgo

    nzj_cgo = {r.cgo_type for r in row_set.rates if r.commodity_group_description == DEFAULT_NZJ_DESCRIPTION}
    assert "DG" in nzj_cgo  # NZJ group is unaffected
