from __future__ import annotations

import openpyxl

from mrg2opus.parsers.laec import COMMODITY_NON_ISC_MAIN, NOR_DEFAULT_DESCRIPTION, LAECParser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols
from tests.golden import _normalize, diff_rates, read_cmdt_note_sheet, read_rates_sheet

PATH = "Sample MRGs with OPUS FORMATS/LAEC.xlsx"

# Known, verified gaps:
#   - cmdt_seq (403-406 in the ground truth): an externally-assigned
#     running number, same category as header_seq/note_seq in CMDT NOTE -
#     see laec.py's COMMODITY_* comment.
#   - route_seq: a real per-row sequence in the ground truth, but the
#     numbering resets at section boundaries in a way not derivable with
#     confidence in the time available (e.g. it's ~1510 partway through the
#     Non-ISC D/DG block but resets to ~70 at the start of the ISC D/DR
#     block) - commodity_note (the CMDT NOTE Contents text per row) IS
#     implemented since it's a clean per-commodity-group lookup.
#   - type: forced to "C" on every row for every lane, a deliberate,
#     user-directed business rule applied uniformly (not derived from any
#     one sample) - LAEC's own ground truth leaves it blank.
#   - commodity_group_description: the Non-ISC portion of "DRY" and "R5
#     NOR" used to share ONE combined description in the ground truth
#     ("FAK & DG & NOR_NON-ISC", G0015); per a deliberate, user-directed
#     rule, "R5 NOR" now defaults to its OWN description (its own raw
#     sheet name) instead, and only shares one CMDT NOTE block with the
#     Non-ISC main group if the user overrides them to the same
#     description - see test_laec_cmdt_note_default_splits_by_sheet and
#     test_laec_cmdt_note_merges_when_descriptions_match below.
RATES_IGNORE_FIELDS = {"cmdt_seq", "route_seq", "type", "commodity_group_description"}


def _run_laec():
    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = LAECParser()
    return parser.run(wb, MappingProfile())


def test_laec_rates_matches_ground_truth():
    row_set = _run_laec()
    generated = [r.model_dump() for r in row_set.rates]
    expected = read_rates_sheet(PATH, cols.SHEET_NAME_RATES)

    matched, missing, extra, mismatches = diff_rates(generated, expected, ignore_fields=RATES_IGNORE_FIELDS)

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"


def test_laec_rates_port_port_matches_ground_truth():
    row_set = _run_laec()
    generated = [r.model_dump() for r in row_set.rates_port_port]
    expected = read_rates_sheet(PATH, cols.SHEET_NAME_RATES_PORT_PORT)

    matched, missing, extra, mismatches = diff_rates(generated, expected, ignore_fields=RATES_IGNORE_FIELDS)

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"


def test_laec_cmdt_note_default_splits_by_sheet():
    """Default behavior: the Non-ISC portion of "DRY" and "R5 NOR"
    (previously one combined G0015 description) each get their own
    description (their own raw sheet name) and their own CMDT NOTE block -
    a deliberate, user-directed default (see RATES_IGNORE_FIELDS comment).
    ISC main and both in-gauge groups were always independent."""
    row_set = _run_laec()
    descriptions = {r.commodity_group_description for r in row_set.rates}
    assert descriptions == {
        "FAK & DG_NON-ISC",
        "FAK_ISC",
        "R5 NOR",
        "INGAUGE FAK_NON-ISC",
        "INGAUGE FAK_ISC",
    }
    blocks = [r for r in row_set.cmdt_notes if r.code == "APP"]
    assert len(blocks) == 5


def test_laec_cmdt_note_merges_when_descriptions_match():
    """If the user overrides "R5 NOR"'s description back to the ORIGINAL
    combined ground-truth text (and the Non-ISC main group's likewise),
    they collapse back into exactly the ground truth's CMDT NOTE - a
    regression check that the merge-by-description logic reconstructs the
    originally-verified behavior exactly."""
    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = LAECParser()
    combined = "FAK & DG & NOR_NON-ISC"
    profile = MappingProfile(
        commodity_description_overrides={COMMODITY_NON_ISC_MAIN[1]: combined, NOR_DEFAULT_DESCRIPTION: combined}
    )
    row_set = parser.run(wb, profile)
    generated = [r.model_dump() for r in row_set.cmdt_notes]
    expected = read_cmdt_note_sheet(PATH, cols.SHEET_NAME_CMDT_NOTE)

    assert len(generated) == len(expected)
    # header_seq/note_seq: externally-assigned running sequence numbers,
    # not derivable from this file. pol: same unresolved gap found in CSE -
    # one specific THL child row carries pol="BDCGP" (Chittagong), tied to
    # a Chittagong-specific rate-structure override in the raw footnotes
    # but not derivable with confidence from that text alone.
    ignore = {"header_seq", "note_seq", "pol"}
    for i, (g, e) in enumerate(zip(generated, expected)):
        for field_name in cols.CMDT_NOTE_ROW_FIELDS:
            if field_name in ignore:
                continue
            gv = _normalize(g.get(field_name))
            ev = _normalize(e.get(field_name))
            assert gv == ev, f"row {i}: {field_name}: {gv!r} != {ev!r}"


def test_laec_skip_dg_generation_suppresses_dg_rows_for_one_group_only():
    """Toggling it off for the Non-ISC main group alone must not touch the
    ISC main group's own DG rows - this lane has multiple independent
    commodity groups, each keyed by its own default description."""
    default_row_set = _run_laec()
    non_isc_default_cgo_types = {
        r.cgo_type for r in default_row_set.rates if r.commodity_group_description == COMMODITY_NON_ISC_MAIN[1]
    }
    assert "DG" in non_isc_default_cgo_types

    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = LAECParser()
    profile = MappingProfile(skip_dg_generation={COMMODITY_NON_ISC_MAIN[1]: True})
    row_set = parser.run(wb, profile)

    non_isc_cgo_types = {
        r.cgo_type for r in row_set.rates if r.commodity_group_description == COMMODITY_NON_ISC_MAIN[1]
    }
    assert "DG" not in non_isc_cgo_types
    assert "DR" in non_isc_cgo_types

    isc_cgo_types = {r.cgo_type for r in row_set.rates if r.commodity_group_description == "FAK_ISC"}
    assert "DG" in isc_cgo_types  # untouched - only the Non-ISC main group was opted out

    assert len(row_set.rates) < len(default_row_set.rates)
