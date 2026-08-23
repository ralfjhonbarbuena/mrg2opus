from __future__ import annotations

import openpyxl

from mrg2opus.parsers.lawc import COMMODITY_MAIN, NOR_DEFAULT_DESCRIPTION, REEFER_DEFAULT_DESCRIPTION, LAWCParser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols
from tests.golden import _normalize, diff_rates, read_cmdt_note_sheet, read_rates_sheet

PATH = "Sample MRGs with OPUS FORMATS/LAWC.xlsx"

# Known, verified gaps:
#   - cmdt_seq (1-4 in the ground truth): an externally-assigned running
#     number, same category as LAEC's cmdt_seq gap (see
#     test_parsers_laec.py). route_seq is excluded for the same reason.
#   - commodity_note on OPUS RATES: derived correctly per-group (matches
#     the OPUS CMDT NOTE sheet's own boilerplate text 1:1, see the passing
#     cmdt_note test below) but this ground truth sample's RATES-sheet
#     commodity_note column is desynced from its own commodity_group_code -
#     e.g. G0003 (ISC)'s ground-truth note text lists SEA's charge codes
#     (CDD/CSS/DOC/THL), not ISC's. Confirmed by checking that each group
#     code maps to exactly one distinct note text, but the text/group
#     pairing doesn't match any raw sheet's own validity+charge-code data.
#     Stale/inconsistent ground truth, not a derivable rule.
#   - type: forced to "C" on every row for every lane, a deliberate,
#     user-directed business rule applied uniformly (not derived from any
#     one sample) - LAWC's own ground truth leaves it blank.
#   - commodity_group_description: the main dry grid, "Reefer", and "LAWC
#     NOR" used to share ONE combined description in the ground truth
#     ("China_TWN_SIN_HKG Dry and DG & REEFER & NOR"); per a deliberate,
#     user-directed rule, each of these 3 raw sheets now defaults to its
#     OWN description (its own sheet name) instead, and only shares one
#     CMDT NOTE block with another sheet if the user overrides them to the
#     same description - see test_lawc_cmdt_note_default_splits_by_sheet
#     and test_lawc_cmdt_note_merges_when_descriptions_match below.
RATES_IGNORE_FIELDS = {"cmdt_seq", "route_seq", "commodity_note", "type", "commodity_group_description"}
# OPUS RATES PORT-PORT's commodity_note holds a small integer per row (e.g.
# 212, 564) instead of text - not the boilerplate blob at all, and not
# derivable from any raw source data (looks like an internal OPUS row/note
# ID with no equivalent in the MRG). Same treatment as the RATES gap above.
PORT_PORT_IGNORE_FIELDS = RATES_IGNORE_FIELDS


def _run_lawc():
    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = LAWCParser()
    return parser.run(wb, MappingProfile())


def test_lawc_rates_matches_ground_truth():
    row_set = _run_lawc()
    generated = [r.model_dump() for r in row_set.rates]
    expected = read_rates_sheet(PATH, cols.SHEET_NAME_RATES)

    matched, missing, extra, mismatches = diff_rates(generated, expected, ignore_fields=RATES_IGNORE_FIELDS)

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"


def test_lawc_rates_port_port_matches_ground_truth():
    row_set = _run_lawc()
    generated = [r.model_dump() for r in row_set.rates_port_port]
    expected = read_rates_sheet(PATH, cols.SHEET_NAME_RATES_PORT_PORT)

    matched, missing, extra, mismatches = diff_rates(generated, expected, ignore_fields=PORT_PORT_IGNORE_FIELDS)

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"


def test_lawc_cmdt_note_default_splits_by_sheet():
    """Default behavior: the main dry grid, "Reefer", and "LAWC NOR" each
    get their own description (their own raw sheet name) and their own
    CMDT NOTE block, instead of the ground truth's single combined block -
    a deliberate, user-directed default (see RATES_IGNORE_FIELDS comment).
    ISC/SEA/OOG were always independent and are unaffected."""
    row_set = _run_lawc()
    descriptions = {r.commodity_group_description for r in row_set.rates}
    assert descriptions == {
        "ISC_LK_BD_AE_PK Dry & DG",
        "S.E.A_JPN_SA_AU_NZ DRY AND DG",
        "China_TWN_SIN_HKG_KR Dry",
        "Reefer",
        "LAWC NOR",
        "OOG",
    }
    blocks = [r for r in row_set.cmdt_notes if r.code == "APP"]
    assert len(blocks) == 6

    # Reefer and NOR both borrow the main dry grid's charge codes/validity
    # (there's no independent ground truth to derive their own from), so
    # their note TEXT is naturally identical even as two separate blocks -
    # what matters is each is internally consistent and the block count
    # above reflects 6 truly separate groups, not that the text differs.
    reefer_rows = [r for r in row_set.rates if r.commodity_group_description == "Reefer"]
    nor_rows = [r for r in row_set.rates if r.commodity_group_description == "LAWC NOR"]
    assert reefer_rows and len({r.commodity_note for r in reefer_rows}) == 1
    assert nor_rows and len({r.commodity_note for r in nor_rows}) == 1


def test_lawc_cmdt_note_merges_when_descriptions_match():
    """If the user overrides the main dry grid's, Reefer's, and NOR's
    descriptions back to the ORIGINAL combined ground-truth text, they
    collapse back into exactly ONE CMDT NOTE block matching the ground
    truth byte-for-byte - a regression check that the merge-by-description
    logic reconstructs the originally-verified behavior exactly, not just
    something plausible."""
    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = LAWCParser()
    combined = "China_TWN_SIN_HKG Dry and DG & REEFER & NOR"
    profile = MappingProfile(
        commodity_description_overrides={
            COMMODITY_MAIN[1]: combined,
            REEFER_DEFAULT_DESCRIPTION: combined,
            NOR_DEFAULT_DESCRIPTION: combined,
        }
    )
    row_set = parser.run(wb, profile)
    generated = [r.model_dump() for r in row_set.cmdt_notes]
    expected = read_cmdt_note_sheet(PATH, cols.SHEET_NAME_CMDT_NOTE)

    assert len(generated) == len(expected)
    # header_seq/note_seq: externally-assigned running sequence numbers, not
    # derivable from this file. pol: same category of unresolved gap as
    # LAEC's (see test_parsers_laec.py) - the SEA group's CSS/THL/DOC/CDD
    # child rows carry pol="LKCMB" (Colombo) in the ground truth, a single
    # observed instance not confidently generalizable from raw footnote text.
    ignore = {"header_seq", "note_seq", "pol"}
    for i, (g, e) in enumerate(zip(generated, expected)):
        for field_name in cols.CMDT_NOTE_ROW_FIELDS:
            if field_name in ignore:
                continue
            gv = _normalize(g.get(field_name))
            ev = _normalize(e.get(field_name))
            assert gv == ev, f"row {i}: {field_name}: {gv!r} != {ev!r}"


def test_lawc_skip_dg_generation_suppresses_dg_rows_for_one_group_only():
    """Toggling it off for the main dry grid alone must not touch ISC/SEA's
    own DG rows - this lane has multiple independent commodity groups, each
    keyed by its own default description."""
    default_row_set = _run_lawc()
    main_default_cgo_types = {
        r.cgo_type for r in default_row_set.rates if r.commodity_group_description == COMMODITY_MAIN[1]
    }
    assert "DG" in main_default_cgo_types

    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = LAWCParser()
    profile = MappingProfile(skip_dg_generation={COMMODITY_MAIN[1]: True})
    row_set = parser.run(wb, profile)

    main_cgo_types = {r.cgo_type for r in row_set.rates if r.commodity_group_description == COMMODITY_MAIN[1]}
    assert "DG" not in main_cgo_types
    assert "DR" in main_cgo_types

    isc_cgo_types = {
        r.cgo_type for r in row_set.rates if r.commodity_group_description == "ISC_LK_BD_AE_PK Dry & DG"
    }
    assert "DG" in isc_cgo_types  # untouched - only the main dry grid was opted out

    assert len(row_set.rates) < len(default_row_set.rates)
