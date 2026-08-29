from __future__ import annotations

import re
from pathlib import Path

import openpyxl
import pytest

from mrg2opus.audit.compare import _normalize, diff_by_key, rates_row_key, read_cmdt_note_sheet, read_rates_sheet
from mrg2opus.parsers.nz1_sea import DEFAULT_DESCRIPTION, Nz1SeaParser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols

REFERENCE_DIR = Path(__file__).resolve().parents[1] / "reference"

# 2 weekly pairs, both real ground truth (tracker status "Completed"). The
# two reference weeks use DIFFERENT sheet names for the same content
# ("rates"/"sur"/"rnt" vs "RATES"/"SURCHARGE"/"RNT") - not just a casing
# difference (the CMDT NOTE sheet's own name differs, "sur" vs "SURCHARGE"),
# so each pair carries its own sheet names rather than a single hardcoded
# name shared across both, unlike every other lane's test file so far.
PAIRS = [
    (
        REFERENCE_DIR / "1_MRGs" / "41_NZ1 SEA to NZBP FAK" / "ONE SEA to NZ MRG 20260815 to 20260831 (7 Aug 2026).xlsx",
        REFERENCE_DIR / "2_OPUS" / "41_NZ1 SEA to NZBP FAK" / "ONE SEA to NZ MRG 20260815 to 20260831 (7 Aug 2026)_opus.xlsx",
        "rates", "sur",
    ),
    (
        REFERENCE_DIR / "1_MRGs" / "42_NZ1 SEA to NZBP FAK" / "ONE SEA to NZ MRG 20260901 to 20260914 (24 AUG 2026).xlsx",
        REFERENCE_DIR / "2_OPUS" / "42_NZ1 SEA to NZBP FAK" / "ONE SEA to NZ MRG 20260901 to 20260914 (24 AUG 2026)_opus.xlsx",
        "RATES", "SURCHARGE",
    ),
]

pytestmark = pytest.mark.skipif(
    any(not p.exists() for pair in PAIRS for p in (pair[0], pair[1])),
    reason="reference/ ground-truth files not present in this checkout",
)

# Same accepted-gap categories as every other lane (see RATES_IGNORE_FIELDS_BY_LANE
# in audit/compare.py) - OPUS's own running sequence and user-customizable group
# identity aren't derivable from the raw MRG alone. route_seq is verified
# separately below (test_nz1_sea_route_seq_is_one_continuous_counter) since this
# parser's block ordering (DR, RF, RAD, then DG) reproduces the exact literal
# ground-truth numbers, but the project convention still treats it as a
# placeholder OPUS renumbers on import, not a pinned field in the main diff.
# route_note: deliberate scope decision for this pass - ground truth's own
# ROUTE NOTE sheet (Yangon's "Freight Collect" note, the 6 India ICD rows'
# own footnote code lists) is confirmed real (not a data-entry error) but
# intentionally not reproduced here; see nz1_sea.py's module docstring.
RATES_IGNORE_FIELDS = {
    "type", "commodity_group_code", "commodity_group_description", "cmdt_seq", "commodity_note", "route_seq",
    "route_note",
}


def _run(raw_path: Path) -> "OpusRowSet":  # noqa: F821
    wb = openpyxl.load_workbook(raw_path, data_only=True)
    return Nz1SeaParser().run(wb, MappingProfile())


@pytest.mark.parametrize("raw_path,opus_path,rates_sheet,sur_sheet", PAIRS)
def test_nz1_sea_rates_matches_ground_truth(raw_path, opus_path, rates_sheet, sur_sheet):
    row_set = _run(raw_path)
    generated = [r.model_dump() for r in row_set.rates]

    ref_wb = openpyxl.load_workbook(opus_path, data_only=True, read_only=True)
    expected = read_rates_sheet(ref_wb, rates_sheet)

    result = diff_by_key(generated, expected, key_fn=rates_row_key, fields=cols.RATES_ROW_FIELDS, ignore_fields=RATES_IGNORE_FIELDS)
    assert not result.missing, f"missing {len(result.missing)} expected rows, e.g. {list(result.missing)[:5]}"
    assert not result.extra, f"{len(result.extra)} unexpected generated rows, e.g. {list(result.extra)[:5]}"
    assert not result.field_mismatches, f"{len(result.field_mismatches)} field mismatches, e.g. {result.field_mismatches[:10]}"


def test_nz1_sea_cmdt_note_matches_ground_truth_week2():
    """Week 2's ground truth confirms this parser's fully-alphabetized
    "inclusive of" sentence text exactly (DOC, EFS, ISL, OBS, THL). Week 1
    diverges in sentence word order only (EFS, OBS, THL, DOC, ISL) - a
    confirmed, non-derivable per-filing text-order quirk (same category as
    west_asia_waf's own week1/week2 divergence) - and is checked separately,
    by content only, below."""
    raw_path, opus_path, _, sur_sheet = PAIRS[1]
    row_set = _run(raw_path)
    generated = [n.model_dump() for n in row_set.cmdt_notes]

    ref_wb = openpyxl.load_workbook(opus_path, data_only=True, read_only=True)
    expected = read_cmdt_note_sheet(ref_wb, sur_sheet)

    assert len(generated) == len(expected)
    ignore = {"header_seq", "note_seq"}
    for i, (g, e) in enumerate(zip(generated, expected)):
        for field_name in cols.CMDT_NOTE_ROW_FIELDS:
            if field_name in ignore:
                continue
            gv, ev = _normalize(g.get(field_name)), _normalize(e.get(field_name))
            assert gv == ev, f"row {i} {field_name}: {gv!r} != {ev!r}"


def test_nz1_sea_cmdt_note_matches_ground_truth_week1_by_content():
    raw_path, opus_path, _, sur_sheet = PAIRS[0]
    row_set = _run(raw_path)
    generated = [n.model_dump() for n in row_set.cmdt_notes]

    ref_wb = openpyxl.load_workbook(opus_path, data_only=True, read_only=True)
    expected = read_cmdt_note_sheet(ref_wb, sur_sheet)

    ignore = {"header_seq", "note_seq", "contents", "charge_seq"}

    def key(row):
        return tuple(_normalize(row.get(f)) for f in cols.CMDT_NOTE_ROW_FIELDS if f not in ignore)

    assert len(generated) == len(expected)
    assert {key(g) for g in generated} == {key(e) for e in expected}
    # Only the sentence's own word order differs, not its set of codes
    # ("...CHARGE(EFS)" has no space before the code paren; "DOC FEE
    # (ORIGIN)(DOC)" does before its non-code "(ORIGIN)" annotation).
    code_re = re.compile(r"(?<!\s)\(([A-Z]+)\)")
    assert sorted(code_re.findall(generated[0]["contents"])) == sorted(code_re.findall(expected[0]["contents"]))


def test_nz1_sea_route_seq_is_one_continuous_counter():
    """Confirmed against ground truth: Route Seq. is a single running
    counter across the whole block order (DR, then RF, then RAD, then the
    DG duplicate block), matching the exact literal ground-truth numbers
    (e.g. Yangon: DR=35, RF=79, DG=133; ICD cluster rows: DR=52-57,
    DG=150-155)."""
    raw_path, _, _, _ = PAIRS[0]
    row_set = _run(raw_path)
    assert [r.route_seq for r in row_set.rates] == list(range(1, len(row_set.rates) + 1))


def test_nz1_sea_dg_duplicate_at_same_rate():
    raw_path, _, _, _ = PAIRS[0]
    row_set = _run(raw_path)
    dr_rows = [r for r in row_set.rates if r.cgo_type == "DR" and r.prefix == "D"]
    dg_rows = [r for r in row_set.rates if r.cgo_type == "DG"]
    assert len(dr_rows) == len(dg_rows) == 61
    for dr_row, dg_row in zip(dr_rows, dg_rows):
        assert dr_row.origin_code == dg_row.origin_code
        assert dr_row.rate_20 == dg_row.rate_20
        assert dr_row.rate_40 == dg_row.rate_40
        assert dr_row.rate_40hc == dg_row.rate_40hc


def test_nz1_sea_icd_clusters_resolve_via_remark_via_text():
    """The 6 "India ICD" origin rows (ICD Delhi, Ahmedabad ICD, Ludhiana
    Cluster, Indore Cluster, Faridabad Cluster, Moradabad Cluster) don't
    exist in the Location Bank by their own names - confirmed against
    ground truth, they resolve via their own Remark cell's "Via <port(s)>"
    text instead, and are NOT deduped/merged even though 5 of the 6 share
    an identical resolved origin_code."""
    raw_path, _, _, _ = PAIRS[0]
    row_set = _run(raw_path)
    # rate_20 == 1700 excludes the real "Nhava Sheva" origin row (raw row
    # 49, its own rate 1600), which also resolves to origin_code "INNSA"
    # and would otherwise collide with the ICD-cluster "Indore Cluster" row.
    icd_dr_rows = [
        r for r in row_set.rates
        if r.cgo_type == "DR" and r.prefix == "D" and r.origin_code in ("INMUN;INNSA", "INNSA") and r.rate_20 == 1700
    ]
    assert [r.origin_code for r in icd_dr_rows] == [
        "INMUN;INNSA", "INMUN;INNSA", "INMUN;INNSA", "INNSA", "INMUN;INNSA", "INMUN;INNSA",
    ]
    assert all(r.rate_20 == 1700 for r in icd_dr_rows)


def test_nz1_sea_skip_dg_generation():
    raw_path, _, _, _ = PAIRS[0]
    wb = openpyxl.load_workbook(raw_path, data_only=True)
    row_set = Nz1SeaParser().run(wb, MappingProfile(skip_dg_generation={DEFAULT_DESCRIPTION: True}))
    assert "DG" not in {r.cgo_type for r in row_set.rates}
    assert "DR" in {r.cgo_type for r in row_set.rates}


def test_nz1_sea_excluded_charge_codes_drops_both_thl_rows():
    raw_path, _, _, _ = PAIRS[0]
    wb = openpyxl.load_workbook(raw_path, data_only=True)
    row_set = Nz1SeaParser().run(wb, MappingProfile(excluded_charge_codes=["THL"]))
    codes = [n.code for n in row_set.cmdt_notes]
    assert "THL" not in codes
    assert "DOC" in codes  # untouched - only THL was excluded
