from __future__ import annotations

import openpyxl
from openpyxl.styles import Font, PatternFill

from mrg2opus.parsers.cse import COMMODITY_MAIN, COMMODITY_MAOVLD, COMMODITY_NOR_MAOVLD, COMMODITY_NOR_PA, COMMODITY_VE, CSEParser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols
from tests.golden import _normalize, diff_rates, read_arbs_sheet, read_cmdt_note_sheet, read_rates_sheet, read_special_note_sheet

PATH = "Sample MRGs with OPUS FORMATS/CSE.xlsx"

# Known, verified gap: the "In guage guideline" raw sheet's PAMIT 20'OT/FR
# column reads a flat 8200 for 66 of 69 origins (confirmed directly against
# the raw cells - not a formula, not a caching artifact), while the ground
# truth's Prefix O/F PAMIT rate_20 varies per origin like every other field
# does. No other column/sheet in this file carries a per-origin PAMIT 20'
# in-gauge rate that would explain the ground truth values - this looks
# like the raw sheet's own PAMIT 20' column was left at a placeholder/
# un-updated value for most rows. Every other field (rate_40, every other
# destination, every other prefix) matches exactly.
def _is_known_rates_gap(row: dict) -> bool:
    return row.get("destination_code") == "PAMIT" and row.get("prefix") in {"O", "F"}


# Known, deliberate deviations from ground truth, both user-directed
# business rules applied uniformly (not derived from any one sample):
#   - type: forced to "C" on every row for every lane - CSE's own ground
#     truth leaves it blank.
#   - commodity_group_description: "CSE", "NOR(PA)", and "CSE VE" used to
#     share ONE combined description ("FAK & DG & NOR", G0001); "CSE
#     (MAOVLD)" and "NOR (MAOVLD)" shared another ("FAK - MAOVLD & MAOVLD
#     NOR", G0002). Each of these 5 raw sheets now defaults to its own
#     description (its own sheet name) instead, and only shares one CMDT
#     NOTE block with another sheet if the user overrides them to the same
#     description - see test_cse_cmdt_note_default_splits_by_sheet and
#     test_cse_cmdt_note_merges_when_descriptions_match below.
RATES_DEVIATION_IGNORE = {"type", "commodity_group_description"}


def _run_cse():
    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = CSEParser()
    return parser.run(wb, MappingProfile())


def test_cse_rates_matches_ground_truth():
    row_set = _run_cse()
    generated = [r.model_dump() for r in row_set.rates]
    expected = read_rates_sheet(PATH, cols.SHEET_NAME_RATES)

    matched, missing, extra, mismatches = diff_rates(generated, expected, ignore_fields=RATES_DEVIATION_IGNORE)
    mismatches = [m for m in mismatches if not (m[1] == "rate_20" and _is_known_rates_gap({"destination_code": m[0][1], "prefix": m[0][3]}))]

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"


def test_cse_rates_port_port_matches_ground_truth():
    row_set = _run_cse()
    generated = [r.model_dump() for r in row_set.rates_port_port]
    expected = read_rates_sheet(PATH, cols.SHEET_NAME_RATES_PORT_PORT)

    matched, missing, extra, mismatches = diff_rates(generated, expected, ignore_fields=RATES_DEVIATION_IGNORE)
    mismatches = [m for m in mismatches if not (m[1] == "rate_20" and _is_known_rates_gap({"destination_code": m[0][1], "prefix": m[0][3]}))]

    assert not missing, f"missing {len(missing)} expected rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated rows, e.g. {list(extra)[:5]}"
    assert not mismatches, f"{len(mismatches)} field mismatches, e.g. {mismatches[:10]}"


# Known, verified gap: one specific CMDT NOTE child row (the first of two
# THL entries) carries pol="BDCGP" (Chittagong) in the ground truth - tied
# to a Chittagong-specific rate-structure override mentioned in the raw
# sheet's footnotes ("* Chittagong: Rates are incl. THL/PSS/OBS/EFS/MBS...")
# but not derivable with confidence from that text alone (why THL
# specifically, and only the first of two occurrences, isn't clear).
CMDT_NOTE_KNOWN_GAP_FIELDS = {"pol"}


def test_cse_cmdt_note_default_splits_by_sheet():
    """Default behavior: "CSE", "NOR(PA)", and "CSE VE" (previously one
    combined G0001 description) and "CSE (MAOVLD)"/"NOR (MAOVLD)"
    (previously one combined G0002 description) each get their own
    description (their own raw sheet name) and their own CMDT NOTE block -
    a deliberate, user-directed default (see RATES_DEVIATION_IGNORE
    comment). "In guage guideline" (G0003) was always independent."""
    row_set = _run_cse()
    descriptions = {r.commodity_group_description for r in row_set.rates}
    assert descriptions == {"CSE", "CSE (MAOVLD)", "CSE VE", "NOR(PA)", "NOR (MAOVLD)", "IN GUAGE GUIDELINE (IG)"}
    blocks = [r for r in row_set.cmdt_notes if r.code == "APP"]
    assert len(blocks) == 6


def test_cse_cmdt_note_merges_when_descriptions_match():
    """If the user overrides "CSE"/"NOR(PA)"/"CSE VE"'s descriptions back
    to the ORIGINAL combined ground-truth text (and "CSE (MAOVLD)"/"NOR
    (MAOVLD)" similarly), they collapse back into exactly the ground
    truth's CMDT NOTE - a regression check that the merge-by-description
    logic reconstructs the originally-verified behavior exactly."""
    wb = openpyxl.load_workbook(PATH, data_only=True)
    parser = CSEParser()
    main_combined = "FAK & DG & NOR"
    maovld_combined = "FAK - MAOVLD & MAOVLD NOR"
    profile = MappingProfile(
        commodity_description_overrides={
            COMMODITY_MAIN[0]: main_combined,
            COMMODITY_VE[0]: main_combined,
            COMMODITY_NOR_PA[0]: main_combined,
            COMMODITY_MAOVLD[0]: maovld_combined,
            COMMODITY_NOR_MAOVLD[0]: maovld_combined,
        }
    )
    row_set = parser.run(wb, profile)
    generated = [r.model_dump() for r in row_set.cmdt_notes]
    expected = read_cmdt_note_sheet(PATH, cols.SHEET_NAME_CMDT_NOTE)

    assert len(generated) == len(expected)
    ignore = {"header_seq", "note_seq", *CMDT_NOTE_KNOWN_GAP_FIELDS}
    for i, (g, e) in enumerate(zip(generated, expected)):
        for field_name in cols.CMDT_NOTE_ROW_FIELDS:
            if field_name in ignore:
                continue
            gv = _normalize(g.get(field_name))
            ev = _normalize(e.get(field_name))
            assert gv == ev, f"row {i}: {field_name}: {gv!r} != {ev!r}"


def _arbs_key(row: dict) -> tuple:
    return (row.get("point"), row.get("over"), row.get("per"))


# Known, verified gap: a handful of inland Chinese ARBS origins whose
# ground-truth description matches neither the raw sheet's own text nor
# this Location Bank cleanly:
#   - CNLUN, CNMAA, CNTNL, CNXAN: ground truth uses a different
#     romanization/level of detail than the raw text ("Lu'An" -> "LIUAN"
#     not "LUAN"; "Inner Mongolia" -> "NEI MONGOL"; "XIANGYANG, HUBEI" ->
#     plain "XIANGYANG") - not derivable from anything in this file.
#   - CNCLJ, CNCGO: the OPPOSITE problem - this Location Bank's own mined
#     entry is wrong for these two codes (says "YUEYANG"/"SHANDONG" instead
#     of "CHENGLINGJI"/"HENAN"), most likely a bad mining result from a
#     coincidental code collision in another lane's sample data. The raw
#     sheet's own text is correct here, but preferring raw text universally
#     would break CNCTU (Chengdu) the other way (see cse.py's fallback
#     comment) - there's no single rule that's right for both directions
#     without a real UN/LOCODE-equivalent reference to arbitrate.
ARBS_DESCRIPTION_GAP_CODES = {"CNLUN", "CNMAA", "CNTNL", "CNXAN", "CNCLJ", "CNCGO"}


def test_cse_arbs_matches_ground_truth():
    row_set = _run_cse()
    generated = [r.model_dump() for r in row_set.arbs]
    expected = read_arbs_sheet(PATH)

    gen_by_key = {_arbs_key(r): r for r in generated}
    exp_by_key = {_arbs_key(r): r for r in expected}

    missing = set(exp_by_key) - set(gen_by_key)
    extra = set(gen_by_key) - set(exp_by_key)
    assert not missing, f"missing {len(missing)} expected ARBS rows, e.g. {list(missing)[:5]}"
    assert not extra, f"{len(extra)} unexpected generated ARBS rows, e.g. {list(extra)[:5]}"

    mismatches = []
    for key in gen_by_key:
        if key[0] in ARBS_DESCRIPTION_GAP_CODES:
            continue
        g, e = gen_by_key[key], exp_by_key[key]
        for field_name in cols.ARBS_ROW_FIELDS:
            gv, ev = _normalize(g.get(field_name)), _normalize(e.get(field_name))
            if gv != ev:
                mismatches.append((key, field_name, gv, ev))
    assert not mismatches, f"{len(mismatches)} ARBS field mismatches, e.g. {mismatches[:10]}"


def test_cse_grid_sheet_handles_shifted_anchor_row():
    """Regression test for a real user file: an extra note row inserted
    above "SERVICE SCOPE = CSE" shifted every row below it down by 1 versus
    the bundled sample's layout, which the hardcoded GridSheetConfig row
    numbers didn't account for. The parser read the real POD-name row as
    if it were the POD-code row, every container_map.suffix_for() lookup
    failed, and the whole "CSE" sheet silently produced zero GridRows (the
    reported bug) while sibling sheets without the shift ("CSE (MAOVLD)")
    parsed fine. _resolve_grid_config must detect the anchor's real row via
    its literal text and re-derive the header/data rows from it."""
    from mrg2opus.parsers.cse import GRID_SHEETS, _resolve_grid_config

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CSE"
    ws.cell(row=7, column=3, value="SERVICE SCOPE = CSE")  # sample has this at row 6
    ws.cell(row=8, column=3, value="Test Port")
    ws.cell(row=9, column=3, value="TSTPT")
    ws.cell(row=10, column=3, value="D2")
    ws.cell(row=11, column=1, value="Test Origin")
    ws.cell(row=11, column=2, value="TSTORG")
    ws.cell(row=11, column=3, value=1234)

    base_cfg = next(c for c in GRID_SHEETS if c.sheet_name == "CSE")
    cfg = _resolve_grid_config(ws, base_cfg)
    assert cfg.data_min_row == base_cfg.data_min_row + 1

    rows = CSEParser()._parse_grid_sheet(ws, cfg)
    assert len(rows) == 1
    assert rows[0].origin_code_raw == "TSTORG"
    assert rows[0].dest_code == "TSTPT"
    assert rows[0].sizes.get("20") == 1234


def test_cse_grid_sheet_skips_withdrawn_origin():
    """A struck-through or blacked-out origin name/code cell marks that
    whole origin as not offered (withdrawn), not just one rate cell - the
    row must be dropped entirely rather than parsed with a valid-looking
    origin code. Rate cells on the row are otherwise normal (unstruck), so
    this specifically exercises the origin-label check, not the existing
    per-cell rate exclusion."""
    from mrg2opus.parsers.cse import GRID_SHEETS, _resolve_grid_config

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CSE"
    ws.cell(row=6, column=3, value="SERVICE SCOPE = CSE")
    ws.cell(row=7, column=3, value="Test Port")
    ws.cell(row=8, column=3, value="TSTPT")
    ws.cell(row=9, column=3, value="D2")
    ws.cell(row=10, column=1, value="Withdrawn Origin")
    ws.cell(row=10, column=2, value="OLDORG")
    ws.cell(row=10, column=2).font = Font(strike=True)
    ws.cell(row=10, column=3, value=1234)
    ws.cell(row=11, column=1, value="Active Origin")
    ws.cell(row=11, column=2, value="NEWORG")
    ws.cell(row=11, column=3, value=5678)

    base_cfg = next(c for c in GRID_SHEETS if c.sheet_name == "CSE")
    cfg = _resolve_grid_config(ws, base_cfg)
    rows = CSEParser()._parse_grid_sheet(ws, cfg)

    assert len(rows) == 1
    assert rows[0].origin_code_raw == "NEWORG"


def test_cse_grid_sheet_skips_withdrawn_destination():
    """A struck-through/blacked-out POD label withdraws that whole
    destination - covered generically by flatten_pod_header's own tests,
    this confirms it actually wires through CSE's grid parsing."""
    from mrg2opus.parsers.cse import GRID_SHEETS, _resolve_grid_config

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CSE"
    ws.cell(row=6, column=3, value="SERVICE SCOPE = CSE")
    ws.cell(row=7, column=3, value="Withdrawn Port")
    ws.cell(row=7, column=6, value="Active Port")
    ws.cell(row=8, column=3, value="OLDPT")
    ws.cell(row=8, column=3).fill = PatternFill(start_color="FF000000", end_color="FF000000", fill_type="solid")
    ws.cell(row=8, column=6, value="NEWPT")
    ws.cell(row=9, column=3, value="D2")
    ws.cell(row=9, column=6, value="D2")
    ws.cell(row=10, column=1, value="Origin")
    ws.cell(row=10, column=2, value="ORG")
    ws.cell(row=10, column=3, value=1000)
    ws.cell(row=10, column=6, value=2000)

    base_cfg = next(c for c in GRID_SHEETS if c.sheet_name == "CSE")
    cfg = _resolve_grid_config(ws, base_cfg)
    rows = CSEParser()._parse_grid_sheet(ws, cfg)

    assert len(rows) == 1
    assert rows[0].dest_code == "NEWPT"


def test_cse_special_note_matches_ground_truth():
    row_set = _run_cse()
    generated = [r.model_dump() for r in row_set.special_notes]
    expected = read_special_note_sheet(PATH)

    assert len(generated) == len(expected)
    ignore = {"header_seq", "note_seq"}  # externally-assigned running sequence numbers, not derivable from this file
    for i, (g, e) in enumerate(zip(generated, expected)):
        for field_name in cols.SPECIAL_NOTE_ROW_FIELDS:
            if field_name in ignore:
                continue
            gv = _normalize(g.get(field_name))
            ev = _normalize(e.get(field_name))
            assert gv == ev, f"row {i}: {field_name}: {gv!r} != {ev!r}"
