from __future__ import annotations

import io

import openpyxl

from mrg2opus.audit.compare import (
    AUDIT_CONCAT_FIELDS,
    RATES_PRESENTATION_FIELDS,
    arbs_row_key,
    audit_concat,
    audit_row_key,
    diff_by_key,
    diff_vertical_blocks,
    explain_profile_overrides,
    find_duplicate_filings,
    find_sheet,
    freetime_compared_fields,
    profile_skips_rows,
    profile_without_row_skips,
    rates_row_key,
    read_arbs_sheet,
    read_rates_sheet,
    reconstruct_vertical_blocks,
    route_note_counts,
    split_mismatches_by_tier,
)
from mrg2opus.audit.side_by_side import build_side_by_side_workbook
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols


def _rates_sheet_wb(rows: list[list]) -> "openpyxl.Workbook":
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = cols.SHEET_NAME_RATES
    # Rows start at row 3 (rows 1-2 are the 2-row header) - matches every
    # lane's writer output and golden.py's existing read_rates_sheet.
    for offset, row in enumerate(rows):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=3 + offset, column=col_idx, value=value)
    return wb


def test_find_sheet_exact_match():
    wb = openpyxl.Workbook()
    wb.active.title = "OPUS RATES"
    assert find_sheet(wb, "OPUS RATES") == "OPUS RATES"


def test_find_sheet_loose_whitespace_hyphen_match():
    wb = openpyxl.Workbook()
    wb.active.title = "OPUS RATES PORT - PORT"  # SAF's spacing quirk
    assert find_sheet(wb, "OPUS RATES PORT-PORT") == "OPUS RATES PORT - PORT"


def test_find_sheet_raises_when_missing():
    wb = openpyxl.Workbook()
    wb.active.title = "Something Else"
    try:
        find_sheet(wb, "OPUS RATES")
        assert False, "expected KeyError"
    except KeyError:
        pass


def test_read_rates_sheet_skips_blank_rows_and_reads_by_field_order():
    row = [None] * len(cols.RATES_ROW_FIELDS)
    row[cols.RATES_ROW_FIELDS.index("origin_code")] = "CNSHA"
    row[cols.RATES_ROW_FIELDS.index("destination_code")] = "USLAX"
    wb = _rates_sheet_wb([row, [None] * len(cols.RATES_ROW_FIELDS)])
    rows = read_rates_sheet(wb, cols.SHEET_NAME_RATES)
    assert len(rows) == 1
    assert rows[0]["origin_code"] == "CNSHA"
    assert rows[0]["destination_code"] == "USLAX"


def test_arbs_row_key():
    assert arbs_row_key({"point": "CNSHA", "over": "20MT", "per": "MT", "seq": 1}) == ("CNSHA", "20MT", "MT")


def test_diff_by_key_matched_missing_extra_and_field_mismatch():
    generated = [
        {"k": "A", "v": 1},
        {"k": "B", "v": 2},
        {"k": "C", "v": 99},
    ]
    expected = [
        {"k": "A", "v": 1},
        {"k": "C", "v": 3},
        {"k": "D", "v": 4},
    ]
    result = diff_by_key(generated, expected, key_fn=lambda r: (r["k"],), fields=["v"])
    assert result.matched == 1
    assert result.missing == {("D",)}
    assert result.extra == {("B",)}
    assert result.field_mismatches == [(("C",), "v", 99, 3)]


def test_diff_by_key_respects_ignore_fields():
    generated = [{"k": "A", "v": 1, "note": "x"}]
    expected = [{"k": "A", "v": 1, "note": "y"}]
    result = diff_by_key(
        generated, expected, key_fn=lambda r: (r["k"],), fields=["v", "note"], ignore_fields={"note"}
    )
    assert result.matched == 1
    assert result.field_mismatches == []


from mrg2opus.audit.compare import diff_cmdt_blocks, reconstruct_blocks


def _cmdt_row(header_seq=None, note_seq=None, contents=None, charge_seq=None, code=None, amount=None) -> dict:
    row = dict.fromkeys(cols.CMDT_NOTE_ROW_FIELDS)
    row.update(header_seq=header_seq, note_seq=note_seq, contents=contents, charge_seq=charge_seq, code=code, amount=amount)
    return row


def test_reconstruct_blocks_groups_children_under_parent():
    rows = [
        _cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS"),
        _cmdt_row(charge_seq=2, code="OBS"),
        _cmdt_row(header_seq=2, note_seq=2, contents="Block B", charge_seq=1, code="EFS"),
    ]
    blocks = reconstruct_blocks(rows)
    assert [b.key for b in blocks] == ["Block A", "Block B"]
    assert len(blocks[0].children) == 1
    assert blocks[0].children[0]["code"] == "OBS"
    assert len(blocks[1].children) == 0


def test_reconstruct_blocks_detects_parent_via_contents_when_seq_fields_are_none():
    """Regression test: a freshly-parsed OpusRowSet's CmdtNoteRows never
    have header_seq/note_seq set (those are writer-assigned at Excel
    export time) - reconstruct_blocks must still detect block boundaries
    via `contents` alone in that case."""
    rows = [
        _cmdt_row(header_seq=None, note_seq=None, contents="Block A", charge_seq=1, code="PSS"),
        _cmdt_row(header_seq=None, note_seq=None, contents=None, charge_seq=2, code="OBS"),
    ]
    blocks = reconstruct_blocks(rows)
    assert [b.key for b in blocks] == ["Block A"]
    assert len(blocks[0].children) == 1
    assert blocks[0].children[0]["code"] == "OBS"


def test_diff_cmdt_blocks_matched_identical():
    generated = [
        _cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS"),
        _cmdt_row(charge_seq=2, code="OBS"),
    ]
    expected = [
        _cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS"),
        _cmdt_row(charge_seq=2, code="OBS"),
    ]
    result = diff_cmdt_blocks(generated, expected, cols.CMDT_NOTE_ROW_FIELDS)
    assert result.missing_blocks == []
    assert result.extra_blocks == []
    assert result.field_mismatches == []


def test_diff_cmdt_blocks_missing_and_extra():
    generated = [_cmdt_row(header_seq=1, note_seq=1, contents="Only Generated", charge_seq=1, code="PSS")]
    expected = [_cmdt_row(header_seq=1, note_seq=1, contents="Only Reference", charge_seq=1, code="PSS")]
    result = diff_cmdt_blocks(generated, expected, cols.CMDT_NOTE_ROW_FIELDS)
    assert result.missing_blocks == ["Only Reference"]
    assert result.extra_blocks == ["Only Generated"]


def test_diff_cmdt_blocks_field_mismatch_within_matched_block():
    generated = [_cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS", amount=100)]
    expected = [_cmdt_row(header_seq=1, note_seq=1, contents="Block A", charge_seq=1, code="PSS", amount=200)]
    result = diff_cmdt_blocks(generated, expected, cols.CMDT_NOTE_ROW_FIELDS)
    assert result.missing_blocks == []
    assert result.extra_blocks == []
    assert len(result.field_mismatches) == 1
    key, idx, field_name, gv, ev = result.field_mismatches[0]
    assert (key, idx, field_name, gv, ev) == ("Block A", 0, "amount", 100, 200)


# --- field tiers ------------------------------------------------------------
# A renamed commodity code produces one mismatch per row, so reported in the
# same bucket as a wrong rate it buries it. See compare.py's tier comment.

def test_substance_and_presentation_are_reported_separately():
    mismatches = [
        {"key": ("A",), "field": "rate_20", "generated": 100, "reference": 120},
        {"key": ("A",), "field": "commodity_group_code", "generated": "LWE01", "reference": "G0001"},
        {"key": ("B",), "field": "commodity_group_code", "generated": "LWE01", "reference": "G0001"},
    ]

    substance, presentation = split_mismatches_by_tier(mismatches, RATES_PRESENTATION_FIELDS)

    assert [m["field"] for m in substance] == ["rate_20"]
    assert [m["field"] for m in presentation] == ["commodity_group_code", "commodity_group_code"]


def test_unlisted_fields_count_as_substance():
    """Substance is the default, so a newly added column is treated as
    load-bearing until someone decides otherwise."""
    substance, presentation = split_mismatches_by_tier(
        [{"key": ("A",), "field": "some_new_column", "generated": 1, "reference": 2}],
        RATES_PRESENTATION_FIELDS,
    )
    assert len(substance) == 1 and not presentation


# --- explaining differences from the profile --------------------------------

def test_a_default_profile_explains_nothing():
    assert explain_profile_overrides(MappingProfile()) == {}


def test_each_override_names_the_column_it_changes_and_an_example():
    profile = MappingProfile(
        commodity_code_overrides={"CSE": "LWE01", "NOR": "LWE02"},
        commodity_sequence_overrides={"CSE": 4},
    )

    explained = explain_profile_overrides(profile)

    assert set(explained) == {"commodity_group_code", "cmdt_seq"}
    assert "'CSE' → 'LWE01'" in explained["commodity_group_code"]
    assert "and 1 more" in explained["commodity_group_code"]
    assert "and 1 more" not in explained["cmdt_seq"]


def test_group_order_is_not_an_expected_difference():
    """It changes the order rows are written in, and the diff is keyed
    rather than positional, so it cannot produce one."""
    assert explain_profile_overrides(MappingProfile(commodity_group_order=["A", "B"])) == {}


# --- row-removing settings --------------------------------------------------

def test_row_skips_are_detected_only_when_actually_set():
    assert not profile_skips_rows(MappingProfile())
    assert not profile_skips_rows(MappingProfile(skip_commodity_filing={"CSE": False}))
    assert profile_skips_rows(MappingProfile(skip_commodity_filing={"CSE": True}))
    assert profile_skips_rows(MappingProfile(skip_dg_generation={"CSE": True}))


def test_disabling_skips_leaves_every_other_setting_alone():
    profile = MappingProfile(
        commodity_code_overrides={"CSE": "LWE01"},
        skip_commodity_filing={"CSE": True},
        skip_dg_generation={"NOR": True},
        include_vertical_rates=False,
    )

    without = profile_without_row_skips(profile)

    assert without.skip_commodity_filing == {} and without.skip_dg_generation == {}
    assert without.commodity_code_overrides == {"CSE": "LWE01"}
    assert without.include_vertical_rates is False


# --- the auditor's own row identity -----------------------------------------
# The audit is done by CONCATENATEing these columns in Excel and matching
# that against the other draft; Compare matches on the same thing.

def _rates_dict(**overrides):
    row = dict.fromkeys(cols.RATES_ROW_FIELDS)
    row.update(
        origin_code="MYPKG", origin_term="CY", destination_code="MXZLO",
        destination_term="CY", prefix="D", cgo_type="DG", rate_20=5800,
    )
    row.update(overrides)
    return row


def test_the_audit_key_separates_rows_the_route_key_collapses():
    """Two real LAWC rows: same route pair and cargo type, different rates
    and route note. rates_row_key makes them one key, and diff_by_key's
    {key: row} dict would drop one of them unseen."""
    a = _rates_dict(route_note=None, rate_20=5800, rate_40=6100)
    b = _rates_dict(route_note="REEFER DRY AS DANGEROUS", rate_20=None, rate_40hc=6000)

    assert rates_row_key(a) == rates_row_key(b)
    assert audit_row_key(a) != audit_row_key(b)


def test_rates_stay_out_of_the_match_key():
    """So a wrong rate reports as a rate difference on a matched route,
    not as the route vanishing from one side."""
    cheap, dear = _rates_dict(rate_20=100), _rates_dict(rate_20=999)
    assert audit_row_key(cheap) == audit_row_key(dear)


def test_the_concat_does_include_the_rates():
    assert audit_concat(_rates_dict(rate_20=100)) != audit_concat(_rates_dict(rate_20=999))


def test_a_row_repeated_exactly_is_a_duplicate_filing():
    """What OPUS rejects, and what the audit checks for by hand."""
    row = _rates_dict()
    assert find_duplicate_filings([row]) == []
    assert find_duplicate_filings([row, dict(row)]) == [(audit_concat(row), 2)]


def test_rows_differing_only_on_a_rate_are_not_duplicates():
    assert find_duplicate_filings([_rates_dict(rate_20=100), _rates_dict(rate_20=999)]) == []


# --- the three sheets that used not to be compared at all -------------------

def test_a_blank_cell_and_an_empty_string_are_the_same_value():
    """References use both - LAWC's own FREETIME writes "" where we write
    nothing, which reported as 176 differences across 22 identical rows."""
    a = {"seq": 1, "coverage_rgn": None}
    b = {"seq": 1, "coverage_rgn": ""}
    assert not diff_by_key([a], [b], key_fn=lambda r: (r["seq"],), fields=["coverage_rgn"]).field_mismatches


def test_route_notes_are_compared_by_text_and_lane_not_by_row():
    """RN is addressed by (Header Seq, Route Seq), both assigned by OPUS -
    LAWC's real filing numbers its headers from 1015 - so rows can't be
    matched one to one."""
    rows = [
        {"contents": "Vessel Service Lane: MX2", "lane": "MX2", "header_seq": 1015},
        {"contents": "Vessel Service Lane: MX2", "lane": "MX2", "header_seq": 9999},
        {"contents": "  ", "lane": None},
    ]
    counts = route_note_counts(rows)
    assert counts == {("Vessel Service Lane: MX2", "MX2"): 2}


def _vertical(route_seq=None, origin=None, dest=None, per=None, cargo=None, rate=None):
    return {
        "route_seq": route_seq, "origin_code": origin, "destination_code": dest,
        "origin_term": "CY", "destination_term": "CY", "o_via_code": None, "d_via_code": None,
        "per": per, "cargo_type": cargo, "rate": rate,
    }


def test_a_vertical_block_gathers_the_rows_below_its_route_seq():
    """The sheet is columnar: a row carrying only an origin is the second
    origin of the block above, not a route of its own."""
    rows = [
        _vertical(route_seq=17, origin="VNBHA", dest="MXZLO", per="D2", cargo="DR", rate=5800),
        _vertical(origin="VNCMP", per="D4", cargo="DR", rate=6100),
        _vertical(origin="VNSGN"),
        _vertical(route_seq=18, origin="VNDAD", dest="MXZLO", per="D2", cargo="DR", rate=5850),
    ]

    blocks = reconstruct_vertical_blocks(rows)

    assert len(blocks) == 2
    assert blocks[0].origins == frozenset({"VNBHA", "VNCMP", "VNSGN"})
    assert blocks[0].destinations == frozenset({"MXZLO"})
    assert blocks[0].rates == (("D2", "DR", 5800), ("D4", "DR", 6100))


def test_cargo_type_is_part_of_the_block_key():
    """One route files a dry block and a reefer block back to back; without
    cargo type in the key they collapse and one is dropped unseen."""
    dry = reconstruct_vertical_blocks([_vertical(route_seq=1, origin="AUADL", dest="BEANR", per="D2", cargo="DR", rate=1941)])
    reefer = reconstruct_vertical_blocks([_vertical(route_seq=2, origin="AUADL", dest="BEANR", per="R2", cargo="RF", rate=1755)])
    assert dry[0].key != reefer[0].key


def test_vertical_blocks_report_a_rate_difference_on_a_matched_route():
    ours = [_vertical(route_seq=1, origin="AUADL", dest="BEANR", per="D2", cargo="DR", rate=1941)]
    theirs = [_vertical(route_seq=1, origin="AUADL", dest="BEANR", per="D2", cargo="DR", rate=9999)]

    missing, extra, differing = diff_vertical_blocks(ours, theirs)

    assert not missing and not extra
    assert len(differing) == 1


def test_freetime_skips_the_columns_we_deliberately_leave_blank():
    compared = freetime_compared_fields()
    assert "rfa_no" not in compared and "status" not in compared and "dar_no" not in compared
    assert "tariff" in compared and "free_time_total" in compared


# --- the side-by-side workbook ----------------------------------------------

def test_side_by_side_pairs_each_sheet_and_points_the_lookups_at_each_other():
    rows = [dict.fromkeys(cols.RATES_ROW_FIELDS) | {"origin_code": "CNSHA", "rate_20": 100}]
    data = build_side_by_side_workbook([("RATES", "RATES", rows, rows)])

    wb = openpyxl.load_workbook(io.BytesIO(data))
    assert wb.sheetnames == ["HOW TO USE", "RATES (ours)", "RATES (ref)"]

    ours = wb["RATES (ours)"]
    assert ours["A1"].value == "MATCH KEY"
    # BOTH columns are live formulas, not baked answers, so they
    # re-evaluate after an edit to either sheet.
    assert ours["A2"].value.startswith("=") and '&"|"&' in ours["A2"].value
    assert ours["B2"].value.startswith("=_xlfn.XLOOKUP(") and "'RATES (ref)'" in ours["B2"].value
    assert "'RATES (ours)'" in wb["RATES (ref)"]["B2"].value


def test_side_by_side_writes_an_empty_reference_side_rather_than_omitting_it():
    """LAWC files no VERTICAL RATES at all; the pairing should still be
    visible instead of the sheet quietly missing."""
    rows = [dict.fromkeys(cols.RATES_ROW_FIELDS) | {"origin_code": "CNSHA"}]
    wb = openpyxl.load_workbook(io.BytesIO(build_side_by_side_workbook([("RATES", "RATES", rows, [])])))
    assert wb["RATES (ref)"].max_row == 1  # header only


def test_a_long_sheet_label_is_trimmed_to_excels_limit():
    rows = [dict.fromkeys(cols.RATES_ROW_FIELDS)]
    wb = openpyxl.load_workbook(io.BytesIO(build_side_by_side_workbook([("RATES", "X" * 40, rows, rows)])))
    assert all(len(name) <= 31 for name in wb.sheetnames)


def test_the_match_key_formula_points_at_this_row_and_covers_the_concat():
    rows = [dict.fromkeys(cols.RATES_ROW_FIELDS)]
    wb = openpyxl.load_workbook(io.BytesIO(build_side_by_side_workbook([("RATES", "RATES", rows, rows)])))
    formula = wb["RATES (ours)"]["A2"].value

    # one reference per concat field, all on row 2, none on the key or
    # lookup columns themselves
    refs = formula.removeprefix("=").split('&"|"&')
    assert len(refs) == len(AUDIT_CONCAT_FIELDS)
    assert all(ref.endswith("2") for ref in refs)
    # not the key or lookup columns themselves - compared as whole column
    # letters, since "AA" also starts with "A"
    assert not {ref.rstrip("0123456789") for ref in refs} & {"A", "B"}


def test_the_rate_columns_are_part_of_the_key_formula():
    """The audit's own concat includes the whole Rate section, so a wrong
    rate has to fail the lookup rather than match."""
    rows = [dict.fromkeys(cols.RATES_ROW_FIELDS)]
    wb = openpyxl.load_workbook(io.BytesIO(build_side_by_side_workbook([("RATES", "RATES", rows, rows)])))
    formula = wb["RATES (ours)"]["A2"].value

    from openpyxl.utils import get_column_letter
    rate_col = get_column_letter(3 + cols.RATES_ROW_FIELDS.index("rate_20"))
    assert f"{rate_col}2" in formula


def test_the_workbook_asks_excel_to_calculate_on_open():
    """Every cell we write is a formula with no cached result, and Excel
    trusts a saved workbook's stored results - so without this it opens
    blank until each cell is entered by hand."""
    import re
    import zipfile

    rows = [dict.fromkeys(cols.RATES_ROW_FIELDS)]
    data = build_side_by_side_workbook([("RATES", "RATES", rows, rows)])
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        book = archive.read("xl/workbook.xml").decode()
    assert re.search(r'<calcPr[^>]*fullCalcOnLoad="1"', book)


def test_post_spec_functions_carry_the_prefix_excel_needs_in_a_file():
    """XLOOKUP and TEXTJOIN postdate the xlsx format: written plainly they
    evaluate to #NAME?, and Excel only accepts them in a FILE as
    _xlfn.NAME. It hides the prefix in the formula bar and re-adds it on
    save, so a cell retyped by hand looks identical and works - which is
    what made this read as a calculation problem rather than a spelling
    one. The key avoids the issue entirely by using "&"."""
    import zipfile

    rows = [dict.fromkeys(cols.RATES_ROW_FIELDS)]
    data = build_side_by_side_workbook([("RATES", "RATES", rows, rows)])
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        sheet = archive.read("xl/worksheets/sheet2.xml").decode()

    assert "_xlfn.XLOOKUP(" in sheet
    assert "TEXTJOIN" not in sheet
    # no bare XLOOKUP anywhere once the prefixed ones are removed
    assert "XLOOKUP" not in sheet.replace("_xlfn.XLOOKUP", "")


def test_every_sheet_type_has_a_layout_and_a_key():
    """Each sheet has its own columns, so one hardcoded RATES layout
    cannot serve them - a note sheet written with rate columns would be
    silently wrong rather than obviously so."""
    from mrg2opus.audit.side_by_side import SHEET_SPECS

    for sheet_type, spec in SHEET_SPECS.items():
        assert spec.headers and len(spec.headers) == len(spec.row_fields), sheet_type
        assert spec.key_fields, sheet_type
        # every key field is a real column on that sheet, or the formula
        # would reference a column that isn't there
        assert set(spec.key_fields) <= set(spec.row_fields), sheet_type


def test_sequence_numbers_are_never_part_of_a_key():
    """OPUS assigns them and neither draft controls them, so including one
    would fail every row on the sheet."""
    from mrg2opus.audit.side_by_side import SHEET_SPECS

    for sheet_type, spec in SHEET_SPECS.items():
        assert not {"cmdt_seq", "route_seq", "header_seq", "note_seq"} & set(spec.key_fields), sheet_type


def test_a_note_sheet_is_written_with_its_own_columns():
    rows = [dict.fromkeys(cols.CMDT_NOTE_ROW_FIELDS) | {"contents": "Rates are valid...", "code": "APP"}]
    wb = openpyxl.load_workbook(io.BytesIO(
        build_side_by_side_workbook([("CMDT NOTE", "CMDT NOTE", rows, rows)])
    ))
    ws = wb["CMDT NOTE (ours)"]
    assert ws.max_column == len(cols.CMDT_NOTE_ROW_FIELDS) + 2
    assert ws["C1"].value == "Header Seq"


def test_an_unknown_sheet_type_is_skipped_not_written_with_the_wrong_columns():
    wb = openpyxl.load_workbook(io.BytesIO(
        build_side_by_side_workbook([("NOT A SHEET", "X", [{}], [{}])])
    ))
    assert wb.sheetnames == ["HOW TO USE"]
