from __future__ import annotations

import openpyxl

from mrg2opus.schema import opus_columns as cols
from mrg2opus.schema.opus_rows import OpusRowSet, RatesPortPortRow, RatesRow
from mrg2opus.ui.compare_page import _run_comparison, _summary_row


def _rates_row(**overrides) -> RatesRow:
    base = dict(
        commodity_group_code="G0001", commodity_group_description="FAK",
        origin_code="CNSHA", origin_description="Shanghai",
        destination_code="USLAX", destination_description="Los Angeles",
        prefix="D", cgo_type="DR", rate_20=1000,
    )
    base.update(overrides)
    return RatesRow(**base)


def _reference_workbook_with_rates_sheet(rows: list[list]) -> "openpyxl.Workbook":
    # "RATES" matches _run_comparison's own real-filing-name convention
    # (excel_io/writer.py's _sheet_names_for_suffix) - NOT
    # cols.SHEET_NAME_RATES, which is the older bundled-sample fixtures'
    # naming (see that constant's own docstring).
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RATES"
    for offset, row in enumerate(rows):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=3 + offset, column=col_idx, value=value)
    return wb


def test_run_comparison_matches_identical_rates_row():
    generated_row = _rates_row()
    row_sets = {"": OpusRowSet(rates=[generated_row])}

    ref_row = [None] * len(cols.RATES_ROW_FIELDS)
    for field_name, value in generated_row.model_dump().items():
        ref_row[cols.RATES_ROW_FIELDS.index(field_name)] = value
    ref_wb = _reference_workbook_with_rates_sheet([ref_row])

    results, _duplicates, _pairs = _run_comparison(row_sets, ref_wb, "Both", "SAF")
    rates_result = next(r for r in results if r["sheet_type"] == "RATES")
    assert rates_result["found_in_reference"] is True
    assert rates_result["matched"] == 1
    assert rates_result["missing"] == []
    assert rates_result["extra"] == []


def test_run_comparison_reports_missing_sheet_as_all_extra():
    generated_row = _rates_row()
    row_sets = {"": OpusRowSet(rates=[generated_row])}
    ref_wb = openpyxl.Workbook()
    ref_wb.active.title = "Unrelated Sheet"

    results, _duplicates, _pairs = _run_comparison(row_sets, ref_wb, "Grouped (RATES)", "SAF")
    rates_result = next(r for r in results if r["sheet_type"] == "RATES")
    assert rates_result["found_in_reference"] is False
    assert len(rates_result["extra"]) == 1


def test_run_comparison_respects_rates_mode_grouped_only():
    generated_row = _rates_row()
    row_set = OpusRowSet(rates=[generated_row], rates_port_port=[RatesPortPortRow(**generated_row.model_dump())])
    row_sets = {"": row_set}
    ref_wb = openpyxl.Workbook()
    ref_wb.active.title = "Unrelated Sheet"

    results, _duplicates, _pairs = _run_comparison(row_sets, ref_wb, "Grouped (RATES)", "SAF")
    sheet_types = {r["sheet_type"] for r in results}
    assert sheet_types == {"RATES"}


def test_run_comparison_both_mode_includes_grouped_and_exploded():
    generated_row = _rates_row()
    row_set = OpusRowSet(rates=[generated_row], rates_port_port=[RatesPortPortRow(**generated_row.model_dump())])
    row_sets = {"": row_set}
    ref_wb = openpyxl.Workbook()
    ref_wb.active.title = "Unrelated Sheet"

    results, _duplicates, _pairs = _run_comparison(row_sets, ref_wb, "Both", "SAF")
    sheet_types = {r["sheet_type"] for r in results}
    assert sheet_types == {"RATES", "RATES PORT-PORT"}


# --- the summary row --------------------------------------------------------
# A "Matched" column used to count rows with NO difference at all - across
# every column, sequence numbers included, which a reference practically
# never reproduces. It read 0 on sheets where nothing was wrong.

def _result(**over):
    base = {
        "sheet_type": "RATES", "sub_lane": "(default)", "sheet_name": "RATES",
        "found_in_reference": True, "matched": 0, "routes_matched": 7208, "substance_ok": 7208,
        "missing": [], "intentionally_absent": [], "extra": [],
        "field_mismatches": [], "substance_mismatches": [], "presentation_mismatches": [],
    }
    base.update(over)
    return base


def test_a_sheet_with_no_real_difference_reads_OK_however_much_presentation_differs():
    row = _summary_row(_result(presentation_mismatches=[{"x": 1}] * 1952))
    assert row["Verdict"] == "OK"
    assert row["Matched"] == 7208 and row["Agree on substance"] == 7208


def test_a_substance_difference_reads_CHECK():
    assert _summary_row(_result(substance_mismatches=[{"key": ("a",)}]))["Verdict"] == "CHECK"


def test_an_unaccounted_row_reads_CHECK():
    assert _summary_row(_result(missing=[{"key": ("a",)}]))["Verdict"] == "CHECK"
    assert _summary_row(_result(extra=[{"key": ("a",)}]))["Verdict"] == "CHECK"


def test_a_missing_sheet_says_so_rather_than_CHECK():
    assert _summary_row(_result(found_in_reference=False))["Verdict"] == "not in reference"


def test_the_counts_stay_whole_numbers():
    """None in a count column makes pandas widen it to float and print
    7208.0 and NaN."""
    import pandas as pd

    df = pd.DataFrame([_summary_row(_result()), _summary_row(_result(found_in_reference=False, routes_matched=0, substance_ok=0))])
    for column in ("Matched", "Agree on substance", "Substance", "Unaccounted", "Presentation"):
        assert str(df[column].dtype) == "int64", column
