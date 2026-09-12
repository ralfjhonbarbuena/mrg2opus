"""The Utilities tab: tools that start from a filing you already have."""
from __future__ import annotations

from datetime import date

import openpyxl
import pytest

from mrg2opus.schema.opus_rows import RatesRow
from mrg2opus.utilities.checks import run_all
from mrg2opus.utilities.delta import compare_filings
from mrg2opus.utilities.reshape import from_vertical, regroup_port_port, to_port_port, to_vertical
from mrg2opus.utilities.summary import summarize
from mrg2opus.utilities.workbook import LoadedFiling, load_filing


def _rate(origin="CNSHA", destination="CLVAP", **over):
    base = dict(
        commodity_group_code="G0001", commodity_group_description="FAK",
        origin_code=origin, origin_description="Shanghai", origin_term="CY",
        destination_code=destination, destination_description="Valparaiso", destination_term="CY",
        prefix="D", cgo_type="DR", cmdt_seq=1, route_seq=1,
        cur_20="USD", rate_20=1000, cur_40hc="USD", rate_40hc=1800,
    )
    base.update(over)
    return base


def _filing(**sheets) -> LoadedFiling:
    return LoadedFiling(sheets=sheets, found_as={k: k.upper() for k in sheets})


# --- reading a workbook whoever prepared it ----------------------------------

def _wb(sheet_names):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name in sheet_names:
        wb.create_sheet(name)
    return wb


@pytest.mark.parametrize("name", ["CMDT NOTE", "SRCHG", "SUR", "SURCHARGE"])
def test_the_surcharge_sheet_is_found_under_any_of_its_names(name):
    """One sheet, five spellings across the 45 real filings - SRCHG in 19
    of them, CMDT NOTE in 18, SUR in 9. A tool that only knew one name
    would open a third of them."""
    assert load_filing(_wb(["RATES", name])).found_as.get("cmdt_notes") == name


@pytest.mark.parametrize("name,kind", [("RN", "route_notes"), ("RNT", "route_notes"),
                                       ("V RATES", "vertical_rates")])
def test_the_other_drifting_sheet_names_are_found_too(name, kind):
    assert load_filing(_wb(["RATES", name])).found_as.get(kind) == name


def test_a_sheet_nobody_recognizes_is_reported_rather_than_dropped():
    """Silence about a sheet we didn't read would be the worst answer:
    the user would assume it was included."""
    loaded = load_filing(_wb(["RATES", "Tier 1 List"]))
    assert loaded.unread == ["Tier 1 List"]


# --- reshaping ----------------------------------------------------------------

def test_splitting_ports_gives_one_row_per_port():
    out = to_port_port([_rate(origin="CNSHA;CNNGB;CNYTN")])
    assert [r.origin_code for r in out] == ["CNSHA", "CNNGB", "CNYTN"]
    assert {r.destination_code for r in out} == {"CLVAP"}


def test_vertical_puts_each_container_size_on_its_own_row():
    out = to_vertical([_rate()])
    assert [(r.per, r.rate) for r in out if r.per] == [("D2", 1000), ("D5", 1800)]


def test_a_filing_survives_the_round_trip_through_vertical():
    """The long format is columnar, not a list of records - a route opens
    where Route Seq is written and its ports run down their own columns.
    Those boundaries survive, so this direction restores the grouping
    rather than guessing at it."""
    original = [_rate(origin="CNSHA;CNNGB"), _rate(origin="KRPUS", destination="PECLL", route_seq=2)]

    back = from_vertical([r.model_dump() for r in to_vertical(original)])

    assert len(back) == len(original)
    assert {r.origin_code for r in back} == {"CNNGB;CNSHA", "KRPUS"}
    assert {r.destination_code for r in back} == {"CLVAP", "PECLL"}
    assert {(r.rate_20, r.rate_40hc) for r in back} == {(1000, 1800)}


def test_regrouping_port_port_folds_rather_than_restores():
    """The split threw away which ports shared a row and nothing records
    it, so two origins that agree on everything else are folded whether
    or not they started together. Worth pinning: it is the one conversion
    that can return fewer rows than were filed."""
    separately_filed = [_rate(origin="CNSHA"), _rate(origin="CNNGB")]

    out = regroup_port_port(separately_filed)

    assert len(out) == 1
    assert out[0].origin_code == "CNNGB;CNSHA"


def test_regrouping_keeps_rows_that_differ_in_anything_apart():
    out = regroup_port_port([_rate(origin="CNSHA"), _rate(origin="CNNGB", rate_20=1200)])
    assert len(out) == 2


# --- checking -----------------------------------------------------------------

def _result(filing, name):
    return next(r for r in run_all(filing) if r.name == name)


def test_a_route_filed_twice_is_an_error():
    """OPUS rejects the upload outright, which makes this the check to
    run first."""
    filing = _filing(rates=[_rate(), _rate()])

    result = _result(filing, "Duplicate filings")

    assert not result.passed
    assert result.findings and result.findings[0].severity == "error"


def test_a_clean_filing_says_so_rather_than_saying_nothing():
    """"No duplicates" is something the auditor confirms before
    approving, not merely the absence of a warning."""
    result = _result(_filing(rates=[_rate()]), "Duplicate filings")

    assert result.passed
    assert result.headline == "No route is filed twice"


def test_an_unknown_port_is_a_warning_not_an_error():
    """The bank holds 331 ports and the world holds more - a typo and a
    genuinely new port look identical from here."""
    result = _result(_filing(rates=[_rate(origin="ZZNOP")]), "Location codes")

    assert not result.passed
    assert [f.summary for f in result.findings] == ["ZZNOP is not in the Location Bank"]
    assert [f.severity for f in result.findings] == ["warning"]


def test_a_row_with_no_rate_in_any_size_is_flagged():
    filing = _filing(rates=[_rate(rate_20=None, rate_40hc=None)])

    result = _result(filing, "Required fields")

    assert not result.passed
    assert "no rate" in result.findings[0].summary


def test_route_notes_are_not_read_for_charge_codes():
    """They are header-only by definition - charge_seq and code are always
    1 and APP. Reading them also went wrong in practice: some lanes' RN
    sheets sit a column across, and the sequence number came back as a
    charge code named "1"."""
    filing = _filing(rates=[_rate()], route_notes=[{"code": "1", "contents": "x"}] * 99)

    assert _result(filing, "Charge codes").passed


# --- summarizing --------------------------------------------------------------

def test_the_summary_counts_ports_inside_a_group():
    """A ";"-joined row covers several ports; counting the cell would say
    one."""
    s = summarize(_filing(rates=[_rate(origin="CNSHA;CNNGB;CNYTN")]))

    assert s.origins == 3
    assert s.destinations == 1
    assert s.rate_rows == 1


def test_the_summary_reads_validity_off_the_note_sheet():
    """An OPUS workbook states its window nowhere else."""
    filing = _filing(
        rates=[_rate()],
        cmdt_notes=[{"application_effective": date(2026, 8, 15), "application_expires": date(2026, 8, 31)}],
    )

    assert summarize(filing).validity == (date(2026, 8, 15), date(2026, 8, 31))


# --- the week-over-week delta -------------------------------------------------

def test_a_rate_that_moved_reports_as_a_change_not_as_two_rows():
    """Rates are deliberately out of the match key. With them in, a route
    whose rate moved would vanish from one side and reappear as new in
    the other - which is exactly what this tool exists to avoid saying."""
    delta = compare_filings([_rate(rate_20=1000)], [_rate(rate_20=1150)])

    assert not delta.added and not delta.removed
    assert [(c.size, c.was, c.now) for c in delta.changes] == [("20'", 1000, 1150)]
    assert delta.changes[0].difference == 150
    assert round(delta.changes[0].percent, 1) == 15.0


def test_added_and_removed_routes_are_told_apart():
    delta = compare_filings([_rate(destination="CLVAP")], [_rate(destination="PECLL")])

    assert [r["destination_code"] for r in delta.added] == ["PECLL"]
    assert [r["destination_code"] for r in delta.removed] == ["CLVAP"]
    assert delta.unchanged == 0


def test_an_unchanged_filing_reports_every_route_as_unchanged():
    delta = compare_filings([_rate(), _rate(destination="PECLL")],
                            [_rate(), _rate(destination="PECLL")])

    assert delta.unchanged == 2
    assert not delta.changes and not delta.added and not delta.removed


def test_the_biggest_movers_come_first():
    before = [_rate(destination="CLVAP", rate_20=1000), _rate(destination="PECLL", rate_20=1000)]
    after = [_rate(destination="CLVAP", rate_20=1050), _rate(destination="PECLL", rate_20=1900)]

    delta = compare_filings(before, after)

    assert [c.route for c in delta.changes][0].endswith("PECLL")
