"""VERTICAL RATES derivation - the two-level blank-fill in particular.
See schema/opus_rows.py::build_vertical_rates."""
from __future__ import annotations

from mrg2opus.schema.opus_rows import (
    OpusRowSet,
    RatesRow,
    build_vertical_rates,
    explode_to_vertical_rates,
)


def _rate_row(cmdt_seq, route_seq, rate_20, rate_40=None):
    return RatesRow(
        cmdt_seq=cmdt_seq, commodity_group_code="G0001", commodity_group_description="FAK",
        route_seq=route_seq, origin_code="AUADL", origin_description="ADELAIDE, SA", origin_term="CY",
        destination_code="BEANR", destination_description="ANTWERP", destination_term="CY",
        prefix="D", cgo_type="DR",
        cur_20="USD", rate_20=rate_20,
        cur_40=("USD" if rate_40 is not None else None), rate_40=rate_40,
    )


def test_vertical_rates_state_commodity_block_once_per_cmdt_seq():
    """Confirmed against reference/2_OPUS/27's real WEW VERTICAL RATES:
    5,209 data rows with CMDT Seq on just 4 of them, one per commodity
    group, while Route Seq and origin/destination repeat per rate row.
    OPUS reads a blank commodity block as "same group as above"."""
    row_set = OpusRowSet(rates=[
        _rate_row(1, 1, 100, 200),   # opens group 1
        _rate_row(1, 2, 300),        # same group - must NOT restate it
        _rate_row(2, 1, 400),        # opens group 2
    ])
    vertical = build_vertical_rates(row_set).vertical_rates

    assert [v.cmdt_seq for v in vertical] == [1, None, None, 2]
    assert [v.commodity_group_code for v in vertical] == ["G0001", None, None, "G0001"]
    # Route Seq and the origin block DO repeat, on each rate row's first line.
    assert [v.route_seq for v in vertical] == [1, None, 2, 1]
    assert [v.origin_code for v in vertical] == ["AUADL", None, "AUADL", "AUADL"]
    assert [v.per for v in vertical] == ["D2", "D4", "D2", "D2"]


def test_vertical_rates_make_each_group_contiguous():
    """RATES row order doesn't guarantee contiguous groups - TAD WMW/WEW's
    WEW scope runs 1,2,3,2,1,2,3,2 once DG duplicates are appended. That's
    fine on RATES (matched by identity, not position) but here the
    blank-fill makes order load-bearing, and every real VERTICAL RATES
    sheet states each group exactly once."""
    row_set = OpusRowSet(rates=[_rate_row(1, 1, 100), _rate_row(2, 1, 200), _rate_row(1, 2, 300)])
    vertical = build_vertical_rates(row_set).vertical_rates

    stated = [v.cmdt_seq for v in vertical if v.cmdt_seq is not None]
    assert stated == [1, 2], "each group stated once, in first-appearance order"
    # group 1's two rows are now adjacent, group 2's follows
    assert [v.route_seq for v in vertical] == [1, 2, 1]


def _grouped_row() -> RatesRow:
    """LAWC's real shape: four origins, one destination, three rate slots.
    Codes are sorted by CODE and names by NAME, so the two orders differ -
    which is exactly why the names can't be paired positionally."""
    return RatesRow(
        commodity_group_code="G0001",
        commodity_group_description="S.E.A DRY",
        origin_code="VNBHA;VNCMP;VNDIA;VNSGN",
        origin_description="CAI MEP;DI AN, BINH DUONG;DONG NAI, BIEN HOA;HO CHI MINH",
        origin_term="CY",
        destination_code="MXZLO",
        destination_description="MANZANILLO",
        destination_term="CY",
        prefix="D",
        cgo_type="DR",
        rate_20=5800,
        rate_40=6100,
        rate_40hc=6100,
        route_seq=17,
    )


def test_each_location_gets_its_own_row_listed_downwards():
    rows = explode_to_vertical_rates(_grouped_row())

    # 4 origins vs 1 destination vs 3 rate slots -> the longest wins.
    assert len(rows) == 4
    assert [r.origin_code for r in rows] == ["VNBHA", "VNCMP", "VNDIA", "VNSGN"]
    # Independent columns, not a cartesian product: the single destination
    # and the three rates stop where they run out.
    assert [r.destination_code for r in rows] == ["MXZLO", None, None, None]
    assert [r.per for r in rows] == ["D2", "D4", "D5", None]
    assert [r.rate for r in rows] == [5800, 6100, 6100, None]


def test_route_seq_and_commodity_block_are_written_once_at_the_top():
    rows = explode_to_vertical_rates(_grouped_row())

    assert [r.route_seq for r in rows] == [17, None, None, None]
    assert [r.commodity_group_code for r in rows] == ["G0001", None, None, None]


def test_each_location_takes_its_own_name_not_the_one_beside_it():
    """The grouped row sorts codes by code and names by name, so zipping
    them pairs VNBHA with CAI MEP. Names come from the Location Bank."""
    rows = explode_to_vertical_rates(_grouped_row())

    assert [r.origin_description for r in rows] == [
        "DONG NAI, BIEN HOA", "CAI MEP", "DI AN, BINH DUONG", "HO CHI MINH",
    ]
    assert rows[0].destination_description == "MANZANILLO"


def test_a_single_location_row_is_unaffected():
    row = _grouped_row()
    row.origin_code = "CNSHA"
    row.origin_description = "SHANGHAI"

    rows = explode_to_vertical_rates(row)

    assert len(rows) == 3  # now the three rate slots are the longest column
    assert [r.origin_code for r in rows] == ["CNSHA", None, None]
