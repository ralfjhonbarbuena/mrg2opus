"""VERTICAL RATES derivation - the two-level blank-fill in particular.
See schema/opus_rows.py::build_vertical_rates."""
from __future__ import annotations

from mrg2opus.schema.opus_rows import OpusRowSet, RatesRow, build_vertical_rates


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
