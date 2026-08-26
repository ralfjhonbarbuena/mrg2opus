from __future__ import annotations

from mrg2opus.schema.opus_rows import OpusRowSet, RatesRow
from mrg2opus.ui.commodity_utils import assign_sequential_default_codes, distinct_commodity_groups

_BASE_ROW_KWARGS = dict(
    origin_code="CNSHA", origin_description="Shanghai",
    destination_code="USLAX", destination_description="Los Angeles",
    prefix="D", cgo_type="DR",
)


def _row(code: str, description: str) -> RatesRow:
    return RatesRow(commodity_group_code=code, commodity_group_description=description, **_BASE_ROW_KWARGS)


def test_distinct_commodity_groups_preserves_first_encounter_order():
    """Order matters: assign_sequential_default_codes numbers groups in
    this same order, so it must be parse/encounter order, not sorted."""
    row_set = OpusRowSet(
        rates=[
            _row("G0001", "Zebra Group"),
            _row("G0001", "Alpha Group"),  # shares a structural code with Zebra Group
            _row("G0002", "Middle Group"),
        ]
    )
    groups = distinct_commodity_groups({"": row_set})
    assert groups == [("G0001", "Zebra Group"), ("G0001", "Alpha Group"), ("G0002", "Middle Group")]


def test_assign_sequential_default_codes_gives_every_group_its_own_code():
    """Two groups sharing one structural code (e.g. LAWC's main dry grid
    and Reefer both defaulting to G0001 internally) must NOT share the
    new sequential default - each gets its own unique output code."""
    groups = [("G0001", "Zebra Group"), ("G0001", "Alpha Group"), ("G0002", "Middle Group")]
    assert assign_sequential_default_codes(groups) == {
        "Zebra Group": "G0001",
        "Alpha Group": "G0002",
        "Middle Group": "G0003",
    }
