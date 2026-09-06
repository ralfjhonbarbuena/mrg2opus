from __future__ import annotations

from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema.opus_rows import OpusRowSet, RatesRow
from mrg2opus.ui.commodity_utils import (
    assign_sequential_default_codes,
    dg_twin_probe,
    distinct_commodity_groups,
    groups_offering_dg_twins,
)

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


def test_dg_twin_probe_records_the_groups_a_parser_asks_about():
    """A parser asks skip_dg_generation about a group only where it would
    build that group's DG twin, so the questions asked are the answer to
    "which groups is Skip DG meaningful for?"."""
    probe = dg_twin_probe(MappingProfile())

    # Stand-in for a parse: two groups offer a twin, a third never does.
    probe.skip_dg_generation.get("Dry FAK", False)
    probe.skip_dg_generation.get("Dry ISC", False)

    assert groups_offering_dg_twins(probe) == frozenset({"Dry FAK", "Dry ISC"})


def test_dg_twin_probe_keeps_the_skips_the_profile_already_had():
    """Parsing through the probe has to produce exactly what parsing
    without it would, or the preview would show DG rows a loaded preset
    had already turned off."""
    profile = MappingProfile(skip_dg_generation={"Dry FAK": True})
    probe = dg_twin_probe(profile)

    assert probe.skip_dg_generation.get("Dry FAK", False) is True
    assert probe.skip_dg_generation.get("Dry ISC", False) is False
    # Asked about either way - a skipped group still has a twin to skip.
    assert groups_offering_dg_twins(probe) == frozenset({"Dry FAK", "Dry ISC"})


def test_groups_offering_dg_twins_on_an_unprobed_profile_is_empty():
    """TAD lanes never consult the dict - their DG duplicate is one
    lane-wide flag - and "no group has a twin" is the right reading."""
    assert groups_offering_dg_twins(MappingProfile()) == frozenset()
