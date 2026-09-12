"""Per-scope commodity settings.

A lane with sub-lanes files each as its own OPUS workbook, so one
commodity group can be coded, named, ordered or dropped differently in
each - TAD AEW/AMW's two Japan scopes file the same group as G0011 and
G0001 in the same round.
"""
from __future__ import annotations

from mrg2opus.pipeline import run_parser
from mrg2opus.presets.models import MappingProfile, ScopeOverrides
from mrg2opus.ui.commodity_utils import commodity_blocks
from mrg2opus.ui.filing_settings import _with_scope_overrides
from mrg2opus.schema.opus_rows import OpusRowSet, RatesRow


def _row(code: str, description: str) -> RatesRow:
    return RatesRow(
        commodity_group_code=code, commodity_group_description=description,
        origin_code="CNSHA", origin_description="Shanghai",
        destination_code="USLAX", destination_description="Los Angeles",
        prefix="D", cgo_type="DR",
    )


def test_a_profile_nobody_scoped_is_returned_unchanged():
    """The common case by far - every single-scope lane, and every
    multi-scope one nobody has customized per scope - so it takes exactly
    the path it took before per-scope settings existed."""
    profile = MappingProfile(commodity_code_overrides={"FAK": "G0001"})
    assert profile.for_scope("AEW") is profile
    assert MappingProfile().for_scope("") is not None


def test_a_scope_overrides_only_the_groups_it_names():
    """The point of merging rather than replacing: a shared setting still
    reaches every scope that hasn't overridden that particular group."""
    profile = MappingProfile(
        commodity_code_overrides={"FAK": "G0001", "FAK - JAPAN": "G0001"},
        by_scope={"JAPAN AEW": ScopeOverrides(commodity_code_overrides={"FAK - JAPAN": "G0011"})},
    )

    japan = profile.for_scope("JAPAN AEW")
    assert japan.commodity_code_overrides == {"FAK": "G0001", "FAK - JAPAN": "G0011"}
    assert profile.for_scope("JAPAN AMW").commodity_code_overrides["FAK - JAPAN"] == "G0001"
    # The profile itself is untouched - for_scope is a view, not a change.
    assert profile.commodity_code_overrides["FAK - JAPAN"] == "G0001"


def test_group_order_replaces_rather_than_merges():
    """An order is a whole answer; half of one means nothing."""
    profile = MappingProfile(
        commodity_group_order=["A", "B", "C"],
        by_scope={"OMW": ScopeOverrides(commodity_group_order=["C", "A"])},
    )
    assert profile.for_scope("OMW").commodity_group_order == ["C", "A"]
    assert profile.for_scope("OEW").commodity_group_order == ["A", "B", "C"]


def test_every_commodity_setting_is_scopeable():
    profile = MappingProfile(
        by_scope={"AMW": ScopeOverrides(
            commodity_code_overrides={"FAK": "X"},
            commodity_description_overrides={"FAK": "Renamed"},
            commodity_sequence_overrides={"FAK": 7},
            skip_commodity_filing={"FAK": True},
            skip_dg_generation={"FAK": True},
        )}
    )
    amw = profile.for_scope("AMW")
    assert amw.commodity_code_overrides == {"FAK": "X"}
    assert amw.commodity_description_overrides == {"FAK": "Renamed"}
    assert amw.commodity_sequence_overrides == {"FAK": 7}
    assert amw.skip_commodity_filing == {"FAK": True}
    assert amw.skip_dg_generation == {"FAK": True}


def test_only_the_differences_are_stored():
    """A scope entry equal to the shared answer would freeze that group
    there, so a later change to the shared setting would silently stop
    reaching this scope."""
    profile = MappingProfile(commodity_code_overrides={"FAK": "G0001", "OOG": "G0002"})

    by_scope = _with_scope_overrides(profile, "AMW", {
        "commodity_code_overrides": {"FAK": "G0001", "OOG": "G0009"},
        "commodity_group_order": [],
    })

    assert by_scope["AMW"].commodity_code_overrides == {"OOG": "G0009"}


def test_a_scope_that_differs_in_nothing_is_dropped_entirely():
    profile = MappingProfile(
        commodity_code_overrides={"FAK": "G0001"},
        by_scope={"AMW": ScopeOverrides(commodity_code_overrides={"FAK": "X"})},
    )

    by_scope = _with_scope_overrides(profile, "AMW", {
        "commodity_code_overrides": {"FAK": "G0001"},
        "commodity_group_order": [],
    })

    assert "AMW" not in by_scope


def test_commodity_blocks_carry_the_scope_they_belong_to():
    """The settings' scope picker filters on this, and two sub-lanes both
    number their blocks from 1 - so the scope is part of what makes a
    block findable, not decoration."""
    row_sets = {
        "AEW": OpusRowSet(rates=[_row("G0001", "FAK")]),
        "JAPAN AEW": OpusRowSet(rates=[_row("G0011", "FAK - JAPAN")]),
    }

    blocks = commodity_blocks(row_sets)

    assert [(b.scope, b.key) for b in blocks] == [("AEW", "FAK"), ("JAPAN AEW", "FAK - JAPAN")]


class _StubParser:
    """Two sub-lanes with the same two commodity groups in each."""

    def run_multi(self, workbook, config):
        return {
            scope: OpusRowSet(rates=[_row("G0001", "FAK"), _row("G0002", "OOG")])
            for scope in ("OEW", "OMW")
        }


def test_the_pipeline_skips_a_group_in_one_scope_only():
    """Skip Filing is applied by the pipeline rather than by the parser,
    so it has to read each scope's own view too."""
    profile = MappingProfile(by_scope={"OMW": ScopeOverrides(skip_commodity_filing={"OOG": True})})

    row_sets = run_parser(_StubParser(), None, profile)

    assert {r.commodity_group_description for r in row_sets["OEW"].rates} == {"FAK", "OOG"}
    assert {r.commodity_group_description for r in row_sets["OMW"].rates} == {"FAK"}


def test_the_pipeline_orders_one_scope_differently():
    profile = MappingProfile(
        commodity_group_order=["FAK", "OOG"],
        by_scope={"OMW": ScopeOverrides(commodity_group_order=["OOG", "FAK"])},
    )

    row_sets = run_parser(_StubParser(), None, profile)

    assert [r.commodity_group_description for r in row_sets["OEW"].rates] == ["FAK", "OOG"]
    assert [r.commodity_group_description for r in row_sets["OMW"].rates] == ["OOG", "FAK"]


def test_folding_one_scope_back_does_not_drop_another():
    """The reset. Each scope's table is folded back on its own, and the
    fold used to rebuild the map from the profile every time - so the
    second scope's settings wiped the first's, and edits made in a scope
    then left behind by switching to the next were never in the profile
    to apply."""
    profile = MappingProfile(commodity_code_overrides={"FAK": "G0001"})

    by_scope = _with_scope_overrides(
        profile, "AEW", {"commodity_code_overrides": {"FAK #1": "G0009"}}, None
    )
    by_scope = _with_scope_overrides(
        profile, "AMW", {"commodity_code_overrides": {"FAK #1": "G0008"}}, by_scope
    )

    assert sorted(by_scope) == ["AEW", "AMW"]
    assert by_scope["AEW"].commodity_code_overrides == {"FAK #1": "G0009"}
    assert by_scope["AMW"].commodity_code_overrides == {"FAK #1": "G0008"}


def test_re_folding_a_scope_replaces_only_that_scope():
    profile = MappingProfile()
    by_scope = _with_scope_overrides(profile, "AEW", {"commodity_code_overrides": {"FAK": "X"}}, None)
    by_scope = _with_scope_overrides(profile, "AMW", {"commodity_code_overrides": {"FAK": "Y"}}, by_scope)

    by_scope = _with_scope_overrides(profile, "AEW", {"commodity_code_overrides": {"FAK": "Z"}}, by_scope)

    assert by_scope["AEW"].commodity_code_overrides == {"FAK": "Z"}
    assert by_scope["AMW"].commodity_code_overrides == {"FAK": "Y"}
