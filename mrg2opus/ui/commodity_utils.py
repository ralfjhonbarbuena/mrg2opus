"""Shared by step2_preview.py (to snapshot the parser's own default codes
right after the first parse) and step3_customize.py (to render/apply the
override editor) - kept out of either step module to avoid one importing
a "private" helper from the other.
"""
from __future__ import annotations

from mrg2opus.presets.models import MappingProfile


def distinct_commodity_groups(row_sets: dict) -> list[tuple[str, str]]:
    """Called on the FIRST (override-free) parse, so each row's
    commodity_group_description here IS that group's default description -
    the identity every MappingProfile.commodity_*_overrides dict is keyed
    by (see parsers/common/commodity.py's module docstring). Deduped by
    DESCRIPTION, not code: several groups can share one default code (e.g.
    LAWC's main dry grid/"Reefer"/"LAWC NOR" all default to G0001), so
    deduping by code would silently drop all but one of them from the
    Step 3 editor.

    Order is first-encounter (not sorted) - see
    assign_sequential_default_codes(), which numbers groups G0001, G0002,
    ... in this same order as the new out-of-the-box default, replacing
    whatever structural code (shown here) the parser's own internal
    joins happen to use.
    """
    seen: dict[str, str] = {}
    for row_set in row_sets.values():
        for row in row_set.rates:
            seen.setdefault(row.commodity_group_description, row.commodity_group_code)
    return [(code, description) for description, code in seen.items()]


def assign_sequential_default_codes(groups: list[tuple[str, str]]) -> dict[str, str]:
    """Every distinct commodity group gets its OWN unique output code by
    default - G0001, G0002, G0003, ... in the order groups were first
    encountered while parsing (per distinct_commodity_groups' order) -
    instead of silently sharing whatever structural code a lane's parser
    happens to use internally for unrelated joins (e.g. LAWC's main dry
    grid/Reefer/LAWC NOR all default to the same internal "G0001", but
    each should still get its own distinct OUTPUT code unless the user
    deliberately merges them). User-directed (2026-08-27): before this,
    unoverridden sibling groups would silently write the same code to the
    OPUS output. Returns {description: new_code}, meant to auto-seed
    MappingProfile.commodity_code_overrides right after the first parse -
    see step2_preview.py - so the user's own further overrides (Step 3)
    apply on top of this, exactly like any other override."""
    return {description: f"G{i + 1:04d}" for i, (_structural_code, description) in enumerate(groups)}


class _DGTwinProbe(dict):
    """A skip_dg_generation dict that remembers which groups were looked up.

    Every parser decides whether to build a group's D/DG (or R/DG) twin
    by asking `config.skip_dg_generation.get(<that group's key>, False)`,
    and it only asks where a twin is actually on offer - LAWC never asks
    about OOG, and the guards in front of those lookups mean it doesn't
    ask about a section this particular file left empty either. So the
    set of keys asked about IS the set of groups the checkbox means
    anything for, read from the code that implements the rule rather
    than from a second list of lanes that would drift from it.

    Deriving this from the OUTPUT instead would get it wrong in both
    directions: WAF puts its twins in a separate "- DG" commodity group,
    so the group producing DG rows is not the group whose checkbox
    suppresses them.
    """

    def __init__(self, existing: dict | None = None) -> None:
        # Seeded from whatever the profile already held, so parsing with
        # the probe gives exactly the output parsing without it would -
        # a preset that already skips a group keeps skipping it.
        super().__init__(existing or {})
        self.asked: set[str] = set()

    def get(self, key, default=None):  # noqa: D102 - dict.get
        self.asked.add(key)
        return super().get(key, default)


def dg_twin_probe(profile: MappingProfile) -> MappingProfile:
    """`profile` with its skip_dg_generation swapped for a probe.

    Parse with this and the output is identical (the probe is empty, so
    every lookup answers "don't skip"); read the answer back with
    groups_offering_dg_twins(). model_copy skips validation, which is
    what lets the subclass survive being assigned to a pydantic field.
    """
    return profile.model_copy(update={"skip_dg_generation": _DGTwinProbe(profile.skip_dg_generation)})


def groups_offering_dg_twins(probe: MappingProfile) -> frozenset[str]:
    """The groups a profile from dg_twin_probe() saw asked about, after
    parsing with it. Empty for a profile that was never used as a probe,
    which reads as "no group has a twin" - correct for TAD lanes, whose
    DG duplicate is one lane-wide flag rather than a per-group dict."""
    asked = getattr(probe.skip_dg_generation, "asked", frozenset())
    return frozenset(asked)
