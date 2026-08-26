"""Shared by step2_preview.py (to snapshot the parser's own default codes
right after the first parse) and step3_customize.py (to render/apply the
override editor) - kept out of either step module to avoid one importing
a "private" helper from the other.
"""
from __future__ import annotations


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
