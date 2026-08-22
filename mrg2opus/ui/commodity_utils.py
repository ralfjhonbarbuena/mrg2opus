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
    """
    seen: dict[str, str] = {}
    for row_set in row_sets.values():
        for row in row_set.rates:
            seen.setdefault(row.commodity_group_description, row.commodity_group_code)
    return sorted((code, description) for description, code in seen.items())
