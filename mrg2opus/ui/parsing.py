"""Shared by step2_preview.py and step3_customize.py: runs a lane parser
and applies the user's chosen commodity-group order (if any) to the
result. Centralized here so the ordering step is never accidentally
skipped by one caller and not the other.
"""
from __future__ import annotations

from mrg2opus.parsers.base import BaseMRGParser
from mrg2opus.parsers.common.ordering import reorder_row_set
from mrg2opus.presets.models import MappingProfile


def run_parser(parser: BaseMRGParser, workbook, profile: MappingProfile) -> dict:
    row_sets = parser.run_multi(workbook, profile)
    if not profile.commodity_group_order:
        return row_sets
    return {suffix: reorder_row_set(row_set, profile.commodity_group_order) for suffix, row_set in row_sets.items()}
