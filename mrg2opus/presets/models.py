"""MappingProfile: the Step 3 "customization" contract. A saved/reloadable
named profile (e.g. "Asia-Europe Standard Profile") so recurring lanes don't
need re-configuring every run.
"""
from __future__ import annotations

from typing import Optional

from pydantic import BaseModel, Field


class MappingProfile(BaseModel):
    name: str = "default"
    created_by: Optional[str] = None
    updated_at: Optional[str] = None

    # All three commodity_*_overrides dicts below are keyed by the SAME
    # thing: a commodity group's DEFAULT (override-free) description - not
    # its code. Every raw sheet defaults to its own description (see
    # parsers/common/commodity.py), which is guaranteed unique per group,
    # unlike code (several sheets can share one default code, e.g. LAWC's
    # main dry grid/"Reefer"/"LAWC NOR" all default to G0001 - keying by
    # code there would collide their overrides into one). The UI captures
    # these default descriptions once, before any override is applied
    # (WizardState.default_commodity_groups), so they stay valid lookup
    # keys across repeated rounds of editing.
    #
    # default_description -> override code. Codes should mainly come from
    # the user, not a mined bank (there is no commodity_bank/ - see
    # README) - a parser's hardcoded code is only ever the starting
    # suggestion. A lane's internal logic that depends on its own
    # hardcoded code (charge-code selection, the RATES-vs-PORT-PORT code
    # remap in lawc.py, etc.) keeps working regardless of what the user
    # renames the output to, since only the OUTPUT code is swapped.
    commodity_code_overrides: dict[str, str] = Field(default_factory=dict)
    # default_description -> override description. Every lane ships a
    # sensible default description, but it's always user-editable. Two (or
    # more) groups overridden to the exact same description merge into one
    # CMDT NOTE block - see parsers/common/commodity.py::build_notes_by_description().
    commodity_description_overrides: dict[str, str] = Field(default_factory=dict)
    # default_description -> override cmdt_seq
    commodity_sequence_overrides: dict[str, int] = Field(default_factory=dict)
    # The FINAL (post-override) commodity group descriptions, in the order
    # the user wants their blocks to appear in the generated OPUS RATES /
    # RATES PORT-PORT / CMDT NOTE sheets - set by Step 3's "Order" column.
    # Groups not listed here (or a lane run without this set at all) keep
    # their parser-default relative order, appended after every group that
    # IS listed. See parsers/common/ordering.py::reorder_row_set().
    commodity_group_order: list[str] = Field(default_factory=list)
    # OPUS sheet name -> True to skip writing it to the output workbook
    skip_output_sheets: dict[str, bool] = Field(default_factory=dict)
