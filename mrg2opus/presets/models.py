"""MappingProfile: the Step 3 "customization" contract. A saved/reloadable
named profile (e.g. "Asia-Europe Standard Profile") so recurring lanes don't
need re-configuring every run.
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal
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
    # default_description -> True to skip generating that group's DG
    # (Dangerous Goods) duplicate rows. Every lane's base Dry (D/DR) row
    # normally also files an identical D/DG variant at the same rate - a
    # standing filing convention, not derived from the raw sheet (see each
    # parser's own comment at its `cgo_type == "DR"` check). Absent/False
    # keeps that default behavior; True suppresses it for that one group.
    skip_dg_generation: dict[str, bool] = Field(default_factory=dict)
    # default_description -> True to leave that commodity group out of the
    # filing ALTOGETHER - every one of its RATES / RATES PORT-PORT /
    # VERTICAL RATES rows and its whole CMDT NOTE block. Distinct from
    # skip_dg_generation above, which only drops a group's D/DG duplicate
    # and keeps the group itself; and from skip_output_sheets, which drops
    # a whole SHEET across every group. Applied in pipeline.py::
    # run_parser() rather than in each parser, so it works the same on all
    # lanes. The remaining groups keep the CMDT Seq numbers the parser
    # gave them, so skipping a middle group leaves a gap (1, 3, 4) -
    # deliberate, since renumbering would silently move groups the user
    # had pinned with an explicit sequence; the CMDT Seq column is there
    # to close a gap by hand if a filing needs it.
    skip_commodity_filing: dict[str, bool] = Field(default_factory=dict)
    # Charge codes to drop from every CMDT NOTE/SPECIAL NOTE-equivalent
    # block across the whole filing (not per-group - see
    # project-tool-mirrors-mrg-not-human-sop memory) - e.g. a Hong Kong
    # account excluding "BAF" because it duplicates OBS and isn't
    # applicable for their RFAs, even though the raw MRG's own "Includes"
    # text mentions it. Applied uniformly to every lane's charge-code
    # list (see parsers/common/cmdt_notes.py::build_cmdt_notes).
    excluded_charge_codes: list[str] = Field(default_factory=list)
    # Individual charge codes' (CMDT NOTE child rows') Application
    # Effective/Expires dates normally just mirror the weekly rate
    # validity window (see cmdt_notes.py::build_cmdt_notes) - but ground
    # truth shows they're actually meant to be the charge code's own RFA
    # (Rate Filing Agreement) window, a separate, usually longer-lived
    # date pair a human filer enters per their account's own RFA (e.g.
    # West Africa WAF's ground truth: parent/APP row uses that week's
    # validity, every child code uses a much longer 2026-05-20 to
    # 2026-12-31 window instead). Applied filing-wide (not per-group,
    # same scope choice as excluded_charge_codes above) - when unset
    # (None), children keep falling back to the rate validity window,
    # same as before this field existed.
    rfa_effective_date: Optional[date] = None
    rfa_expiry_date: Optional[date] = None
    # Adds an OPUS VERTICAL RATES sheet (a "long format" alternate upload -
    # one row per container size instead of RATES' 4-slots-per-row layout;
    # per the user, faster to upload but capped at 10,000 rows per file).
    # ON by default for every lane (user-directed 2026-08-30, "add the
    # vertical rates for all MRGs") - see schema/opus_rows.py::
    # build_vertical_rates(), applied in pipeline.py::run_parser() so the
    # CLI and the UI both get it.
    #
    # NOTE the 10,000-row cap is a real limit this can exceed: exploding
    # RATES into one row per container size roughly triples the row count,
    # so a large lane overruns it (CSE ~11.9k, LAWC ~19.1k rows on their
    # own reference files). Nothing here truncates or splits - the count is
    # surfaced instead (CLI output, and a Step 2 warning) so the filer can
    # decide, since how to split a filing is their call, not the tool's.
    include_vertical_rates: bool = True
    # TAD FILING lanes only (OEW/OMW, WMW/WEW, AEW/AMW): mirrors the team's
    # own "Tool for TAD.xlsm" VBA export's "Include Dry Dangerous" toggle -
    # duplicates every D/DR row as an identical D/DG row at the same rate.
    # Unlike skip_dg_generation above (which SKIPS an otherwise-default-on
    # duplicate), TAD's own confirmed convention is DG duplication OFF by
    # default - this is an opt-IN, not an opt-out, and applies uniformly
    # across every TAD lane in one run (no per-group key: TAD's raw sheets
    # don't have the same "default_description per group" structure the
    # other lanes key skip_dg_generation by).
    generate_tad_dg_duplicate: bool = False
    # AEW/AMW only: an OFT 45 ("D7") rate slot the raw MRG never carries
    # directly - confirmed against ground truth as OFT 40HC + a fixed
    # add-on (reference/1_MRGs/23_TAD FILING AEW AMW's own "Surcharges"
    # sheet, rows 48-49: "$700 add-on for D7"). Off by default (matching
    # the VBA tool's own "Include D7 (OFT 45)" toggle default) and only
    # ever applied to D/DR rows - a generated D/DG duplicate (see
    # generate_tad_dg_duplicate above) copies the same D7 rate across
    # rather than recomputing it. OEW/OMW and WMW/WEW never generate this;
    # the raw MRG shape doesn't carry an equivalent add-on for them.
    include_tad_d7: bool = False
    tad_d7_addon: Decimal = Decimal("700")
