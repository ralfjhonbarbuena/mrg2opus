"""Resolve OUTPUT commodity_group_code/description, and merge CMDT NOTE
blocks for sub-groups that end up sharing one description.

Every lane parser still hardcodes a default/structural commodity identity
per raw sheet (e.g. LAWC's COMMODITY_MAIN = ("G0001", ...)) because a
group's CODE is used as an internal join key throughout each parser -
charge-code selection, ISC/non-ISC branching, LAWC's RATES-vs-PORT-PORT
code remap (PP_COMMODITY) - not just as display text. But both the code
and the description actually written to the OPUS output should mainly
come from the user (there is no commodity_bank/ registry; see README), so
MappingProfile.commodity_code_overrides / commodity_description_overrides /
commodity_sequence_overrides swap in the user's chosen values only at the
point a row is built.

All three override dicts are keyed by the group's DEFAULT description
(what the group shows when nothing is overridden) - NOT by code. This is
the one identity that's both always unique per group (guaranteed by
construction: every raw sheet defaults to its own description, see below)
and directly visible to the UI (WizardState.default_commodity_groups is
captured from the very first, override-free parse, so it always holds the
true default descriptions to key against, even across repeated rounds of
editing). Code can't serve this role: several groups can share one default
code (e.g. LAWC's main dry grid, "Reefer", and "LAWC NOR" all default to
G0001), so keying by code would collide two groups' overrides into one.

A raw sheet that used to share ANOTHER sheet's description by default
(e.g. LAWC's "Reefer"/"LAWC NOR" sheets used to be folded into the main
dry grid's single combined description) now defaults to its OWN
description - its own raw sheet name - and gets its OWN CMDT NOTE block.
If the user overrides two (or more) sub-groups to the exact same
description, they're meant to collapse back into ONE CMDT NOTE block
(charge codes unioned) - see merge_note_specs()/build_notes_by_description()
below, used by every lane that has this kind of sub-group split (see
lawc.py, cse.py, laec.py for the concrete per-lane wiring).
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import date

from mrg2opus.parsers.common.cmdt_notes import build_cmdt_notes
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema.opus_rows import CmdtNoteRow


def resolve_commodity_code(default_description: str, default_code: str, config: MappingProfile) -> str:
    """default_description is the group's default (override-free)
    description - the lookup key for every override dict, see module
    docstring; default_code is the value used when there's no override,
    which is often a code SHARED with other groups (e.g. Reefer/NOR both
    default to their main grid's code) even though the lookup key isn't."""
    return config.commodity_code_overrides.get(default_description, default_code)


def resolve_commodity_description(default_description: str, config: MappingProfile) -> str:
    return config.commodity_description_overrides.get(default_description, default_description)


@dataclass(frozen=True)
class CommodityNoteSpec:
    description: str
    validity_start: date | None
    validity_end: date | None
    charge_codes: list[str]


def merge_note_specs(specs: list[CommodityNoteSpec]) -> dict[str, CommodityNoteSpec]:
    """Sub-groups that resolve to the SAME description (after any user
    override) share exactly one CMDT NOTE block - merges by unioning
    charge codes (order-preserving, deduped) and keeping the first spec's
    validity window (only relevant if the user deliberately merges two
    groups with genuinely different raw validity periods - there's no
    single "right" answer for that, so the first one wins)."""
    merged: dict[str, CommodityNoteSpec] = {}
    for spec in specs:
        if spec.description not in merged:
            merged[spec.description] = spec
        else:
            existing = merged[spec.description]
            combined_codes = list(existing.charge_codes)
            for code in spec.charge_codes:
                if code not in combined_codes:
                    combined_codes.append(code)
            merged[spec.description] = replace(existing, charge_codes=combined_codes)
    return merged


def build_notes_by_description(
    specs: list[CommodityNoteSpec], **build_kwargs
) -> tuple[list[CmdtNoteRow], dict[str, str | None]]:
    """Runs merge_note_specs() then calls build_cmdt_notes() once per
    distinct (post-merge) description, in first-seen order. Returns the
    flat list of all CMDT NOTE rows plus {description: note_contents_text}
    for stamping onto each row's commodity_note field. build_kwargs are
    passed straight through to build_cmdt_notes() (sequential_charge_seq,
    charge_code_names_override, etc.) - uniform per lane, not per group."""
    merged = merge_note_specs(specs)
    all_notes: list[CmdtNoteRow] = []
    note_text_by_description: dict[str, str | None] = {}
    seen: set[str] = set()
    for spec in specs:
        if spec.description in seen:
            continue
        seen.add(spec.description)
        m = merged[spec.description]
        notes = build_cmdt_notes(m.validity_start, m.validity_end, m.charge_codes, **build_kwargs)
        note_text_by_description[spec.description] = notes[0].contents if notes else None
        # Tag every row (parent AND children) with the group it belongs to
        # - see CmdtNoteRow.group_description's docstring.
        all_notes.extend(row.model_copy(update={"group_description": spec.description}) for row in notes)
    return all_notes, note_text_by_description
