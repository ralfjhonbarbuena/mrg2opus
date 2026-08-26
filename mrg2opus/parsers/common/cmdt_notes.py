"""Shared "Rate structure: Includes X/Y/Z..." -> OPUS CMDT NOTE boilerplate
builder, used by every lane whose raw sheet has this pattern (confirmed
against SAF and EAF ground truth so far - same template shape, a couple of
lane-specific knobs below).
"""
from __future__ import annotations

import re
from datetime import date

from mrg2opus.schema.charge_codes import CHARGE_CODE_NAMES, INDIVIDUAL_CHARGE_CODES
from mrg2opus.schema.opus_rows import CmdtNoteRow

_INCLUDES_RE = re.compile(r"\b(?:incl\.|includes?)\s+([A-Z/]+)", re.IGNORECASE)


def parse_included_charge_codes(text: str) -> list[str]:
    """Extract charge codes from a raw 'Rate structure: Includes X/Y/Z, subj
    to ...' line, keeping only codes confirmed (via INDIVIDUAL_CHARGE_CODES)
    to actually get their own CMDT NOTE child row - some mentioned codes
    (e.g. EAF's BAF/BRS) never do, per ground truth."""
    m = _INCLUDES_RE.search(text)
    if not m:
        return []
    codes = [c.strip().upper() for c in m.group(1).split("/") if c.strip()]
    return sorted(c for c in codes if c in INDIVIDUAL_CHARGE_CODES)


def build_cmdt_notes(
    validity_start: date | None,
    validity_end: date | None,
    included_codes: list[str],
    sequential_charge_seq: bool = False,
    sort_text_names: bool = True,
    charge_code_names_override: dict[str, str] | None = None,
    excluded_codes: frozenset[str] = frozenset(),
) -> list[CmdtNoteRow]:
    """sequential_charge_seq: SAF's ground truth leaves child rows' Charge Seq
    blank (only the parent gets 1); EAF's numbers every row 1, 2, 3, ...
    Both are real, lane-specific ground-truth behaviors, not a guess.
    sort_text_names: whether the "inclusive of X and Y" text lists codes
    alphabetically (SAF/EAF/LAEC) or in the same order as included_codes
    (CSE) - the two conventions genuinely contradict each other across
    lanes, so this isn't a guessable default; pass False for CSE-style.
    charge_code_names_override: per-lane charge-code full-name overrides -
    e.g. LAEC's ground truth says "HEAVY SURCHARGE(HEA)" where EAF's says
    "HEAVY WEIGHT SURCHARGE(HEA)" for the same code; the shared
    CHARGE_CODE_NAMES can't hold two different names for one code, so a
    lane with its own confirmed wording passes it here instead.
    excluded_codes: user-directed, filing-wide charge codes to drop
    entirely (see MappingProfile.excluded_charge_codes) - e.g. a Hong Kong
    account excluding BAF because it duplicates OBS and isn't applicable
    for their RFAs. Applied before anything else so an excluded code never
    appears in the "inclusive of" text or gets its own child row."""
    if excluded_codes:
        included_codes = [c for c in included_codes if c not in excluded_codes]
    if not included_codes or validity_start is None or validity_end is None:
        return []

    # The child-row list can legitimately repeat a code (confirmed: CSE's
    # ground truth lists THL as two separate charge-seq child rows), but the
    # "inclusive of X and Y" text names each surcharge only once - dedupe
    # here while `included_codes` (and the children built from it below)
    # keeps every occurrence.
    unique_codes = list(dict.fromkeys(included_codes))
    if sort_text_names:
        unique_codes = sorted(unique_codes)
    names = {**CHARGE_CODE_NAMES, **(charge_code_names_override or {})}
    names_line = " and the ".join(f"{names.get(code, code)}({code})" for code in unique_codes)
    lines = [
        f"Rates are valid from {validity_start:%Y%m%d} to {validity_end:%Y%m%d}",
        f"Rates are inclusive of the {names_line}",
        "Rates are subject to all other surcharges, including those, if any, specified in "
        "the contract and those published in the Governing Tariff(s) at the time of shipment.",
    ]
    contents = "\n".join(lines)

    parent = CmdtNoteRow(
        contents=contents,
        charge_seq=1,
        code="APP",
        application_effective=validity_start,
        application_expires=validity_end,
        application="S",
    )
    children = [
        CmdtNoteRow(
            charge_seq=(i + 2) if sequential_charge_seq else None,
            code=code,
            application_effective=validity_start,
            application_expires=validity_end,
            application="I",
        )
        for i, code in enumerate(included_codes)
    ]
    return [parent, *children]
