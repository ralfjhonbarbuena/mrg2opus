"""Shared "Rate structure: Includes X/Y/Z..." -> OPUS CMDT NOTE boilerplate
builder, used by every lane whose raw sheet has this pattern (confirmed
against SAF and EAF ground truth so far - same template shape, a couple of
lane-specific knobs below).
"""
from __future__ import annotations

import re
from datetime import date

from mrg2opus.schema.charge_codes import CHARGE_CODE_NAMES, is_known_charge_code
from mrg2opus.schema.opus_rows import CmdtNoteRow

_INCLUDES_RE = re.compile(r"\b(?:incl\.|includes?)\s+([A-Z/]+)", re.IGNORECASE)


def parse_included_charge_codes(text: str) -> list[str]:
    """Extract charge codes from a raw 'Rate structure: Includes X/Y/Z, subj
    to ...' line. Every real charge code is recognized (gated only on
    having a known name - see charge_codes.py::is_known_charge_code);
    suppressing one this account shouldn't file is the user's call, via
    MappingProfile.excluded_charge_codes."""
    m = _INCLUDES_RE.search(text)
    if not m:
        return []
    codes = [c.strip().upper() for c in m.group(1).split("/") if c.strip()]
    return sorted(c for c in codes if is_known_charge_code(c))


def build_cmdt_notes(
    validity_start: date | None,
    validity_end: date | None,
    included_codes: list[str],
    sequential_charge_seq: bool = False,
    sort_text_names: bool = True,
    charge_code_names_override: dict[str, str] | None = None,
    excluded_codes: frozenset[str] = frozenset(),
    rfa_effective: date | None = None,
    rfa_expiry: date | None = None,
    service_lane: str | None = None,
    scope_values: list[str | None] | None = None,
    scope_field: str = "pol",
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
    appears in the "inclusive of" text or gets its own child row.
    rfa_effective/rfa_expiry: user-directed, filing-wide override for each
    CHILD row's own Application Effective/Expires dates (see
    MappingProfile.rfa_effective_date/rfa_expiry_date) - a charge code's
    real-world RFA (Rate Filing Agreement) window is usually a separate,
    longer-lived date pair a human filer enters, not the weekly rate
    validity window every child row defaults to below. The PARENT (APP)
    row always keeps the weekly validity window regardless - only
    children are affected. Either left None falls back to
    validity_start/validity_end for that one bound, same as before this
    parameter existed.
    service_lane: LAEC LUX-specific (confirmed against its own real ground
    truth) - when set, inserts an extra "Rates are applicable for Vessel
    Service Lane: {service_lane}" line right after the validity line, and
    stamps the parent row's own Lane column with the same value. None
    (default) leaves both untouched, so every other lane is unaffected.
    scope_values/scope_field: for a lane whose child rows carry a per-code
    scope (e.g. AUBP's POR-scoped THL/ISL/DOC - see aubp.py), pass a list
    positionally aligned with included_codes (None for an unscoped/blanket
    code) plus which CmdtNoteRow field to stamp it on ("por" or "pol").
    Left as None by default - every existing caller is unaffected."""
    if excluded_codes:
        keep = [c not in excluded_codes for c in included_codes]
        if scope_values is not None:
            scope_values = [v for v, k in zip(scope_values, keep) if k]
        included_codes = [c for c, k in zip(included_codes, keep) if k]
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
    lines = [f"Rates are valid from {validity_start:%Y%m%d} to {validity_end:%Y%m%d}"]
    if service_lane:
        lines.append(f"Rates are applicable for Vessel Service Lane: {service_lane}")
    lines.append(f"Rates are inclusive of the {names_line}")
    lines.append(
        "Rates are subject to all other surcharges, including those, if any, specified in "
        "the contract and those published in the Governing Tariff(s) at the time of shipment."
    )
    contents = "\n".join(lines)

    parent = CmdtNoteRow(
        contents=contents,
        charge_seq=1,
        code="APP",
        application_effective=validity_start,
        application_expires=validity_end,
        application="S",
        lane=service_lane,
    )
    child_effective = rfa_effective if rfa_effective is not None else validity_start
    child_expires = rfa_expiry if rfa_expiry is not None else validity_end
    children = [
        CmdtNoteRow(
            charge_seq=(i + 2) if sequential_charge_seq else None,
            code=code,
            application_effective=child_effective,
            application_expires=child_expires,
            application="I",
            **({scope_field: scope_values[i]} if scope_values is not None else {}),
        )
        for i, code in enumerate(included_codes)
    ]
    return [parent, *children]
