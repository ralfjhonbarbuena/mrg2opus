"""Surcharge/charge code -> full display name, used to expand abbreviations
(e.g. "EFS" -> "EMERGENCY FUEL SURCHARGE") when generating OPUS CMDT NOTE
boilerplate text. Mined from the one worked example available (SAF); meant
to grow as more lanes are added - treat as user-editable reference data,
not an exhaustive authority.
"""
from __future__ import annotations

CHARGE_CODE_NAMES: dict[str, str] = {
    "EFS": "EMERGENCY FUEL SURCHARGE",
    "MBS": "MONTHLY ONE BUNKER SURCHARGE",
    "OBS": "ONE BUNKER SURCHARGE",
    "HEA": "HEAVY WEIGHT SURCHARGE",
    "LSF": "LOW SULPHUR FUEL SURCHARGE",
    "PSS": "PEAK SEASON SURCHARGE",
    "THL": "TERMINAL HANDLING CHARGE (L)",  # confirmed against CSE ground truth
    "CSS": "CARRIER SECURITY SURCHARGE",
    "SLF": "SEAL FEE",
}

# Codes confirmed (against ground truth) to become their own CMDT NOTE child
# row when mentioned in a raw sheet's "Rate structure: Includes X/Y/Z" line.
# Deliberately a whitelist, not "everything after Includes": EAF's raw text
# lists BAF and BRS as included too, but neither gets a child row in the
# ground truth - only extend this set when a new code is directly confirmed
# against a real OPUS CMDT NOTE sheet, not by guessing from raw text alone.
INDIVIDUAL_CHARGE_CODES: frozenset[str] = frozenset({"EFS", "MBS", "OBS", "HEA", "LSF", "PSS", "THL", "CSS", "SLF"})
