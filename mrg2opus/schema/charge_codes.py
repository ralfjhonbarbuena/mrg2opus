"""Surcharge/charge code -> full display name, used to expand abbreviations
(e.g. "EFS" -> "EMERGENCY FUEL SURCHARGE") when generating OPUS CMDT NOTE
boilerplate text. Mined from the one worked example available (SAF); meant
to grow as more lanes are added - treat as user-editable reference data,
not an exhaustive authority.
"""
from __future__ import annotations

CHARGE_CODE_NAMES: dict[str, str] = {
    "BAF": "BUNKER ADJUSTMENT FACTOR",
    "EFS": "EMERGENCY FUEL SURCHARGE",
    "MBS": "MONTHLY ONE BUNKER SURCHARGE",
    "OBS": "ONE BUNKER SURCHARGE",
    "HEA": "HEAVY WEIGHT SURCHARGE",
    "LSF": "LOW SULPHUR FUEL SURCHARGE",
    "PSS": "PEAK SEASON SURCHARGE",
    "THL": "TERMINAL HANDLING CHARGE (L)",  # confirmed against CSE ground truth
    "CSS": "CARRIER SECURITY SURCHARGE",
    "SLF": "SEAL FEE",
    "CGD": "CONGESTION SURCHARGE (D)",  # confirmed against West Africa WAF ground truth
    # 3 literal spaces between PAYMENT/HANDLING, verbatim from ground truth text.
    "EPH": "ELSEWHERE PAYMENT   HANDLING FEE",
    "ISL": "INTERNATIONAL SECURITY FEE AT LOCAL",  # confirmed against AUS NEA to AUEC ground truth
}

# Codes confirmed (against ground truth) to become their own CMDT NOTE child
# row when mentioned in a raw sheet's "Rate structure: Includes X/Y/Z" line.
# Deliberately a whitelist, not "everything after Includes": EAF's raw text
# lists BRS as included too, but it never gets a child row in the ground
# truth - only extend this set when a new code is directly confirmed
# against a real OPUS CMDT NOTE sheet, not by guessing from raw text alone.
#
# CORRECTION (2026-08-26): BAF was originally excluded here too, based on
# the older bundled EAF.xlsx sample's ground truth. User-clarified reason:
# their SOP tells human filing agents NOT to file BAF - a special case for
# people, not a filing-format rule. This tool should reproduce whatever
# the raw MRG's "Includes" text says, BAF included, rather than mimicking
# that human-only SOP exclusion. Confirmed via reference/2_OPUS/7_EAF-KEMBA
# and 8_EAF-KEMBA, both of which do include it.
INDIVIDUAL_CHARGE_CODES: frozenset[str] = frozenset(
    {"BAF", "EFS", "MBS", "OBS", "HEA", "LSF", "PSS", "THL", "CSS", "SLF", "CGD", "EPH", "ISL"}
)
