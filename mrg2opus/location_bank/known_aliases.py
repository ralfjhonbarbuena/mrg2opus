"""Curated raw-text -> code aliases for real spelling/naming variants that
score just under the fuzzy-match confidence threshold against the mined
primary name (see fuzzy_match.DEFAULT_THRESHOLD). Each entry here has been
manually confirmed against the OPUS ground truth - this is NOT a substitute
for real Location Bank coverage (UN/LOCODE or Step 3 manual overrides), just
a small seed for variants already verified while building the SAF parser.

Loaded by bootstrap_from_samples.bootstrap() as source="manual_override".
"""
from __future__ import annotations

KNOWN_ALIASES: dict[str, str] = {
    "Busan": "KRPUS",  # raw MRG spelling vs. mined name "PUSAN" (older romanization)
    "Lakbang": "THLKR",  # raw MRG spelling vs. mined name "LAT KRABANG"
    "HoChiMing": "VNSGN",  # EAF KEMBA raw sheet typo for "HoChiMinh" (Ho Chi Minh City)
    # NZ1 SEA raw sheet's own "Kolkata / Calcutta" combined label - "Calcutta"
    # alone fuzzy-matches an unrelated location (score 62.5, needs_review)
    # since it isn't a substring of mined primary name "KOLKATA".
    "Calcutta": "INCCU",
}
