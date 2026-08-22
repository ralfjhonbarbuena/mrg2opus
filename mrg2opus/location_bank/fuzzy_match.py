"""Resolve raw MRG free-text location names to Location Bank codes.

Raw origin cells are often a comma-separated list of city names sharing one
rate (e.g. "Hong Kong,Shekou,Nansha"). Each comma-separated token is matched
independently against the bank's exact aliases first, then primary names via
rapidfuzz. Anything below the confidence threshold is flagged needs_review
rather than silently accepted - this is billing-relevant data.
"""
from __future__ import annotations

import re

from rapidfuzz import fuzz, process
from rapidfuzz.utils import default_process

from mrg2opus.location_bank.models import LocationMatch, LocationRecord
from mrg2opus.location_bank.store import LocationBankStore

DEFAULT_THRESHOLD = 85.0

_VIA_RE = re.compile(r"\s+via\s+.*$", re.IGNORECASE)
_INLINE_CODE_RE = re.compile(r"\(([A-Z]{2}[A-Z0-9]{3})\)")


def split_location_text(raw_text: str) -> list[str]:
    """Split a raw origin/destination cell into individual location tokens.

    Drops a trailing "via <transship port>" clause (e.g. "Ganzhou via
    Yantian") - that names routing, not an additional origin location, and
    the OPUS ground truth confirms O.Via/D.Via stay blank for these rows
    rather than resolving the via-port as a second origin.
    """
    parts = [p.strip() for p in raw_text.replace("/", ",").split(",")]
    cleaned = []
    for p in parts:
        p = _VIA_RE.sub("", p).strip()
        if p:
            cleaned.append(p)
    return cleaned


class LocationResolver:
    def __init__(self, store: LocationBankStore | None = None, threshold: float = DEFAULT_THRESHOLD):
        self.store = store or LocationBankStore()
        self.threshold = threshold
        self._locations: list[LocationRecord] = self.store.all_locations()
        self._choices: dict[str, str] = {loc.primary_name: loc.code for loc in self._locations}

    def match_token(self, token: str) -> LocationMatch | None:
        token = token.strip()
        if not token:
            return None

        inline_code = _INLINE_CODE_RE.search(token)
        if inline_code:
            code = inline_code.group(1)
            loc = self.store.get_by_code(code)
            if loc is not None:
                return LocationMatch(code=loc.code, primary_name=loc.primary_name, score=100.0, needs_review=False)
            token = _INLINE_CODE_RE.sub("", token).strip()

        exact_aliases = self.store.get_by_alias(token)
        if exact_aliases:
            loc = exact_aliases[0]
            return LocationMatch(code=loc.code, primary_name=loc.primary_name, score=100.0, needs_review=False)

        for loc in self._locations:
            if loc.primary_name.lower() == token.lower() or loc.code.lower() == token.lower():
                return LocationMatch(code=loc.code, primary_name=loc.primary_name, score=100.0, needs_review=False)

        if not self._choices:
            return None

        best = process.extractOne(token, self._choices.keys(), scorer=fuzz.WRatio, processor=default_process)
        if best is None:
            return None
        name, score, _ = best
        code = self._choices[name]
        return LocationMatch(code=code, primary_name=name, score=score, needs_review=score < self.threshold)

    def match_text(self, raw_text: str) -> list[LocationMatch]:
        """Split raw_text on commas/slashes and match each token."""
        matches = []
        for token in split_location_text(raw_text):
            m = self.match_token(token)
            if m is not None:
                matches.append(m)
        return matches
