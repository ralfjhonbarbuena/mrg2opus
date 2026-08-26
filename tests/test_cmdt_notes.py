from __future__ import annotations

from datetime import date

from mrg2opus.parsers.common.cmdt_notes import build_cmdt_notes


def test_build_cmdt_notes_excludes_requested_code():
    """excluded_charge_codes (MappingProfile) lets a user drop a code from
    both the "inclusive of" text and its own child row, filing-wide - e.g.
    a Hong Kong account excluding BAF because it duplicates OBS and isn't
    applicable for their RFAs, even though the raw MRG mentions it."""
    notes = build_cmdt_notes(
        date(2026, 8, 15), date(2026, 8, 21), ["BAF", "EFS", "OBS"], excluded_codes=frozenset({"BAF"})
    )

    codes = [n.code for n in notes]
    assert "BAF" not in codes
    assert codes == ["APP", "EFS", "OBS"]
    assert "BAF" not in notes[0].contents


def test_build_cmdt_notes_excluding_every_code_yields_nothing():
    notes = build_cmdt_notes(date(2026, 8, 15), date(2026, 8, 21), ["BAF"], excluded_codes=frozenset({"BAF"}))
    assert notes == []


def test_build_cmdt_notes_default_excludes_nothing():
    notes = build_cmdt_notes(date(2026, 8, 15), date(2026, 8, 21), ["BAF", "EFS"])
    assert [n.code for n in notes] == ["APP", "BAF", "EFS"]
