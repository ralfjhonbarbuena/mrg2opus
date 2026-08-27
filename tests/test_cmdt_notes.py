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


def test_build_cmdt_notes_default_children_use_rate_validity():
    """Absent an rfa_effective/rfa_expiry override, children fall back to
    the weekly rate validity window - the pre-existing behavior, unchanged
    by adding the new parameter."""
    parent, *children = build_cmdt_notes(date(2026, 8, 15), date(2026, 8, 21), ["BAF", "EFS"])
    assert parent.application_effective == date(2026, 8, 15)
    assert parent.application_expires == date(2026, 8, 21)
    for child in children:
        assert child.application_effective == date(2026, 8, 15)
        assert child.application_expires == date(2026, 8, 21)


def test_build_cmdt_notes_rfa_override_applies_to_children_only():
    """rfa_effective/rfa_expiry represent the charge code's own RFA (Rate
    Filing Agreement) window - a separate, usually longer-lived date pair
    a human filer enters, distinct from the weekly rate validity. Only
    CHILD rows are affected; the parent (APP) row always keeps the rate
    validity window."""
    parent, *children = build_cmdt_notes(
        date(2026, 8, 15),
        date(2026, 8, 21),
        ["BAF", "EFS"],
        rfa_effective=date(2026, 5, 20),
        rfa_expiry=date(2026, 12, 31),
    )
    assert parent.application_effective == date(2026, 8, 15)
    assert parent.application_expires == date(2026, 8, 21)
    for child in children:
        assert child.application_effective == date(2026, 5, 20)
        assert child.application_expires == date(2026, 12, 31)


def test_build_cmdt_notes_rfa_override_bounds_are_independent():
    """Either bound can be overridden without the other - each falls back
    to the corresponding rate validity date independently."""
    _, child = build_cmdt_notes(
        date(2026, 8, 15), date(2026, 8, 21), ["BAF"], rfa_effective=date(2026, 5, 20)
    )
    assert child.application_effective == date(2026, 5, 20)
    assert child.application_expires == date(2026, 8, 21)  # fell back to validity_end
