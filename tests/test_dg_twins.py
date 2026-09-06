"""Reefer/NOR dangerous twins - the opt-in any lane can turn on.

Only LAWC's ground truth files these, and its parser builds its own, so
these cover the shared rule the settings expose to every other lane (see
mrg2opus/parsers/common/dg_twins.py).
"""
from __future__ import annotations

from mrg2opus.parsers.common.dg_twins import (
    REEFER_DRY_AS_DANGEROUS,
    add_reefer_and_nor_dg_twins,
    reefer_and_nor_groups,
)
from mrg2opus.pipeline import run_parser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema.opus_rows import CmdtNoteRow, OpusRowSet, RatesRow

REEFER_GROUP = "FAK (RF)"
NOR_GROUP = "FAK (NOR)"
DRY_GROUP = "FAK"


def _row(group: str, prefix: str, cgo: str, destination: str = "AUSYD", **extra) -> RatesRow:
    return RatesRow(
        commodity_group_code="G0001", commodity_group_description=group,
        origin_code="CNSHA", origin_description="Shanghai",
        destination_code=destination, destination_description="Sydney",
        prefix=prefix, cgo_type=cgo, rate_40hc=1200, **extra,
    )


def _cgo_types(rows) -> set[tuple[str, str, str]]:
    return {(r.prefix, r.cgo_type, r.commodity_group_description) for r in rows}


def test_reefer_and_nor_groups_finds_both_and_leaves_dry_alone():
    row_set = OpusRowSet(rates=[
        _row(DRY_GROUP, "D", "DR"), _row(REEFER_GROUP, "R", "RF"), _row(NOR_GROUP, "R", "DR"),
    ])
    assert reefer_and_nor_groups({"": row_set}) == frozenset({REEFER_GROUP, NOR_GROUP})


def test_the_two_twins_are_different_shapes():
    """A reefer box carrying dry cargo is filed as dry when that cargo is
    dangerous, which is what the note explains - so NOR's twin changes
    prefix as well as CGO type, and Reefer's does not."""
    row_set = OpusRowSet(rates=[_row(REEFER_GROUP, "R", "RF"), _row(NOR_GROUP, "R", "DR")])

    out = add_reefer_and_nor_dg_twins(row_set, frozenset({REEFER_GROUP, NOR_GROUP}))

    assert _cgo_types(out.rates) == {
        ("R", "RF", REEFER_GROUP), ("R", "DG", REEFER_GROUP),
        ("R", "DR", NOR_GROUP), ("D", "DG", NOR_GROUP),
    }
    reefer_twin = next(r for r in out.rates if (r.prefix, r.cgo_type) == ("R", "DG"))
    nor_twin = next(r for r in out.rates if (r.prefix, r.cgo_type) == ("D", "DG"))
    assert reefer_twin.route_note is None
    assert nor_twin.route_note == REEFER_DRY_AS_DANGEROUS
    # Same rate, same route, same group as the row it came from.
    assert (reefer_twin.rate_40hc, reefer_twin.destination_code) == (1200, "AUSYD")


def test_the_nor_note_joins_a_note_the_row_already_had():
    row_set = OpusRowSet(rates=[
        _row(NOR_GROUP, "R", "DR", route_note="Rates are applicable for Vessel Service Lane: AX3"),
    ])

    out = add_reefer_and_nor_dg_twins(row_set, frozenset({NOR_GROUP}))

    twin = next(r for r in out.rates if (r.prefix, r.cgo_type) == ("D", "DG"))
    assert twin.route_note == "REEFER DRY AS DANGEROUS | Rates are applicable for Vessel Service Lane: AX3"


def test_a_group_that_was_not_turned_on_gets_nothing():
    row_set = OpusRowSet(rates=[_row(REEFER_GROUP, "R", "RF"), _row(NOR_GROUP, "R", "DR")])

    out = add_reefer_and_nor_dg_twins(row_set, frozenset({REEFER_GROUP}))

    assert _cgo_types(out.rates) == {
        ("R", "RF", REEFER_GROUP), ("R", "DG", REEFER_GROUP), ("R", "DR", NOR_GROUP),
    }
    assert add_reefer_and_nor_dg_twins(row_set, frozenset()).rates == row_set.rates


def test_a_twin_the_lane_already_filed_is_not_filed_twice():
    """LAWC builds its own, and OPUS rejects a duplicate filing outright -
    so turning the setting on for a lane that already does it has to be a
    no-op rather than a second copy of every row."""
    row_set = OpusRowSet(rates=[
        _row(REEFER_GROUP, "R", "RF"),
        _row(REEFER_GROUP, "R", "DG"),  # the lane's own twin
    ])

    out = add_reefer_and_nor_dg_twins(row_set, frozenset({REEFER_GROUP}))

    assert len(out.rates) == 2


class _StubParser:
    """Stands in for a lane parser: run_parser only calls run_multi."""

    def __init__(self, row_set: OpusRowSet) -> None:
        self._row_set = row_set

    def run_multi(self, workbook, config):
        return {"": self._row_set.model_copy(deep=True)}


def test_the_pipeline_leaves_reefer_and_nor_alone_unless_asked():
    """Absent means OFF for these, the opposite of a dry group's default -
    no lane but LAWC files them, and LAWC builds its own."""
    parser = _StubParser(OpusRowSet(rates=[_row(REEFER_GROUP, "R", "RF"), _row(NOR_GROUP, "R", "DR")]))

    out = run_parser(parser, None, MappingProfile())

    assert _cgo_types(out[""].rates) == {("R", "RF", REEFER_GROUP), ("R", "DR", NOR_GROUP)}


def test_turning_nor_on_files_the_route_note_on_the_rn_sheet_too():
    """A route note addresses an RN row by (header_seq, route_seq), so a
    note with no RN entry behind it says nothing."""
    parser = _StubParser(OpusRowSet(
        rates=[_row(NOR_GROUP, "R", "DR")],
        cmdt_notes=[CmdtNoteRow(header_seq=1, note_seq=1, contents="Rates are valid...", charge_seq=1, code="APP")],
    ))

    out = run_parser(parser, None, MappingProfile(skip_dg_generation={NOR_GROUP: False}))

    twin = next(r for r in out[""].rates if (r.prefix, r.cgo_type) == ("D", "DG"))
    notes = [n for n in out[""].route_notes if n.contents == REEFER_DRY_AS_DANGEROUS]
    assert len(notes) == 1
    assert (notes[0].header_seq, notes[0].route_seq) == (twin.cmdt_seq, twin.route_seq)


def test_the_rn_entry_is_not_added_twice_when_the_lane_wrote_it():
    parser = _StubParser(OpusRowSet(rates=[_row(REEFER_GROUP, "R", "RF")]))

    out = run_parser(parser, None, MappingProfile(skip_dg_generation={REEFER_GROUP: False}))

    # Reefer's twin carries no note at all, so nothing lands on RN.
    assert out[""].route_notes == []
