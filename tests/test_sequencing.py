"""Cross-sheet sequence fixups - the pipeline pass that runs after every
lane has finished building (mrg2opus/parsers/common/sequencing.py).

Covers the three bugs reported 2026-08-31: PORT-PORT rows came out with no
CMDT Seq at all, CMDT NOTE repeated one Header Seq down every row of a
block, and Note Seq was never filled in.
"""
from __future__ import annotations

from mrg2opus.parsers.common.sequencing import (
    collapse_note_block_sequences,
    finalize_sequences,
    sync_port_port_cmdt_seq,
)
from mrg2opus.schema.opus_rows import (
    CmdtNoteRow,
    OpusRowSet,
    RatesRow,
    SpecialNoteRow,
    explode_rates_row,
)


def _rates_row(group: str, origin: str = "CNSHA;CNNGB", dest: str = "AEJEA") -> RatesRow:
    return RatesRow(
        commodity_group_code="G0001",
        commodity_group_description=group,
        origin_code=origin,
        origin_description=origin,
        destination_code=dest,
        destination_description=dest,
        prefix="D",
        cgo_type="DR",
    )


def _note(contents: str | None, charge_seq: int, header_seq: int | None) -> CmdtNoteRow:
    return CmdtNoteRow(contents=contents, charge_seq=charge_seq, header_seq=header_seq)


def test_port_port_takes_its_source_row_cmdt_seq():
    row = _rates_row("Dry")
    port_port = explode_rates_row(row)  # exploded BEFORE the number exists
    assert len(port_port) == 2
    row.cmdt_seq = 4

    sync_port_port_cmdt_seq([row], port_port)

    assert [r.cmdt_seq for r in port_port] == [4, 4]
    # ...and the commodity code was already carried through the explode.
    assert {r.commodity_group_code for r in port_port} == {"G0001"}


def test_port_port_keyed_on_source_group_survives_a_group_remap():
    """LAWC rewrites commodity_group_description on its PORT-PORT rows, so
    the exploded row's own description can't be the lookup key."""
    dry, nor = _rates_row("Dry"), _rates_row("NOR")
    dry.cmdt_seq, nor.cmdt_seq = 1, 2
    port_port = explode_rates_row(dry) + explode_rates_row(nor)
    for r in port_port:
        r.commodity_group_description = "FAK - COMBINED"  # the remap

    sync_port_port_cmdt_seq([dry, nor], port_port)

    assert [r.cmdt_seq for r in port_port] == [1, 1, 2, 2]


def test_port_port_row_from_an_unknown_group_is_left_alone():
    orphan = explode_rates_row(_rates_row("Gone"))[0]
    orphan.cmdt_seq = 9
    sync_port_port_cmdt_seq([_rates_row("Dry")], [orphan])
    assert orphan.cmdt_seq == 9


def test_header_seq_survives_on_the_parent_row_only():
    notes = [_note("Rates are valid from...", 1, 7), _note(None, 2, 7), _note("", 3, 7)]

    collapse_note_block_sequences(notes)

    assert [n.header_seq for n in notes] == [7, None, None]


def test_note_seq_is_1_beside_the_header_seq():
    notes = [_note("Rates are valid from...", 1, 7), _note(None, 2, 7)]

    collapse_note_block_sequences(notes)

    assert [n.note_seq for n in notes] == [1, None]


def test_each_block_gets_its_own_parent_row():
    notes = [
        _note("block one", 1, 1), _note(None, 2, 1),
        _note("block two", 1, 2), _note(None, 2, 2),
    ]

    collapse_note_block_sequences(notes)

    assert [n.header_seq for n in notes] == [1, None, 2, None]
    assert [n.note_seq for n in notes] == [1, None, 1, None]


def test_a_lane_specific_note_seq_is_not_overwritten():
    """CSE files its SPECIAL NOTE parent under 3, LAEC's under 1 - both
    confirmed against their own ground truth, so 1 is only a default."""
    notes = [_note("valid from...", 1, 1)]
    notes[0].note_seq = 3

    collapse_note_block_sequences(notes)

    assert notes[0].note_seq == 3


def test_notes_with_no_contents_anywhere_are_left_untouched():
    notes = [_note(None, 1, 5), _note(None, 2, 5)]

    collapse_note_block_sequences(notes)

    assert [n.header_seq for n in notes] == [5, 5]


def test_finalize_sequences_covers_cmdt_and_special_notes_and_port_port():
    row = _rates_row("Dry")
    row_set = OpusRowSet(
        rates=[row],
        rates_port_port=explode_rates_row(row),
        cmdt_notes=[_note("cmdt block", 1, 1), _note(None, 2, 1)],
        special_notes=[
            SpecialNoteRow(contents="special block", charge_seq=1, header_seq=2),
            SpecialNoteRow(contents=None, charge_seq=2, header_seq=2),
        ],
    )
    row.cmdt_seq = 1

    finalize_sequences(row_set)

    assert [r.cmdt_seq for r in row_set.rates_port_port] == [1, 1]
    assert [n.header_seq for n in row_set.cmdt_notes] == [1, None]
    assert [n.note_seq for n in row_set.special_notes] == [1, None]
