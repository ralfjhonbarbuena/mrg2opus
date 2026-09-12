"""Commodity BLOCKS: the unit the settings act on.

For most lanes a block and a commodity group are the same thing. TAD
groups its rows by their own validity window and Include Surcharge list,
so one name covers several blocks - four in a September AEW filing, two
snapshots crossed with two surcharge lists, all called "FAK".
"""
from __future__ import annotations

from datetime import date

from mrg2opus.parsers.common.blocks import block_key, describes_block, resolve_for_block
from mrg2opus.parsers.common.ordering import drop_commodity_groups
from mrg2opus.schema.opus_rows import CmdtNoteRow, OpusRowSet, RatesRow
from mrg2opus.ui.commodity_utils import commodity_blocks


def _row(description="FAK", seq=1, code="G0001", **over):
    return RatesRow(
        commodity_group_code=code, commodity_group_description=description, cmdt_seq=seq,
        origin_code="CNSHA", origin_description="Shanghai",
        destination_code="BEANR", destination_description="Antwerp",
        prefix="D", cgo_type="DR", **over,
    )


# --- naming a block -----------------------------------------------------------

def test_a_group_that_is_one_block_keeps_its_plain_name():
    """Every lane but TAD, so nothing about their settings - or the
    presets holding them - changes."""
    assert block_key("FAK", 1, blocks_in_group=1) == "FAK"
    assert block_key("LAWC NOR", None, blocks_in_group=1) == "LAWC NOR"


def test_a_group_with_several_blocks_names_each_one():
    assert block_key("FAK", 2, blocks_in_group=4) == "FAK #2"


def test_a_key_for_the_group_still_describes_every_block_of_it():
    """Which is what lets one setting cover the whole group, and what
    keeps a preset written before blocks were separable working."""
    assert describes_block("FAK", "FAK", 2)
    assert describes_block("FAK #2", "FAK", 2)
    assert not describes_block("FAK #3", "FAK", 2)


def test_the_block_wins_over_the_group():
    """So "all of FAK is G0001, except block 2" says what it looks like."""
    overrides = {"FAK": "G0001", "FAK #2": "G0002"}

    assert resolve_for_block(overrides, "FAK", 1, 4, "fallback") == "G0001"
    assert resolve_for_block(overrides, "FAK", 2, 4, "fallback") == "G0002"


def test_with_nothing_set_the_fallback_stands():
    assert resolve_for_block({}, "FAK", 2, 4, "G0007") == "G0007"


# --- listing them for the editor ----------------------------------------------

def _tad_like() -> OpusRowSet:
    """One name, two blocks, differing by validity window - TAD's shape."""
    return OpusRowSet(
        rates=[_row(seq=1), _row(seq=1), _row(seq=2)],
        cmdt_notes=[
            CmdtNoteRow(header_seq=1, note_seq=1, contents="...", charge_seq=1, code="APP",
                        application_effective=date(2026, 9, 1), application_expires=date(2026, 9, 6)),
            CmdtNoteRow(charge_seq=2, code="LSF"),
            CmdtNoteRow(header_seq=2, note_seq=1, contents="...", charge_seq=1, code="APP",
                        application_effective=date(2026, 9, 7), application_expires=date(2026, 9, 15)),
            CmdtNoteRow(charge_seq=2, code="MBS"),
        ],
    )


def test_one_name_over_several_blocks_becomes_several_rows():
    """The whole point: four blocks under one name were one row in the
    settings, so they shared one code, one description and one skip."""
    blocks = commodity_blocks({"AEW": _tad_like()})

    assert [b.key for b in blocks] == ["FAK #1", "FAK #2"]
    assert [b.rows for b in blocks] == [2, 1]


def test_each_block_says_what_tells_it_apart():
    """Four rows called FAK are unreadable without this."""
    blocks = commodity_blocks({"AEW": _tad_like()})

    assert blocks[0].label == "01 Sep – 06 Sep · LSF"
    assert blocks[1].label == "07 Sep – 15 Sep · MBS"


def test_a_group_that_is_one_block_has_nothing_to_tell_apart():
    """An empty label, so the column stays quiet on the 13 lanes where
    every commodity group is a single block."""
    blocks = commodity_blocks({"": OpusRowSet(rates=[_row(description="OOG", seq=6)])})

    assert [b.key for b in blocks] == ["OOG"]
    assert blocks[0].label == ""


def test_two_sub_lanes_numbering_from_one_keep_their_own_blocks():
    """Deduping across scopes dropped every WMW block whose number WEW
    had already used - three of its four."""
    wew = OpusRowSet(rates=[_row(seq=1), _row(seq=2)])
    wmw = OpusRowSet(rates=[_row(seq=1), _row(seq=2)])

    blocks = commodity_blocks({"WEW": wew, "WMW": wmw})

    assert [(b.scope, b.key) for b in blocks] == [
        ("WEW", "FAK #1"), ("WEW", "FAK #2"), ("WMW", "FAK #1"), ("WMW", "FAK #2"),
    ]


# --- acting on one block ------------------------------------------------------

def test_skipping_one_block_leaves_the_others_and_takes_its_notes():
    row_set = OpusRowSet(
        rates=[_row(seq=1), _row(seq=2)],
        cmdt_notes=[
            CmdtNoteRow(header_seq=1, charge_seq=1, code="APP", group_description="FAK"),
            CmdtNoteRow(header_seq=2, charge_seq=1, code="APP", group_description="FAK"),
        ],
    )

    out = drop_commodity_groups(row_set, {"FAK #2"})

    assert [r.cmdt_seq for r in out.rates] == [1]
    assert [n.header_seq for n in out.cmdt_notes] == [1]


def test_skipping_the_whole_group_still_takes_every_block():
    """A setting made for the group reaches all of it - the behaviour
    every preset written before blocks existed relies on."""
    row_set = OpusRowSet(rates=[_row(seq=1), _row(seq=2)])

    assert drop_commodity_groups(row_set, {"FAK"}).rates == []
