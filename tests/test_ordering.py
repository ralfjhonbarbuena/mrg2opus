from __future__ import annotations

from mrg2opus.parsers.common.ordering import group_by_destination, reorder_by_group, reorder_row_set
from mrg2opus.schema.opus_rows import CmdtNoteRow, OpusRowSet, RatesRow


def _rate(desc: str, dest: str) -> RatesRow:
    return RatesRow(
        commodity_group_code="G0001",
        commodity_group_description=desc,
        origin_code="AA",
        origin_description="Origin",
        destination_code=dest,
        destination_description="Dest",
        prefix="D",
        cgo_type="DR",
    )


def test_reorder_by_group_moves_whole_blocks():
    items = ["a1", "b1", "a2", "c1", "b2"]
    ordered = reorder_by_group(items, ["c", "a", "b"], key_fn=lambda s: s[0])
    assert ordered == ["c1", "a1", "a2", "b1", "b2"]


def test_reorder_by_group_unlisted_groups_appended_in_first_seen_order():
    items = ["a1", "b1", "c1"]
    ordered = reorder_by_group(items, ["b"], key_fn=lambda s: s[0])
    assert ordered == ["b1", "a1", "c1"]


def test_reorder_by_group_empty_order_is_noop():
    items = ["b1", "a1"]
    assert reorder_by_group(items, [], key_fn=lambda s: s[0]) == items


def test_group_by_destination_unaffected_by_new_import():
    rows = [_rate("G", "Y"), _rate("G", "X"), _rate("G", "Y")]
    result = group_by_destination(rows)
    assert [r.destination_code for r in result] == ["Y", "Y", "X"]


def test_reorder_row_set_reorders_rates_and_cmdt_notes_together():
    rates = [_rate("SEA", "P1"), _rate("MAIN", "P2"), _rate("ISC", "P3")]
    notes = [
        CmdtNoteRow(code="APP", group_description="SEA"),
        CmdtNoteRow(code="EFS", group_description="SEA"),
        CmdtNoteRow(code="APP", group_description="MAIN"),
        CmdtNoteRow(code="APP", group_description="ISC"),
    ]
    row_set = OpusRowSet(rates=rates, cmdt_notes=notes)

    reordered = reorder_row_set(row_set, ["ISC", "MAIN", "SEA"])

    assert [r.commodity_group_description for r in reordered.rates] == ["ISC", "MAIN", "SEA"]
    assert [n.group_description for n in reordered.cmdt_notes] == ["ISC", "MAIN", "SEA", "SEA"]


def test_reorder_row_set_no_order_is_noop():
    row_set = OpusRowSet(rates=[_rate("A", "X")])
    assert reorder_row_set(row_set, []) is row_set
