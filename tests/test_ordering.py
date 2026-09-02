from __future__ import annotations

from mrg2opus.parsers.common.ordering import (
    drop_commodity_groups,
    group_by_destination,
    reorder_by_group,
    reorder_row_set,
)
from mrg2opus.schema.opus_rows import CmdtNoteRow, OpusRowSet, RatesRow, explode_rates_row


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


def _rates_row(group: str, code: str = "G0001", origin: str = "CNSHA;CNNGB") -> RatesRow:
    return RatesRow(
        commodity_group_code=code,
        commodity_group_description=group,
        origin_code=origin,
        origin_description=origin,
        destination_code="AEJEA",
        destination_description="AEJEA",
        prefix="D",
        cgo_type="DR",
    )


def _row_set_with(*groups: str) -> OpusRowSet:
    rates, port_port, notes = [], [], []
    for i, g in enumerate(groups, start=1):
        row = _rates_row(g, code=f"G{i:04d}")
        row.cmdt_seq = i
        rates.append(row)
        port_port.extend(explode_rates_row(row))
        notes.append(CmdtNoteRow(contents=f"{g} note", charge_seq=1, group_description=g))
        notes.append(CmdtNoteRow(contents=None, charge_seq=2, group_description=g))
    return OpusRowSet(rates=rates, rates_port_port=port_port, cmdt_notes=notes)


def test_skipping_a_group_drops_its_rates_port_port_rows_and_note_block():
    row_set = drop_commodity_groups(_row_set_with("Dry", "Reefer", "OOG"), {"Reefer"})

    assert [r.commodity_group_description for r in row_set.rates] == ["Dry", "OOG"]
    assert {r.source_group for r in row_set.rates_port_port} == {"Dry", "OOG"}
    assert {n.group_description for n in row_set.cmdt_notes} == {"Dry", "OOG"}
    # the whole block goes, parent row and blank-contents children alike
    assert len(row_set.cmdt_notes) == 4


def test_port_port_is_matched_on_its_source_group_not_a_remapped_one():
    """LAWC rewrites commodity_group_description on PORT-PORT rows, so the
    exploded row's own description can't be the thing matched."""
    row_set = _row_set_with("Dry", "Reefer")
    for r in row_set.rates_port_port:
        r.commodity_group_description = "FAK - COMBINED"

    kept = drop_commodity_groups(row_set, {"Reefer"})

    assert {r.source_group for r in kept.rates_port_port} == {"Dry"}


def test_skipping_nothing_returns_the_row_set_untouched():
    original = _row_set_with("Dry", "Reefer")
    assert drop_commodity_groups(original, set()) is original


def test_remaining_groups_keep_their_own_cmdt_seq_numbers():
    """Deliberately NOT compacted - see MappingProfile.skip_commodity_filing."""
    kept = drop_commodity_groups(_row_set_with("Dry", "Reefer", "OOG"), {"Reefer"})
    assert [r.cmdt_seq for r in kept.rates] == [1, 3]
