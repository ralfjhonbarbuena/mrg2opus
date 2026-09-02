"""Shared CMDT Seq / Route Seq auto-numbering, reused by every lane parser.

Business rule (explicit user direction, 2026-08-30): every generated RATES
row must carry a real CMDT Seq and Route Seq - the tool should never leave
them blank, even where a real ground-truth filing does (confirmed e.g.
EAF-KEMBA's own OPUS output leaves both entirely blank - a real filing
convention, not a parser gap, but going forward the generated output
always fills them in).

CMDT Seq numbers distinct CMDT NOTE blocks, starting at 1, in the order a
lane builds them - NOT by commodity_group_code/description alone: two
blocks can share the identical group code/description and still need
different seq numbers (confirmed pattern: TAD's per-validity-window/
per-charge-code-combo blocks, all filed under the same "FAK"/G0001 group).
Each lane already knows its own block identity when it builds CMDT NOTE
rows - this module only standardizes the NUMBERING mechanics (auto-assign
around any explicit override, lowest free integer first) and the
universal, block-key-agnostic Route Seq reset (1..N within each cmdt_seq,
in existing row order).
"""
from __future__ import annotations

import re
from typing import Hashable, TypeVar

from mrg2opus.schema.opus_rows import CmdtNoteRow, OpusRowSet, RatesPortPortRow, RatesRow, RouteNoteRow

_K = TypeVar("_K", bound=Hashable)


def assign_cmdt_seq_numbers(block_keys_in_order: list[_K], overrides: dict[_K, int] | None = None) -> dict[_K, int]:
    """Given the distinct CMDT NOTE block keys a lane already computed (one
    entry per block, in first-seen/build order - duplicates in the input
    are harmless, only the first occurrence of each key matters), return
    {key: cmdt_seq}. A key present in `overrides` keeps that exact number;
    every other key is auto-numbered around the fixed ones, lowest
    available integer first, so an override never collides with an
    auto-assigned number."""
    overrides = overrides or {}
    seen: list[_K] = []
    for key in block_keys_in_order:
        if key not in seen:
            seen.append(key)

    assigned: dict[_K, int] = {k: v for k, v in overrides.items() if k in seen}
    used = set(assigned.values())
    next_seq = 1
    for key in seen:
        if key in assigned:
            continue
        while next_seq in used:
            next_seq += 1
        assigned[key] = next_seq
        used.add(next_seq)
        next_seq += 1
    return assigned


def assign_route_seq(rates: list[RatesRow]) -> None:
    """Number route_seq 1..N within each existing cmdt_seq value, in the
    rows' current order - call this once every row's cmdt_seq is final."""
    counters: dict[int | None, int] = {}
    for row in rates:
        counters[row.cmdt_seq] = counters.get(row.cmdt_seq, 0) + 1
        row.route_seq = counters[row.cmdt_seq]


def sync_port_port_cmdt_seq(rates: list[RatesRow], rates_port_port: list[RatesPortPortRow]) -> None:
    """Give every exploded OPUS RATES PORT-PORT row the same CMDT Seq as
    the RATES row it came from, so the two sheets line up on commodity
    seq + commodity code (user-reported, 2026-08-31).

    It can't be done at explode time: a lane builds its PORT-PORT rows
    inside the same loop that builds RATES, and only assigns cmdt_seq
    once every group is known - so the source row's number is still None
    when the copy is taken. Keyed on RatesPortPortRow.source_group (the
    ORIGINAL group description, recorded at explode time) rather than the
    exploded row's own commodity_group_description, because LAWC rewrites
    that on its PORT-PORT rows. Rows whose group isn't in `rates` at all
    are left as they are."""
    if not rates_port_port:
        return
    seq_by_group: dict[str, int | None] = {}
    for row in rates:
        seq_by_group.setdefault(row.commodity_group_description, row.cmdt_seq)
    for row in rates_port_port:
        key = row.source_group if row.source_group is not None else row.commodity_group_description
        if key in seq_by_group:
            row.cmdt_seq = seq_by_group[key]


def collapse_note_block_sequences(notes: list[CmdtNoteRow]) -> None:
    """Keep Header Seq on a CMDT NOTE / SPECIAL NOTE block's FIRST row only,
    and put Note Seq 1 beside it (user-reported, 2026-08-31).

    Both were wrong: parsers stamp header_seq onto every row of a block
    (`for note in notes: note.header_seq = seq`), so the same number
    repeated down the sheet, and nothing ever set note_seq, so it came out
    blank everywhere. Real filings put both on the parent row alone and
    leave every child row blank - confirmed against EAF-KEMBA's "CMDT
    NOTE", WAF's "SRCHG" and LAEC FAK's "SPECIAL NOTE" sheets.

    A row with non-blank `contents` starts a block; the blank-contents
    rows after it are its children. That's the same invariant
    audit/compare.py::reconstruct_blocks() already keys on, and the one
    the writer's fill-down produces. Deliberately NOT applied to ROUTE
    NOTE rows: LAWC's real "RN" sheet repeats one header_seq across every
    route row and carries note_seq 1 on all of them, so the same
    collapse there would be a regression.

    An existing note_seq on a parent is left alone - CSE and LAEC file
    theirs under lane-specific numbers confirmed against their own ground
    truth, not 1. If no row carries contents at all the list is left
    untouched, rather than blanking a sheet whose shape this doesn't
    describe."""
    if not any(str(n.contents or "").strip() for n in notes):
        return
    for note in notes:
        if str(note.contents or "").strip():
            if note.note_seq is None:
                note.note_seq = 1
        else:
            note.header_seq = None
            note.note_seq = None


# The service lane a ROUTE NOTE names, e.g. "...Vessel Service Lane: MX2"
# -> "MX2". \S+ stops at the first space, which is what drops LAWC's
# "KCI (OH)" / "KCI (OW)" gauge suffixes - ground truth files those under
# a bare "KCI".
_SERVICE_LANE_RE = re.compile(r"Vessel Service Lane:\s*(\S+)", re.IGNORECASE)


def service_lane_from_note(contents: str | None) -> str | None:
    """The 3-character vessel service lane code a route note names, or None
    when it names none."""
    if not contents:
        return None
    m = _SERVICE_LANE_RE.search(contents)
    return m.group(1) if m else None


def fill_route_note_lanes(route_notes: list[RouteNoteRow]) -> None:
    """Populate each ROUTE NOTE row's own `Lane` column from the service
    lane its Contents names (user-reported, 2026-08-31: the column came
    out empty on every lane).

    Every parser put the code in the note TEXT only, leaving the dedicated
    column - which is what OPUS actually reads the lane off - blank. The
    rule is the same in every ground truth checked: the column carries the
    code exactly when the note names one, and stays blank otherwise
    (LAWC's bare "OH"/"OW"/"REEFER DRY AS DANGEROUS" notes, and TAD's
    transhipment-port notes). Confirmed against LAWC's RN (MAR 196, MX2
    88, AX3 30, KCI 36) and both TAD ROUTE NOTE files (IOM/AEX/MD1, FE3),
    including the combined "T/S port | service lane" and "REEFER DRY AS
    DANGEROUS | ...Lane: AX3" texts, which take the lane from their
    service-lane half. An explicitly-set lane is never overwritten."""
    for note in route_notes:
        if note.lane is None:
            note.lane = service_lane_from_note(note.contents)


def finalize_sequences(row_set: OpusRowSet) -> None:
    """Every cross-sheet sequencing fixup that can only run once a lane has
    finished building, applied in place. Lives here rather than in each
    parser because all 17 lanes need identical behavior and each was
    getting it subtly differently."""
    sync_port_port_cmdt_seq(row_set.rates, row_set.rates_port_port)
    collapse_note_block_sequences(row_set.cmdt_notes)
    collapse_note_block_sequences(row_set.special_notes)
    fill_route_note_lanes(row_set.route_notes)
