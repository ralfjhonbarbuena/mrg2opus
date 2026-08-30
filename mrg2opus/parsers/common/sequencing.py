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

from typing import Hashable, TypeVar

from mrg2opus.schema.opus_rows import RatesRow

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
