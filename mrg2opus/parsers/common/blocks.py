"""Naming a commodity BLOCK, where a description isn't enough.

Every commodity setting is keyed by a group's default description, which
works because a description and a CMDT NOTE block are usually the same
thing - LAWC, LAEC, CSE, AUWC and the rest file one block per group.

TAD doesn't. Its blocks come from the data rather than from a sheet: rows
are grouped by their own validity window and Include Surcharge list, so
one commodity name covers several blocks with different sequence numbers.
A September AEW filing has four of them - two snapshots (1-6 Sep, 7-15
Sep) crossed with two surcharge lists (with LSF, without) - and all four
are called "FAK".

Keyed by description alone those four are one row in the settings, one
code, one skip. So where a description covers more than one block, the
block takes a key of its own: "FAK #2" for the second. Lookups try that
key first and the plain description second, which is what lets a setting
made for the whole group still reach every block of it - and what keeps
a preset written before blocks were separable working unchanged.
"""
from __future__ import annotations

from typing import Any, TypeVar

_T = TypeVar("_T")

_SEPARATOR = " #"


def block_key(default_description: str, cmdt_seq: Any, blocks_in_group: int = 1) -> str:
    """This block's settings key.

    Plain description when the group is one block, which is every lane
    but TAD - so nothing about their settings, or the presets holding
    them, changes.
    """
    if blocks_in_group <= 1 or cmdt_seq in (None, ""):
        return default_description
    return f"{default_description}{_SEPARATOR}{cmdt_seq}"


def describes_block(key: str, default_description: str, cmdt_seq: Any) -> bool:
    """Whether a settings key addresses this block - either by naming it
    outright or by naming the group it belongs to."""
    return key == default_description or key == f"{default_description}{_SEPARATOR}{cmdt_seq}"


def resolve_for_block(
    overrides: dict[str, _T],
    default_description: str,
    cmdt_seq: Any,
    blocks_in_group: int,
    fallback: _T,
) -> _T:
    """An override for this block, for its group, or the fallback.

    The block wins over the group, so "all of FAK is G0001, except block
    2 which is G0002" says what it looks like.
    """
    key = block_key(default_description, cmdt_seq, blocks_in_group)
    if key in overrides:
        return overrides[key]
    return overrides.get(default_description, fallback)
