"""Row-ordering helpers shared across lane parsers, so generated OPUS
sheets read as organized blocks (all same-prefix rows together, grouped by
destination within each block) rather than interleaved by source row.
"""
from __future__ import annotations

from typing import Callable, TypeVar

from mrg2opus.schema.opus_rows import OpusRowSet, RatesRow


def group_by_destination(rows: list[RatesRow]) -> list[RatesRow]:
    """Stable-group rows by destination_code, preserving each destination's
    first-seen order (not alphabetical) and each row's relative order
    within its destination group."""
    order: list[str] = []
    buckets: dict[str, list[RatesRow]] = {}
    for row in rows:
        key = row.destination_code
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(row)
    return [row for key in order for row in buckets[key]]


_T = TypeVar("_T")


def reorder_by_group(items: list[_T], group_order: list[str], key_fn: Callable[[_T], str]) -> list[_T]:
    """Stable-reorders items into blocks matching group_order (a list of
    group-identity values in the user's desired sequence). Groups not
    mentioned in group_order keep their first-seen relative order and are
    appended after every explicitly-ordered group. Never reorders WITHIN a
    group - only changes which group's block comes first."""
    if not group_order:
        return items
    buckets: dict[str, list[_T]] = {}
    first_seen: list[str] = []
    for item in items:
        key = key_fn(item)
        if key not in buckets:
            buckets[key] = []
            first_seen.append(key)
        buckets[key].append(item)
    ordered_keys = [k for k in group_order if k in buckets]
    remaining_keys = [k for k in first_seen if k not in group_order]
    return [item for key in (*ordered_keys, *remaining_keys) for item in buckets[key]]


def reorder_row_set(row_set: OpusRowSet, group_order: list[str]) -> OpusRowSet:
    """Reorders row_set.rates/rates_port_port/cmdt_notes to follow
    group_order (MappingProfile.commodity_group_order - a list of FINAL,
    post-override commodity_group_description values in the user's chosen
    sequence). rates/rates_port_port are keyed by each row's own
    commodity_group_description; cmdt_notes by the internal
    group_description bookkeeping field (see CmdtNoteRow), since a CMDT
    NOTE block's own rows don't carry commodity_group_description at all.

    Known scope limit: on lanes where the OPUS RATES PORT-PORT sheet uses
    a DIFFERENT commodity_group_description than the RATES sheet for the
    same group (LAWC's PP_COMMODITY remap, e.g. "China_TWN_SIN_HKG_KR Dry"
    on RATES vs "FAK - China_TWN_SIN_HKG Dry" on PORT-PORT), group_order
    (built from RATES-level descriptions) won't match PORT-PORT's rows,
    so PORT-PORT keeps its default order for that lane. RATES and CMDT
    NOTE - the two sheets a user actually perceives "order" from - are
    unaffected by this."""
    if not group_order:
        return row_set
    return row_set.model_copy(
        update={
            "rates": reorder_by_group(row_set.rates, group_order, lambda r: r.commodity_group_description),
            "rates_port_port": reorder_by_group(
                row_set.rates_port_port, group_order, lambda r: r.commodity_group_description
            ),
            "cmdt_notes": reorder_by_group(row_set.cmdt_notes, group_order, lambda r: r.group_description or ""),
        }
    )
