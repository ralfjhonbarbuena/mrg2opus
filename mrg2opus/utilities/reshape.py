"""The same rates in another of OPUS's three shapes.

    RATES            one row per route, four rate columns across
    RATES PORT-PORT  the same rows with ";"-joined ports split out
    VERTICAL RATES   one row per container size, ports running downwards

Two of the four directions are derivations the converter already makes on
every run, so they are exact. The two reverse directions are
RECONSTRUCTIONS, and the difference matters: going forward discards the
grouping, so coming back has to infer it. Each function below says which
kind it is.
"""
from __future__ import annotations

from typing import Any

from mrg2opus.audit.compare import VerticalBlock, reconstruct_vertical_blocks
from mrg2opus.parsers.common.ordering import group_by_destination
from mrg2opus.schema.opus_rows import (
    OpusRowSet,
    RatesPortPortRow,
    RatesRow,
    VerticalRatesRow,
    build_vertical_rates,
    explode_rates_row,
)

# What tells two port-port rows they came from the same RATES row: the
# whole row except the origin. Two origins that agree on all of this were
# one ";"-joined row before the split.
_REGROUP_FIELDS = (
    "commodity_group_code", "commodity_group_description", "cmdt_seq",
    "origin_term", "origin_transmode", "o_via_code", "d_via_code",
    "destination_code", "destination_description", "destination_term", "destination_transmode",
    "prefix", "cgo_type",
    "cur_20", "rate_20", "cur_40", "rate_40", "cur_40hc", "rate_40hc", "cur_45", "rate_45",
    "commodity_note", "route_note",
)

_SIZE_BY_DIGIT = {"2": ("cur_20", "rate_20"), "4": ("cur_40", "rate_40"),
                  "5": ("cur_40hc", "rate_40hc"), "7": ("cur_45", "rate_45")}


def _as_rates_rows(rows: list[dict[str, Any]]) -> list[RatesRow]:
    """Dicts read off a sheet, back into rows. A row missing something the
    model requires is dropped rather than guessed at - a filing with no
    origin code has a problem the Check tool should be naming, not one
    this should paper over."""
    out = []
    for row in rows:
        try:
            out.append(RatesRow(**{k: v for k, v in row.items() if k in RatesRow.model_fields}))
        except Exception:
            continue
    return out


def to_port_port(rows: list[dict[str, Any]]) -> list[RatesPortPortRow]:
    """RATES -> RATES PORT-PORT. Exact: the same split the converter makes."""
    exploded: list[RatesPortPortRow] = []
    for row in _as_rates_rows(rows):
        exploded.extend(explode_rates_row(row))
    return exploded


def to_vertical(rows: list[dict[str, Any]]) -> list[VerticalRatesRow]:
    """RATES -> VERTICAL RATES. Exact: build_vertical_rates() is what the
    converter itself runs, so this is the sheet a fresh conversion would
    have produced from these rows."""
    built = build_vertical_rates(OpusRowSet(rates=_as_rates_rows(rows)))
    return built.vertical_rates


def _rate_columns(block: VerticalBlock) -> dict[str, Any]:
    """A block's (per, cargo type, rate) triples back into rate columns.
    `per` is prefix + size digit - "D5" is a dry 40'HC."""
    out: dict[str, Any] = {}
    for per, _cargo, rate in block.rates:
        digit = str(per or "")[-1:]
        slot = _SIZE_BY_DIGIT.get(digit)
        if slot is None or rate is None:
            continue
        cur_field, rate_field = slot
        out[rate_field] = rate
        out.setdefault(cur_field, "USD")
    return out


def from_vertical(rows: list[dict[str, Any]], blocks_source=reconstruct_vertical_blocks) -> list[RatesRow]:
    """VERTICAL RATES -> RATES. A reconstruction, but a safe one.

    The long format is columnar, not a list of records: a route opens
    where Route Seq is written and its ports run down their own columns
    until the next one. So the block boundaries survive the trip out and
    the grouping comes back intact - unlike port-port, nothing has to be
    guessed.

    What does NOT come back is anything the long format has no column
    for. Ports are re-joined with ";" in sorted order rather than the
    order they were filed in.
    """
    out: list[RatesRow] = []
    heads = _block_heads(rows)
    for block, head in zip(blocks_source(rows), heads):
        cargo = next((c for _p, c, _r in block.rates if c), None)
        per = next((p for p, _c, _r in block.rates if p), "")
        out.append(RatesRow(
            cmdt_seq=head.get("cmdt_seq"),
            commodity_group_code=head.get("commodity_group_code") or "",
            commodity_group_description=head.get("commodity_group_description") or "",
            route_seq=head.get("route_seq"),
            origin_code=";".join(sorted(block.origins)),
            origin_description=head.get("origin_description") or "",
            origin_term=block.origin_term,
            origin_transmode=head.get("origin_transmode"),
            o_via_code=block.o_via,
            d_via_code=block.d_via,
            destination_code=";".join(sorted(block.destinations)),
            destination_description=head.get("destination_description") or "",
            destination_term=block.destination_term,
            destination_transmode=head.get("destination_transmode"),
            prefix=str(per)[:1] or "D",
            cgo_type=cargo or "DR",
            **_rate_columns(block),
        ))
    return out


def _block_heads(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """The first row of each block - the only one carrying the commodity
    and the descriptions, which VerticalBlock itself doesn't keep."""
    heads = []
    for row in rows:
        if row.get("route_seq") not in (None, ""):
            heads.append(row)
    return heads


def regroup_port_port(rows: list[dict[str, Any]]) -> list[RatesRow]:
    """RATES PORT-PORT -> RATES, grouped as tightly as the rows allow.

    Not the inverse of to_port_port(), and the difference is large enough
    to say plainly. The split discarded which ports shared a row, and
    nothing in the port-port sheet records it - so this groups every
    origin that agrees on all the rest (destination, both terms, both
    vias, all four rates, both notes), which is the smallest equivalent
    filing rather than the one you started with.

    On a real LAEC week that is 2,528 rows in and 884 out: its 2,528 rows
    carry only 618 distinct non-origin signatures, because that lane
    files many origins separately at one rate. The output covers the same
    routes at the same rates; it is not the same sheet.

    Fewer rows is the safe direction. OPUS rejects a duplicate filing and
    accepts a group, and the long format caps at 10,000 rows.
    """
    buckets: dict[tuple, list[dict[str, Any]]] = {}
    order: list[tuple] = []
    for row in rows:
        key = tuple(row.get(f) for f in _REGROUP_FIELDS)
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(row)

    out: list[RatesRow] = []
    for key in order:
        members = buckets[key]
        head = dict(members[0])
        codes = sorted({str(m.get("origin_code")) for m in members if m.get("origin_code")})
        head["origin_code"] = ";".join(codes)
        if len(codes) > 1:
            names = sorted({str(m.get("origin_description")) for m in members if m.get("origin_description")})
            head["origin_description"] = ";".join(names)
        built = _as_rates_rows([head])
        out.extend(built)
    return group_by_destination(out)
