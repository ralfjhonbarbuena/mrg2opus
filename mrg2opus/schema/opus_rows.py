"""Pydantic models for each OPUS output row type.

These are the contract every lane parser must fill (via BaseMRGParser.to_opus_rows)
and every exporter (excel_io.writer) reads from. Field order matches
schema.opus_columns.*_ROW_FIELDS so the writer can zip them directly.
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal
from typing import Optional

from pydantic import BaseModel, ConfigDict


class RatesRow(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    # Always "C" - a deliberate, user-directed business rule applied uniformly
    # across every lane regardless of what any single ground-truth sample
    # showed (some samples had it blank; that's now treated as sample noise,
    # not a per-lane convention to reproduce).
    type: Optional[str] = "C"
    cmdt_seq: Optional[int] = None
    commodity_group_code: str
    commodity_group_description: str
    actual_customer_code: Optional[str] = None
    actual_customer_description: Optional[str] = None
    route_seq: Optional[int] = None
    origin_code: str
    origin_description: str
    origin_term: Optional[str] = None
    origin_transmode: Optional[str] = None
    o_via_code: Optional[str] = None
    d_via_code: Optional[str] = None
    destination_code: str
    destination_description: str
    destination_term: Optional[str] = None
    destination_transmode: Optional[str] = None
    prefix: str
    cgo_type: str
    cur_20: Optional[str] = None
    rate_20: Optional[Decimal] = None
    cur_40: Optional[str] = None
    rate_40: Optional[Decimal] = None
    cur_40hc: Optional[str] = None
    rate_40hc: Optional[Decimal] = None
    cur_45: Optional[str] = None
    rate_45: Optional[Decimal] = None
    commodity_note: Optional[str] = None
    route_note: Optional[str] = None


class RatesPortPortRow(RatesRow):
    """Same shape as RatesRow, but one row per single port code.

    Built via RatesRow.explode_to_port_port(), which splits origin_code /
    destination_code on ';' into individual rows while copying the combined
    origin_description / destination_description verbatim to every exploded
    row. This reproduces the (slightly odd) behavior observed in the ground
    truth samples: OPUS RATES PORT-PORT does NOT re-split the description
    per port, it just repeats the whole grouped description string.
    """


def _explode_group(row: RatesRow) -> list[RatesPortPortRow]:
    origin_codes = [c for c in row.origin_code.split(";") if c]
    dest_codes = [c for c in row.destination_code.split(";") if c]
    if not origin_codes:
        origin_codes = [row.origin_code]
    if not dest_codes:
        dest_codes = [row.destination_code]

    data = row.model_dump()
    # cmdt_seq and commodity_note are properties of the grouped view, not
    # the exploded one - confirmed against CSE's ground truth (OPUS RATES
    # has cmdt_seq=1/2/3, OPUS RATES PORT-PORT leaves it blank on every
    # row) and LAEC's (OPUS RATES copies the group's CMDT NOTE Contents
    # text into every row's commodity_note, OPUS RATES PORT-PORT doesn't).
    data["cmdt_seq"] = None
    data["commodity_note"] = None
    out: list[RatesPortPortRow] = []
    for o_code in origin_codes:
        for d_code in dest_codes:
            row_data = dict(data)
            row_data["origin_code"] = o_code
            row_data["destination_code"] = d_code
            out.append(RatesPortPortRow(**row_data))
    return out


# attach as a free function usable from parsers: explode_rates_row(row)
def explode_rates_row(row: RatesRow) -> list[RatesPortPortRow]:
    return _explode_group(row)


class ArbsRow(BaseModel):
    """Field names/order verified directly against CSE.xlsx's OPUS ARBS
    ground truth (the Phase-1 placeholder schema was wrong - e.g. Proposal/
    C.Offer/Final are three separate columns, not one combined string, and
    Weight is two columns not one)."""
    model_config = ConfigDict(str_strip_whitespace=True)

    seq: Optional[int] = None
    point: Optional[str] = None
    description: Optional[str] = None
    trans_mode: Optional[str] = None
    term: Optional[str] = None
    service_lane: Optional[str] = None
    trunk_lane: Optional[str] = None
    weight_gte_mt: Optional[str] = None
    weight_lt_mt: Optional[str] = None
    over: Optional[str] = None
    via: Optional[str] = None
    actual_customer: Optional[str] = None
    pay_term: Optional[str] = None
    per: Optional[str] = None
    cgo_type: Optional[str] = None
    cur: Optional[str] = None
    proposal: Optional[Decimal] = None
    c_offer: Optional[Decimal] = None
    final: Optional[Decimal] = None
    eff_date: Optional[date] = None
    exp_date: Optional[date] = None
    source: Optional[str] = None
    status: Optional[str] = None
    seq2: Optional[str] = None  # ground truth header repeats a blank "seq" column - always empty
    note: Optional[str] = None
    remark: Optional[str] = None


class CmdtNoteRow(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    header_seq: Optional[int] = None
    note_seq: Optional[int] = None
    contents: Optional[str] = None
    charge_seq: Optional[int] = None
    code: Optional[str] = None
    application_effective: Optional[date] = None
    application_expires: Optional[date] = None
    application: Optional[str] = None
    cur: Optional[str] = None
    cal: Optional[str] = None
    amount: Optional[Decimal] = None
    pay_term: Optional[str] = None
    pay_ofc: Optional[str] = None
    payer: Optional[str] = None
    per: Optional[str] = None
    cgo_type: Optional[str] = None
    imdg_class: Optional[str] = None
    psa_grp: Optional[str] = None
    food_grade: Optional[str] = None
    lane: Optional[str] = None
    ts_port: Optional[str] = None
    canal: Optional[str] = None
    vvd: Optional[str] = None
    soc: Optional[str] = None
    por: Optional[str] = None
    pol: Optional[str] = None
    pod: Optional[str] = None
    delivery: Optional[str] = None
    node: Optional[str] = None
    cmdt: Optional[str] = None

    # Internal bookkeeping only - NOT one of the OPUS output columns (see
    # schema/opus_columns.py::CMDT_NOTE_ROW_FIELDS, which the writer zips
    # against explicitly and does not include this field, so it never
    # reaches the written file). Tags every row in a block (parent AND
    # children) with the commodity group description it was built for, so
    # parsers/common/ordering.py::reorder_row_set() can reorder whole CMDT
    # NOTE blocks by the user's chosen commodity_group_order without
    # needing to re-derive block boundaries from the flattened row list.
    group_description: Optional[str] = None


class SpecialNoteRow(CmdtNoteRow):
    """Same shape as CmdtNoteRow; written to a separate OPUS SPECIAL NOTE sheet."""


class RouteNoteRow(CmdtNoteRow):
    """Scoped to a specific route pair rather than a whole commodity
    sequence (see project-opus-note-sheet-taxonomy memory) - written to the
    real OPUS system's "RN" sheet. Same shape as CmdtNoteRow plus route_seq
    (links back to the RatesRow it was derived from - see RatesRow.route_seq)
    and the 10 trailing columns confirmed on the real RN sheet
    (schema/opus_columns.py::RN_HEADER) that CmdtNoteRow doesn't carry.
    Unlike CMDT NOTE, real RN rows are header-only (charge_seq/code are
    always 1/"APP", no child charge-code rows) - there's no per-route
    surcharge-code breakdown to attach."""

    route_seq: Optional[int] = None
    receiving_term: Optional[str] = None
    delivery_term: Optional[str] = None
    weight_gte_mt: Optional[str] = None
    weight_lt_mt: Optional[str] = None
    direct_call: Optional[str] = None
    bar_type: Optional[str] = None
    s_i: Optional[str] = None
    mty_pickup_cy: Optional[str] = None
    mty_return_cy: Optional[str] = None
    premium: Optional[str] = None


class OpusRowSet(BaseModel):
    """Everything one parser run produces, ready for excel_io.writer."""

    rates: list[RatesRow] = []
    rates_port_port: list[RatesPortPortRow] = []
    arbs: list[ArbsRow] = []
    cmdt_notes: list[CmdtNoteRow] = []
    special_notes: list[SpecialNoteRow] = []
    route_notes: list[RouteNoteRow] = []
