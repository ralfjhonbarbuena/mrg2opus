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


class VerticalRatesRow(BaseModel):
    """OPUS's alternate "long format" rate upload: one row per container
    size instead of RatesRow's 4-rate-slots-per-row "horizontal" layout.
    Confirmed against reference/2_OPUS/25_TAD FILING OEW OMW's real
    "VERTICAL RATES" sheet - a pure derivative of RatesRow with no
    information of its own (see build_vertical_rates() below), NOT a
    TAD-specific concept: per the user, it's a general OPUS upload option
    (limited to 10,000 rows per file) available for any lane, faster to
    upload than the default "horizontal" RATES format. `per` combines the
    row's own Prefix with a size digit (2=20', 4=40', 5=40'HC, 7=45' -
    confirmed via the "D7 (OFT 45)" naming already seen in the TAD VBA
    tool's own SETTINGS toggle); `cargo_type` repeats the row's CGO TYPE
    alone (no prefix letter) as its own column, matching the real sheet's
    "PER" and "Cargo Type" columns exactly. No commodity_note/route_note
    columns exist on this sheet - confirmed, not an oversight.
    Header-group fields (CMDT Seq/Commodity Group/Actual Customer/Route
    Seq/Origin/O.Via/D.Via/Destination) are only populated on the FIRST
    exploded row for a given source RatesRow, blank on the rest - the
    same "shared header, blank-filled children" convention already used
    by CmdtNoteRow, baked into the data itself rather than the writer.
    """

    model_config = ConfigDict(str_strip_whitespace=True)

    cmdt_seq: Optional[int] = None
    commodity_group_code: Optional[str] = None
    commodity_group_description: Optional[str] = None
    actual_customer_code: Optional[str] = None
    actual_customer_description: Optional[str] = None
    route_seq: Optional[int] = None
    origin_code: Optional[str] = None
    origin_description: Optional[str] = None
    origin_term: Optional[str] = None
    origin_transmode: Optional[str] = None
    o_via_code: Optional[str] = None
    d_via_code: Optional[str] = None
    destination_code: Optional[str] = None
    destination_description: Optional[str] = None
    destination_term: Optional[str] = None
    destination_transmode: Optional[str] = None
    per: Optional[str] = None
    cargo_type: Optional[str] = None
    rate: Optional[Decimal] = None


# rate-field-name -> size digit, in the fixed column order the real sheet
# uses (20', 40', 40'HC, 45').
_VERTICAL_SIZE_SLOTS = [("rate_20", "2"), ("rate_40", "4"), ("rate_40hc", "5"), ("rate_45", "7")]

_VERTICAL_SHARED_FIELDS = [
    "cmdt_seq", "commodity_group_code", "commodity_group_description",
    "actual_customer_code", "actual_customer_description", "route_seq",
    "origin_code", "origin_description", "origin_term", "origin_transmode",
    "o_via_code", "d_via_code",
    "destination_code", "destination_description", "destination_term", "destination_transmode",
]


def explode_to_vertical_rates(row: RatesRow) -> list[VerticalRatesRow]:
    """One RatesRow -> 0-4 VerticalRatesRow, one per populated rate slot."""
    present = [(getattr(row, field), digit) for field, digit in _VERTICAL_SIZE_SLOTS if getattr(row, field) is not None]
    shared = {name: getattr(row, name) for name in _VERTICAL_SHARED_FIELDS}
    out: list[VerticalRatesRow] = []
    for i, (rate, digit) in enumerate(present):
        out.append(
            VerticalRatesRow(
                **(shared if i == 0 else {}),
                per=f"{row.prefix}{digit}",
                cargo_type=row.cgo_type,
                rate=rate,
            )
        )
    return out


def group_vertical_rates_by_cmdt_seq(
    rows: list[VerticalRatesRow],
) -> list[tuple[Optional[int], list[VerticalRatesRow]]]:
    """Split exploded vertical rates into one group per CMDT Seq, keeping
    the order they were built in.

    OPUS caps how many rows one upload may carry, and exploding RATES into
    a row per container size roughly triples the count - so a big lane
    overruns it as a single sheet but fits comfortably once split by
    commodity group (see excel_io/writer.py, which writes one numbered
    sheet per group).

    Grouping reads cmdt_seq off the rows themselves rather than taking it
    as an argument, which works because explode_to_vertical_rates() writes
    the header fields onto the FIRST row of each source RatesRow and
    blanks them on that row's continuation rows - so the last non-None
    value carries forward. A lane whose rows carry no cmdt_seq at all
    degrades to a single group.

    Buckets by cmdt_seq rather than by adjacency, because a commodity
    group's rows are NOT always contiguous: LAWC files its NOR rows under
    the SEA group, so group 2 appears as two separate runs. Grouping on
    adjacency emitted two sheets both named for group 2, and openpyxl
    silently renamed the second - splitting one commodity group across two
    misleadingly-named sheets. One group is always one sheet.
    """
    buckets: dict[Optional[int], list[VerticalRatesRow]] = {}
    order: list[Optional[int]] = []
    current: Optional[int] = None
    for row in rows:
        if row.cmdt_seq is not None:
            current = row.cmdt_seq
        if current not in buckets:
            buckets[current] = []
            order.append(current)
        buckets[current].append(row)
    return [(seq, buckets[seq]) for seq in order]


def build_vertical_rates(row_set: "OpusRowSet") -> "OpusRowSet":
    """User-toggled (MappingProfile.include_vertical_rates), applied
    uniformly across every lane by ui/parsing.py::run_parser() - not
    something individual parsers populate themselves. Derives entirely
    from row_set.rates (not rates_port_port - no evidence any lane needs
    a port-port equivalent of this sheet)."""
    vertical: list[VerticalRatesRow] = []
    for row in row_set.rates:
        vertical.extend(explode_to_vertical_rates(row))
    return row_set.model_copy(update={"vertical_rates": vertical})


class FreetimeRow(BaseModel):
    """A lane's static free-time-allowance/RFA reference table (real sheet
    name "FREETIME", sometimes filed as "FREETIME(NEEDS IMPROVEMENT)" - a
    filer's own WIP marker, not a naming variant to match on). Confirmed
    (LAWC: 3 ground-truth files; LAEC: 6) NOT derived from the raw MRG's
    rate data at all - every column is a static per-lane constant EXCEPT
    eff_dt/exp_dt, which for LAEC tracks that filing's own validity window
    (LAWC's stays a fixed constant even there - see
    parsers/common/freetime.py for the concrete per-lane tables and which
    behavior each lane uses)."""
    model_config = ConfigDict(str_strip_whitespace=True)

    seq: Optional[int] = None
    rfa_no: Optional[str] = None
    status: Optional[str] = None
    tariff: Optional[str] = None
    eff_dt: Optional[date] = None
    exp_dt: Optional[date] = None
    cntr_cargo: Optional[str] = None
    imdg_class: Optional[str] = None
    psa_grp: Optional[str] = None
    coverage_cn: Optional[str] = None
    coverage_rgn: Optional[str] = None
    coverage_loc: Optional[str] = None
    free_time_tier: Optional[str] = None
    free_time_add: Optional[str] = None
    free_time_total: Optional[Decimal] = None
    ftime_excl_sat: Optional[str] = None
    ftime_excl_sun: Optional[str] = None
    ftime_excl_hday: Optional[str] = None
    origin_or_dest_ct: Optional[str] = None
    origin_or_dest_cn: Optional[str] = None
    origin_or_dest_rgn: Optional[str] = None
    origin_or_dest_loc: Optional[str] = None
    bkg_del_cn: Optional[str] = None
    bkg_del_rgn: Optional[str] = None
    bkg_del_loc: Optional[str] = None
    actual_customer_code: Optional[str] = None
    actual_customer_name: Optional[str] = None
    commodity_code: Optional[str] = None
    commodity_name: Optional[str] = None
    curr: Optional[str] = None
    over_day_from: Optional[str] = None
    over_day_upto: Optional[str] = None
    rate_per_day_20: Optional[Decimal] = None
    rate_per_day_40: Optional[Decimal] = None
    rate_per_day_hc: Optional[Decimal] = None
    rate_per_day_45: Optional[Decimal] = None
    cntr_qty_from: Optional[str] = None
    cntr_qty_upto: Optional[str] = None
    tiered_free_time: Optional[str] = None
    remark: Optional[str] = None
    dar_no: Optional[str] = None
    ver: Optional[str] = None
    approval_no: Optional[str] = None
    proposal_no: Optional[str] = None
    customer_code: Optional[str] = None
    customer_name: Optional[str] = None


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
    freetime: list[FreetimeRow] = []
    # Populated only when MappingProfile.include_vertical_rates is set -
    # see build_vertical_rates() above. Empty by default for every lane.
    vertical_rates: list[VerticalRatesRow] = []
