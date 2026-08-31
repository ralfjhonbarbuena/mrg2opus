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

    # Internal bookkeeping, never written to a sheet - the writer picks
    # columns by name from RATES_PORT_PORT_ROW_FIELDS, which doesn't list
    # this. Holds the commodity_group_description of the RatesRow this row
    # was exploded from, BEFORE any lane-specific PORT-PORT remap (LAWC
    # rewrites the group description on its exploded rows), which is what
    # lets the pipeline stamp each row with its group's FINAL cmdt_seq -
    # a number the parsers only assign after exploding. Same idea as
    # CmdtNoteRow.group_description.
    source_group: Optional[str] = None


def _explode_group(row: RatesRow) -> list[RatesPortPortRow]:
    origin_codes = [c for c in row.origin_code.split(";") if c]
    dest_codes = [c for c in row.destination_code.split(";") if c]
    if not origin_codes:
        origin_codes = [row.origin_code]
    if not dest_codes:
        dest_codes = [row.destination_code]

    data = row.model_dump()
    # commodity_note is a property of the grouped view, not the exploded
    # one - confirmed against LAEC's ground truth (OPUS RATES copies the
    # group's CMDT NOTE Contents text into every row's commodity_note,
    # OPUS RATES PORT-PORT doesn't).
    #
    # cmdt_seq used to be blanked here too, on the strength of the old
    # bundled CSE.xlsx fixture; no real filing in reference/2_OPUS has a
    # PORT-PORT sheet at all, so there was never ground truth for it, and
    # per the user PORT-PORT must carry the same CMDT Seq (and commodity
    # code) as RATES. It's left alone here and stamped once the group's
    # final number exists - see sequencing.sync_port_port_cmdt_seq().
    data["commodity_note"] = None
    data["source_group"] = row.commodity_group_description
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
    Blank-fills at TWO levels, not one - the same "shared header,
    blank-filled children" idea as CmdtNoteRow, baked into the data rather
    than the writer, but applied at two different scopes:
      - Route Seq / Origin / O.Via / D.Via / Destination appear on the
        first exploded row of each source RatesRow, blank on that row's
        other container sizes.
      - CMDT Seq / Commodity Group / Actual Customer appear only where a
        new commodity group STARTS, blank for every other row in it -
        OPUS reads a blank commodity block as "same group as above".
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

# The sheet blank-fills at TWO different levels, confirmed against real
# ground truth (reference/2_OPUS/27's own WEW VERTICAL RATES: 5,209 data
# rows, CMDT Seq written on just 4 of them - once per commodity group -
# while Route Seq and the origin/destination block repeat for every rate
# row). OPUS reads a blank commodity block as "same group as above", so
# restating it per row would be wrong, not merely redundant.
_VERTICAL_GROUP_FIELDS = [
    "cmdt_seq", "commodity_group_code", "commodity_group_description",
    "actual_customer_code", "actual_customer_description",
]
_VERTICAL_ROW_FIELDS = [
    "route_seq",
    "origin_code", "origin_description", "origin_term", "origin_transmode",
    "o_via_code", "d_via_code",
    "destination_code", "destination_description", "destination_term", "destination_transmode",
]


def explode_to_vertical_rates(row: RatesRow, *, start_of_group: bool = True) -> list[VerticalRatesRow]:
    """One RatesRow -> 0-4 VerticalRatesRow, one per populated rate slot.

    start_of_group: whether this row opens a new CMDT Seq block. Only then
    does the commodity block get written - see _VERTICAL_GROUP_FIELDS.
    """
    present = [(getattr(row, field), digit) for field, digit in _VERTICAL_SIZE_SLOTS if getattr(row, field) is not None]
    header = {name: getattr(row, name) for name in _VERTICAL_ROW_FIELDS}
    if start_of_group:
        header.update({name: getattr(row, name) for name in _VERTICAL_GROUP_FIELDS})
    out: list[VerticalRatesRow] = []
    for i, (rate, digit) in enumerate(present):
        out.append(
            VerticalRatesRow(
                **(header if i == 0 else {}),
                per=f"{row.prefix}{digit}",
                cargo_type=row.cgo_type,
                rate=rate,
            )
        )
    return out


def build_vertical_rates(row_set: "OpusRowSet") -> "OpusRowSet":
    """User-toggled (MappingProfile.include_vertical_rates), applied
    uniformly across every lane by pipeline.py::run_parser() - not
    something individual parsers populate themselves. Derives entirely
    from row_set.rates (not rates_port_port - no evidence any lane needs
    a port-port equivalent of this sheet).

    Produces ONE flat list, written as one sheet: the commodity block is
    stamped only where cmdt_seq changes, which is what marks a group
    boundary for OPUS.

    Rows are bucketed so each commodity group appears exactly once, as a
    single contiguous block. RATES itself does NOT guarantee that - a
    lane's rows follow raw sheet order and any appended DG duplicates, so
    a group can appear, stop, and resume (TAD WMW/WEW's WEW scope runs
    1,2,3,2,1,2,3,2). That ordering is accepted on RATES, where the tests
    match rows by business identity rather than position, but here the
    blank-fill makes order load-bearing: every real VERTICAL RATES sheet
    seen states each group exactly once (reference/2_OPUS/27's WEW sheet
    restates CMDT Seq 4 times for 4 groups), so a resumed group would be a
    shape OPUS has never been given. Bucketing keeps groups in the order
    they first appear and preserves each row's order within its group.
    """
    buckets: dict[object, list[RatesRow]] = {}
    for row in row_set.rates:
        buckets.setdefault(row.cmdt_seq, []).append(row)

    vertical: list[VerticalRatesRow] = []
    for group_rows in buckets.values():
        for i, row in enumerate(group_rows):
            vertical.extend(explode_to_vertical_rates(row, start_of_group=i == 0))
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
