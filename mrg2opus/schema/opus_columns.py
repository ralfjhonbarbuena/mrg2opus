"""Canonical OPUS sheet header text and column order.

This is the single source of truth for how each OPUS sheet's header row(s)
are rendered. Parsers and the writer both import from here so the header
text never drifts out of sync between what we read/expect and what we emit.
"""

# --- OPUS RATES / OPUS RATES PORT-PORT: genuine 2-row header -----------------
# Row 1 is a "group" label repeated across the columns it spans (not a merged
# cell in the source samples - just literal repeated/blank text per column).
# Row 2 is the field label under each group.

RATES_HEADER_GROUP = [
    "Type", "CMDT\nSeq.",
    "Commodity Group", "Commodity Group",
    "Actual Customer", "Actual Customer",
    "Route\nSeq.",
    "Origin", "Origin", "Origin", "Origin",
    "O.Via", "D.Via",
    "Destination", "Destination", "Destination", "Destination",
    "Rate", "Rate", "Rate", "Rate", "Rate", "Rate", "Rate", "Rate", "Rate", "Rate",
    "Commodity Note", "Route Note",
]

RATES_HEADER_FIELD = [
    None, None,
    "Code", "Description",
    "Code", "Description",
    None,
    "Code", "Description", "Term", "Transmode",
    "Code", " Code",
    "Code", "Description", "Term", "Transmode",
    "Prefix", "CGO TYPE",
    "Cur", "20", "Cur", "40", "Cur", "40HC", "Cur", "45",
    None, None,
]

# Field names on RatesRow, in the same column order as the two header rows
# above. Used by the writer to pull values out of the pydantic model in order.
RATES_ROW_FIELDS = [
    "type", "cmdt_seq",
    "commodity_group_code", "commodity_group_description",
    "actual_customer_code", "actual_customer_description",
    "route_seq",
    "origin_code", "origin_description", "origin_term", "origin_transmode",
    "o_via_code", "d_via_code",
    "destination_code", "destination_description", "destination_term", "destination_transmode",
    "prefix", "cgo_type",
    "cur_20", "rate_20", "cur_40", "rate_40", "cur_40hc", "rate_40hc", "cur_45", "rate_45",
    "commodity_note", "route_note",
]

# --- VERTICAL RATES: OPUS's alternate "long format" rate upload, one row ----
# per container size instead of RATES' 4-slots-per-row layout. Verified
# directly against reference/2_OPUS/25_TAD FILING OEW OMW's real sheet - a
# general OPUS feature (not TAD-specific, per the user), so this lives
# alongside RATES_HEADER_* rather than under a TAD-only module.

VERTICAL_RATES_HEADER_GROUP = [
    "CMDT\nSeq.",
    "Commodity Group", "Commodity Group",
    "Actual Customer", "Actual Customer",
    "Route\nSeq.",
    "Origin", "Origin", "Origin", "Origin",
    "O.Via", "D.Via",
    "Destination", "Destination", "Destination", "Destination",
    "Rate(USD)", "Rate(USD)", "Rate(USD)",
]

VERTICAL_RATES_HEADER_FIELD = [
    None,
    "Code", "Description",
    "Code", "Description",
    None,
    "Code", "Description", "Term", "Transmode",
    "Code", None,
    None, "Description", "Term", "Transmode",
    "PER", "Cargo Type", "Rate",
]

VERTICAL_RATES_ROW_FIELDS = [
    "cmdt_seq",
    "commodity_group_code", "commodity_group_description",
    "actual_customer_code", "actual_customer_description",
    "route_seq",
    "origin_code", "origin_description", "origin_term", "origin_transmode",
    "o_via_code", "d_via_code",
    "destination_code", "destination_description", "destination_term", "destination_transmode",
    "per", "cargo_type", "rate",
]

# --- OPUS ARBS: single header row --------------------------------------------
# Verified directly against CSE.xlsx's OPUS ARBS sheet (26 columns).

ARBS_HEADER = [
    "Seq.", "Point(*)", "Description", "Trans Mode", "Term(*)", "Service\nLane", "Trunk\nLane",
    "Weight\n(=> Metric Ton)", "Weight\n(< Metric Ton)", "Over(*)", "VIA", "Actual\nCustomer",
    "Pay Term", "Per(*)", "CGO\nType", "Cur.(*)", "Proposal(*)", "C.Offer", "Final",
    "Actual\nEFF Date(*)", "Actual\nEXP Date(*)", "Source", "Status", "seq", "Note", "Remark",
]

ARBS_ROW_FIELDS = [
    "seq", "point", "description", "trans_mode", "term", "service_lane", "trunk_lane",
    "weight_gte_mt", "weight_lt_mt", "over", "via", "actual_customer", "pay_term", "per",
    "cgo_type", "cur", "proposal", "c_offer", "final", "eff_date", "exp_date", "source",
    "status", "seq2", "note", "remark",
]

# --- OPUS CMDT NOTE / OPUS SPECIAL NOTE: single header row -------------------

CMDT_NOTE_HEADER = [
    "Header\nSeq", "Note\nSeq", "Contents", "Charge \nSeq", "Code",
    "Application\nEffective", "Application\nExpires", "Application",
    "Cur.", "Cal.", "Amount", "Pay Term", "Pay Ofc", "Payer", "Per", "CGO\nType",
    "IMDG\nClass", "PSA Grp.", "Food\nGrade", "Lane", "T/S\nPort", "Canal", "VVD", "SOC",
    "POR", "POL", "POD", "DEL", "Node", "CMDT",
]

CMDT_NOTE_ROW_FIELDS = [
    "header_seq", "note_seq", "contents", "charge_seq", "code",
    "application_effective", "application_expires", "application",
    "cur", "cal", "amount", "pay_term", "pay_ofc", "payer", "per", "cgo_type",
    "imdg_class", "psa_grp", "food_grade", "lane", "ts_port", "canal", "vvd", "soc",
    "por", "pol", "pod", "delivery", "node", "cmdt",
]

# OPUS SPECIAL NOTE has the same field set as CMDT NOTE (SpecialNoteRow
# subclasses CmdtNoteRow) but a DIFFERENT column order - verified directly
# against CSE.xlsx: "Application" comes right after "Code(*)", before the
# effective/expires dates (CMDT NOTE has Application last, after the dates).
SPECIAL_NOTE_HEADER = [
    "Note\nSeq", "Note Ctnt\nSeq", "Contents", "Charge \nSeq", "Code(*)",
    "Application", "Application\nEffective(*)", "Application\nExpires(*)",
    "Cur.", "Cal.", "Amount", "Pay Term", "Pay Ofc", "Payer", "Per", "CGO\nType",
    "IMDG\nClass", "PSA Grp.", "Food\nGrade", "Lane", "T/S\nPort", "Canal", "VVD", "SOC",
    "POR", "POL", "POD", "DEL", "Node", "CMDT",
]

SPECIAL_NOTE_ROW_FIELDS = [
    "header_seq", "note_seq", "contents", "charge_seq", "code",
    "application", "application_effective", "application_expires",
    "cur", "cal", "amount", "pay_term", "pay_ofc", "payer", "per", "cgo_type",
    "imdg_class", "psa_grp", "food_grade", "lane", "ts_port", "canal", "vvd", "soc",
    "por", "pol", "pod", "delivery", "node", "cmdt",
]

# --- RN (ROUTE NOTE): single header row --------------------------------------
# Verified directly against reference/2_OPUS/15_LAWC FAK's real "RN" sheet -
# same shape as CMDT_NOTE_HEADER (Header Seq through CMDT) plus a Route Seq
# column (2nd position) and 10 trailing columns (Receiving Term through
# Premium) that CMDT_NOTE_HEADER above does NOT have - CMDT_NOTE_HEADER was
# verified against CSE.xlsx's older bundled sample, which appears to have a
# shorter/different real header; not reconciled here, out of scope for the
# ROUTE NOTE feature this was built for.
RN_HEADER = [
    "Header\nSeq", "Route\nSeq", "Note\nSeq", "Contents", "Charge \nSeq", "Code",
    "Application\nEffective", "Application\nExpires", "Application",
    "Cur.", "Cal.", "Amount", "Pay Term", "Pay Ofc", "Payer", "Per", "CGO\nType",
    "IMDG\nClass", "PSA Grp.", "Food\nGrade", "Lane", "T/S\nPort", "Canal", "VVD", "SOC",
    "POR", "POL", "POD", "DEL", "Node", "CMDT",
    "Receiving\nTerm", "Delivery\nTerm", "Weight\n(=> Metric Ton)", "Weight\n( < Metric Ton)",
    "Direct\nCall", "Bar Type", "S/I", "M'ty Pick up CY", "M'ty Return CY", "Premium",
]

RN_ROW_FIELDS = [
    "header_seq", "route_seq", "note_seq", "contents", "charge_seq", "code",
    "application_effective", "application_expires", "application",
    "cur", "cal", "amount", "pay_term", "pay_ofc", "payer", "per", "cgo_type",
    "imdg_class", "psa_grp", "food_grade", "lane", "ts_port", "canal", "vvd", "soc",
    "por", "pol", "pod", "delivery", "node", "cmdt",
    "receiving_term", "delivery_term", "weight_gte_mt", "weight_lt_mt",
    "direct_call", "bar_type", "s_i", "mty_pickup_cy", "mty_return_cy", "premium",
]

# --- FREETIME ----------------------------------------------------------------
# The 46-column shape below is what a DOWNLOADED OPUS filing looks like -
# verified directly against reference/2_OPUS/15_LAWC FAK's and 19_LAEC FAK's
# real "FREETIME" sheets, a per-lane STATIC reference table (RFA/tariff
# free-time-allowance schedule) not derived from the raw MRG at all
# (confirmed byte-identical across every LAWC ground truth sample seen, and
# identical-apart-from-EFF/EXP-DT across every LAEC one) - see
# parsers/common/freetime.py for the concrete per-lane tables.
#
# We generate an UPLOAD, not a download, so the columns OPUS assigns on its
# own side are dropped on the way out (user direction, 2026-08-31): "RFA
# No.", "Status", and everything from column AO ("DAR No.") rightwards.
# They're kept on FreetimeRow, and in the FREETIME_FULL_* lists here,
# because every reference/2_OPUS FREETIME sheet still carries all 46 -
# that's the shape the ground-truth tests read back. FREETIME_* (no
# FULL) is the narrower shape the writer actually emits.
FREETIME_FULL_HEADER_GROUP = [
    "Seq.", "RFA No.", "Status", "Tariff", "EFF DT", "EXP DT", "CNTR/Cargo", "IMDG\nClass", "PSA Grp.",
    "Coverage", "Coverage", "Coverage", "Free Time", "Free Time", "Free Time",
    "F/Time EXCL", "F/Time EXCL", "F/Time EXCL",
    "Origin(I) or Dest.(O)", "Origin(I) or Dest.(O)", "Origin(I) or Dest.(O)", "Origin(I) or Dest.(O)",
    "BKG DEL(I) or POR(O)", "BKG DEL(I) or POR(O)", "BKG DEL(I) or POR(O)",
    "Actual Customer", "Actual Customer", "Commodity", "Commodity", "Curr",
    "Over Day", "Over Day", "Rate per Day", "Rate per Day", "Rate per Day", "Rate per Day",
    "CNTR Q'TY", "CNTR Q'TY", "Tiered\n Free\n Time", "Remark",
    "DAR No.", "Ver.", "Approval No.", "Proposal No.", "Customer", "Customer",
]
FREETIME_FULL_HEADER_FIELD = [
    None, None, None, None, None, None, None, None, None,
    "CN", "RGN", "LOC", "Tier", "Add", "Total", "SAT", "SUN", "H/day",
    "CT", "CN", "RGN", "LOC", "CN", "RGN", "LOC", "Code", "Name", "Code", "Name", None,
    "From", "Up to", "20 FT", "40 FT", "H/C", "45 FT",
    "From", "Up to", None, None, None, None, None, None, "Code", "Name",
]
FREETIME_FULL_ROW_FIELDS = [
    "seq", "rfa_no", "status", "tariff", "eff_dt", "exp_dt", "cntr_cargo", "imdg_class", "psa_grp",
    "coverage_cn", "coverage_rgn", "coverage_loc", "free_time_tier", "free_time_add", "free_time_total",
    "ftime_excl_sat", "ftime_excl_sun", "ftime_excl_hday",
    "origin_or_dest_ct", "origin_or_dest_cn", "origin_or_dest_rgn", "origin_or_dest_loc",
    "bkg_del_cn", "bkg_del_rgn", "bkg_del_loc",
    "actual_customer_code", "actual_customer_name", "commodity_code", "commodity_name", "curr",
    "over_day_from", "over_day_upto", "rate_per_day_20", "rate_per_day_40", "rate_per_day_hc", "rate_per_day_45",
    "cntr_qty_from", "cntr_qty_upto", "tiered_free_time", "remark",
    "dar_no", "ver", "approval_no", "proposal_no", "customer_code", "customer_name",
]

# Assigned by OPUS itself (the RFA/DAR paperwork identifiers and the
# approval status), so an upload must leave them off entirely rather than
# send back the reference table's values.
FREETIME_UNFILED_FIELDS = frozenset({
    "rfa_no", "status",
    "dar_no", "ver", "approval_no", "proposal_no", "customer_code", "customer_name",
})

_FREETIME_KEPT = [i for i, f in enumerate(FREETIME_FULL_ROW_FIELDS) if f not in FREETIME_UNFILED_FIELDS]
FREETIME_HEADER_GROUP = [FREETIME_FULL_HEADER_GROUP[i] for i in _FREETIME_KEPT]
FREETIME_HEADER_FIELD = [FREETIME_FULL_HEADER_FIELD[i] for i in _FREETIME_KEPT]
FREETIME_ROW_FIELDS = [FREETIME_FULL_ROW_FIELDS[i] for i in _FREETIME_KEPT]

# --- Legacy bundled-sample sheet names ---------------------------------------
# These match the literal sheet names inside the older, hand-prepared
# "Sample MRGs with OPUS FORMATS/*.xlsx" fixtures every tests/test_parsers_*.py
# golden test reads from - NOT the real OPUS filing convention (see
# excel_io/writer.py's _sheet_names_for_suffix for that, and the
# project-opus-note-sheet-taxonomy memory for the full naming story). Do not
# change these values - every lane's regression test depends on them matching
# those static files' actual sheet names.
SHEET_NAME_RATES = "OPUS RATES"
SHEET_NAME_RATES_PORT_PORT = "OPUS RATES PORT-PORT"
SHEET_NAME_ARBS = "OPUS ARBS"
SHEET_NAME_CMDT_NOTE = "OPUS CMDT NOTE"
SHEET_NAME_SPECIAL_NOTE = "OPUS SPECIAL NOTE"
