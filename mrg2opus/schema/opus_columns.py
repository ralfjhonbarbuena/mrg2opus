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

SHEET_NAME_RATES = "OPUS RATES"
SHEET_NAME_RATES_PORT_PORT = "OPUS RATES PORT-PORT"
SHEET_NAME_ARBS = "OPUS ARBS"
SHEET_NAME_CMDT_NOTE = "OPUS CMDT NOTE"
SHEET_NAME_SPECIAL_NOTE = "OPUS SPECIAL NOTE"
