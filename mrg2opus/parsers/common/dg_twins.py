"""Dangerous-goods twins for reefer and NOR groups, for any lane.

Most lanes' dry groups get a D/DG twin of every D/DR row, built by the
parser itself. Reefer and NOR are different: only LAWC's ground truth
files them as dangerous at all, so no other parser builds them - but the
user reports (2026-09-06) that another MRG can carry them, and a filing
that needs them should not need a code change to produce them.

So the rule lives here, off by default, and the settings turn it on per
commodity group. The two twins are not the same shape:

    Reefer  R/RF  ->  R/DG   an otherwise identical row.
    NOR     R/DR  ->  D/DG   carrying "REEFER DRY AS DANGEROUS".

A reefer box carrying dry cargo that happens to be dangerous is filed as
dry, which is what the note exists to explain; it joins any note the row
already has with " | ". Both stay in the commodity group they came from.
Confirmed against reference/2_OPUS/15_LAWC FAK, whose G0004 "NOR & REEFER"
holds all four of R/RF, R/DG, R/DR and D/DG.

LAWC builds its own (they are that lane's default, not an opt-in), and
this module never doubles them: an identical row already in the set is
left alone rather than filed twice, which OPUS rejects outright.
"""
from __future__ import annotations

from mrg2opus.schema.opus_rows import OpusRowSet, RouteNoteRow

REEFER_DRY_AS_DANGEROUS = "REEFER DRY AS DANGEROUS"

_REEFER = ("R", "RF")
_NOR = ("R", "DR")


def dangerous_route_note(existing: str | None) -> str:
    return f"{REEFER_DRY_AS_DANGEROUS} | {existing}" if existing else REEFER_DRY_AS_DANGEROUS


def reefer_and_nor_groups(row_sets: dict[str, OpusRowSet]) -> frozenset[str]:
    """The commodity groups holding reefer or NOR rows.

    These are the groups a DG twin can be asked for even where the lane
    doesn't file one by default - so they get a Skip DG checkbox, ticked,
    rather than no checkbox at all.
    """
    return frozenset(
        row.commodity_group_description
        for row_set in row_sets.values()
        for row in row_set.rates
        if (row.prefix, row.cgo_type) in (_REEFER, _NOR)
    )


def _twin(row):
    """The dangerous counterpart of one reefer or NOR row, or None if the
    row is neither."""
    if (row.prefix, row.cgo_type) == _REEFER:
        return row.model_copy(update={"cgo_type": "DG"})
    if (row.prefix, row.cgo_type) == _NOR:
        return row.model_copy(
            update={"prefix": "D", "cgo_type": "DG", "route_note": dangerous_route_note(row.route_note)}
        )
    return None


def _filing_key(row) -> tuple:
    """What makes two rows the same filing. Sequence numbers are left out
    (nothing has assigned them yet at this point) and so are the rates -
    two rows for one route at different rates are still one duplicate
    filing as far as OPUS is concerned."""
    return (
        row.commodity_group_description, row.origin_code, row.destination_code,
        row.prefix, row.cgo_type, row.o_via_code, row.d_via_code,
    )


def add_reefer_and_nor_dg_twins(row_set: OpusRowSet, groups: frozenset[str]) -> OpusRowSet:
    """`row_set` with a DG twin appended for each reefer/NOR row in
    `groups`. Returns the row set unchanged when nothing is enabled.

    Twins go in right after the rows they came from, so they read the way
    a filing does rather than in a block at the end.
    """
    if not groups:
        return row_set

    updates: dict[str, list] = {}
    for field in ("rates", "rates_port_port"):
        rows = getattr(row_set, field)
        # PORT-PORT commodity names are remapped by some lanes (LAWC files
        # its reefer and NOR rows as one "NOR & REEFER" group there), so a
        # group enabled on RATES may not be found here - in which case that
        # sheet simply gets no twin. LAWC, the one lane that both remaps
        # and files these twins, builds its own on both sheets.
        existing = {_filing_key(r) for r in rows}
        out = []
        for row in rows:
            out.append(row)
            if row.commodity_group_description not in groups:
                continue
            twin = _twin(row)
            if twin is None or _filing_key(twin) in existing:
                continue
            existing.add(_filing_key(twin))
            out.append(twin)
        if len(out) != len(rows):
            updates[field] = out
    return row_set.model_copy(update=updates) if updates else row_set


def add_missing_dangerous_route_notes(row_set: OpusRowSet) -> None:
    """Give every "REEFER DRY AS DANGEROUS" rates row an RN entry, in place.

    A route note on a RATES row addresses an RN row by (header_seq,
    route_seq), so a note with no RN entry behind it says nothing. Those
    numbers are assigned by the sequencing pass, which is why this runs
    after it rather than beside add_reefer_and_nor_dg_twins.

    Only this one note text is topped up. Every other route note is its
    lane's own business, and a lane that deliberately leaves one off the
    RN sheet should keep doing so.
    """
    covered = {(n.header_seq, n.route_seq, n.contents) for n in row_set.route_notes}
    # The dates the rest of the filing uses for this commodity group. Read
    # off its CMDT NOTE block, which is where a group's validity already
    # lives - there is nothing else to derive them from here.
    dates_by_header = {
        n.header_seq: (n.application_effective, n.application_expires)
        for n in row_set.cmdt_notes
        if n.header_seq is not None
    }
    for row in row_set.rates:
        note = row.route_note or ""
        if not note.startswith(REEFER_DRY_AS_DANGEROUS):
            continue
        key = (row.cmdt_seq, row.route_seq, note)
        if key in covered:
            continue
        covered.add(key)
        start, end = dates_by_header.get(row.cmdt_seq, (None, None))
        row_set.route_notes.append(
            RouteNoteRow(
                header_seq=row.cmdt_seq,
                route_seq=row.route_seq,
                note_seq=1,
                contents=note,
                charge_seq=1,
                code="APP",
                application="S",
                application_effective=start,
                application_expires=end,
            )
        )
