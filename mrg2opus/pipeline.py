"""The one place a lane parser is actually run.

Parsing is never just `parser.run_multi()` - three steps have to follow
it (commodity-group ordering, the cross-sheet sequence fixups, then the
VERTICAL RATES derivation), and any caller that forgets one silently
produces different output from the others. That happened: the CLI called run_multi directly
and so could never emit a VERTICAL RATES sheet at all, no matter what the
profile said. Both the CLI and the Streamlit wizard go through here now.
"""
from __future__ import annotations

from mrg2opus.parsers.base import BaseMRGParser
from mrg2opus.parsers.common.ordering import drop_commodity_groups, reorder_row_set
from mrg2opus.parsers.common.sequencing import finalize_sequences
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema.opus_rows import OpusRowSet, build_vertical_rates

# OPUS may not accept a VERTICAL RATES upload past this many rows
# (user-stated). Deliberately advisory: the sheet is still generated in
# full and the count reported, so the filer can trim or split it in Excel.
# Nothing here truncates or splits - real ground truth keeps every
# commodity group on ONE sheet, so splitting would produce a shape OPUS
# doesn't expect just to satisfy a limit the filer can handle themselves.
VERTICAL_RATES_ROW_CAP = 10_000


def run_parser(parser: BaseMRGParser, workbook, profile: MappingProfile) -> dict[str, OpusRowSet]:
    row_sets = parser.run_multi(workbook, profile)
    # Before everything else, so the dropped groups are absent from the
    # ordering, the sequence fixups and the VERTICAL RATES derivation
    # alike - that last one is how they stay out of that sheet.
    skipped = {desc for desc, skip in profile.skip_commodity_filing.items() if skip}
    if skipped:
        row_sets = {suffix: drop_commodity_groups(rs, skipped) for suffix, rs in row_sets.items()}
    if profile.commodity_group_order:
        row_sets = {suffix: reorder_row_set(rs, profile.commodity_group_order) for suffix, rs in row_sets.items()}
    # After reordering (blocks move as units, so this stays correct) and
    # before the VERTICAL RATES derivation, which reads rates' cmdt_seq.
    for rs in row_sets.values():
        finalize_sequences(rs)
    if profile.include_vertical_rates:
        row_sets = {suffix: build_vertical_rates(rs) for suffix, rs in row_sets.items()}
    return row_sets


def vertical_rates_over_cap(row_sets: dict[str, OpusRowSet]) -> dict[str, int]:
    """{scope: row count} for every VERTICAL RATES sheet over the cap -
    empty when they all fit. Advisory only; the sheet is written either
    way."""
    return {
        suffix: len(rs.vertical_rates)
        for suffix, rs in row_sets.items()
        if len(rs.vertical_rates) > VERTICAL_RATES_ROW_CAP
    }
