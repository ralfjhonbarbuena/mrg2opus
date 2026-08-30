"""The one place a lane parser is actually run.

Parsing is never just `parser.run_multi()` - two profile-driven steps have
to follow it (commodity-group ordering, then the VERTICAL RATES
derivation), and any caller that forgets one silently produces different
output from the others. That happened: the CLI called run_multi directly
and so could never emit a VERTICAL RATES sheet at all, no matter what the
profile said. Both the CLI and the Streamlit wizard go through here now.
"""
from __future__ import annotations

from mrg2opus.parsers.base import BaseMRGParser
from mrg2opus.parsers.common.ordering import reorder_row_set
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema.opus_rows import OpusRowSet, build_vertical_rates, group_vertical_rates_by_cmdt_seq

# OPUS rejects an upload past this many rows on one sheet (user-stated).
# VERTICAL RATES is written one sheet per CMDT Seq (see excel_io/writer.py),
# so the number that matters is the biggest SINGLE group, not the lane
# total - splitting already keeps every other lane under the limit. A
# single group over the cap is the one case splitting can't fix, so it is
# reported rather than silently truncated: how to break up a commodity
# group is a filing decision, not something to guess at.
VERTICAL_RATES_ROW_CAP = 10_000


def run_parser(parser: BaseMRGParser, workbook, profile: MappingProfile) -> dict[str, OpusRowSet]:
    row_sets = parser.run_multi(workbook, profile)
    if profile.commodity_group_order:
        row_sets = {suffix: reorder_row_set(rs, profile.commodity_group_order) for suffix, rs in row_sets.items()}
    if profile.include_vertical_rates:
        row_sets = {suffix: build_vertical_rates(rs) for suffix, rs in row_sets.items()}
    return row_sets


def vertical_rates_over_cap(row_sets: dict[str, OpusRowSet]) -> dict[tuple[str, int | None], int]:
    """{(scope, cmdt_seq): row count} for every VERTICAL RATES sheet still
    over the cap after the per-CMDT-Seq split - empty when they all fit,
    which is the normal case."""
    over: dict[tuple[str, int | None], int] = {}
    for suffix, rs in row_sets.items():
        for cmdt_seq, rows in group_vertical_rates_by_cmdt_seq(rs.vertical_rates):
            if len(rows) > VERTICAL_RATES_ROW_CAP:
                over[(suffix, cmdt_seq)] = len(rows)
    return over
