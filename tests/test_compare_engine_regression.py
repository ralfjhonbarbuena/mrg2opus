"""Regression check: run the NEW production compare engine (audit/compare.py,
not tests/golden.py's own helpers) against each bundled sample's own
ground-truth RATES sheet. Every row that should exist must exist on both
sides - the same row-level guarantee test_parsers_*.py already asserts
via `assert not missing` / `assert not extra`, re-derived through the new
production code path as a consistency check on top of what those tests
already cover.

Deliberately does NOT assert zero field_mismatches - several lanes have
documented, accepted field-level gaps (see each lane's own test file
comments, e.g. test_parsers_cse.py's PAMIT rate_20 note). This test
checks row existence, not the field-level nuances the existing
test_parsers_*.py files already own and assert on directly.
"""
from __future__ import annotations

import openpyxl
import pytest

from mrg2opus.audit.compare import diff_by_key, rates_row_key, read_rates_sheet
from mrg2opus.parsers.cse import CSEParser
from mrg2opus.parsers.eaf import EAFParser
from mrg2opus.parsers.laec import LAECParser
from mrg2opus.parsers.lawc import LAWCParser
from mrg2opus.parsers.saf import SAFParser
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols


@pytest.mark.parametrize(
    "path,parser_cls",
    [
        ("Sample MRGs with OPUS FORMATS/SAF.xlsx", SAFParser),
        ("Sample MRGs with OPUS FORMATS/EAF.xlsx", EAFParser),
        ("Sample MRGs with OPUS FORMATS/CSE.xlsx", CSEParser),
        ("Sample MRGs with OPUS FORMATS/LAEC.xlsx", LAECParser),
        ("Sample MRGs with OPUS FORMATS/LAWC.xlsx", LAWCParser),
    ],
)
def test_compare_engine_finds_no_missing_or_extra_rates_rows_against_own_ground_truth(path, parser_cls):
    wb = openpyxl.load_workbook(path, data_only=True)
    parser = parser_cls()
    row_sets = parser.run_multi(wb, MappingProfile())
    for suffix, row_set in row_sets.items():
        tag = f"-{suffix}" if suffix else ""
        generated = [r.model_dump() for r in row_set.rates]
        expected = read_rates_sheet(wb, f"OPUS RATES{tag}")

        result = diff_by_key(generated, expected, key_fn=rates_row_key, fields=cols.RATES_ROW_FIELDS)
        assert not result.missing, f"{path} [{suffix or '(default)'}]: missing {len(result.missing)} rows"
        assert not result.extra, f"{path} [{suffix or '(default)'}]: {len(result.extra)} unexpected rows"
