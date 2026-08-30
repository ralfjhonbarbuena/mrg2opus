"""Kept as the import path Step 2 / Step 3 already use - the real
implementation moved to mrg2opus.pipeline so the CLI shares it too (it
previously called parser.run_multi directly and silently skipped both
post-parse steps).
"""
from __future__ import annotations

from mrg2opus.pipeline import VERTICAL_RATES_ROW_CAP, run_parser, vertical_rates_over_cap

__all__ = ["run_parser", "vertical_rates_over_cap", "VERTICAL_RATES_ROW_CAP"]
