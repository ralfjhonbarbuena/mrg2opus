"""Phase 1 CLI entrypoint (no UI yet - that's Phase 2's Streamlit wizard).

    python -m mrg2opus.cli parse "path/to/MRG.xlsx" --out out.xlsx

Accepts more than one input file for lanes that ship as multiple
real-world files (e.g. CSE's main file plus a separate "...for VELAG and
VEPBL" file) - they're merged into one workbook before classification,
same as the web UI's multi-file upload (see excel_io/merge.py):

    python -m mrg2opus.cli parse "CSE Tier 1.xlsx" "CSE Tier 1 for VELAG and VEPBL.xlsx" --out out.xlsx
"""
from __future__ import annotations

import argparse
from pathlib import Path

import openpyxl

# Importing the lane modules registers their LayoutProfile as a side effect.
from mrg2opus.parsers import (  # noqa: F401
    aubp, auec, auwc, cse, eaf, laec, lawc, nz1_sea, nzj, saf, tad_aew_amw, tad_oew_omw, tad_wmw_wew, waf, west_asia_multi, west_asia_waf,
)
from mrg2opus.parsers.registry import classify
from mrg2opus.excel_io.merge import merge_workbooks
from mrg2opus.excel_io.writer import write_opus_workbook_multi
from mrg2opus.pipeline import VERTICAL_RATES_ROW_CAP, run_parser, vertical_rates_over_cap
from mrg2opus.presets.models import MappingProfile


def parse_command(args: argparse.Namespace) -> None:
    workbooks = [openpyxl.load_workbook(path, data_only=True) for path in args.input]
    wb = merge_workbooks(workbooks)
    result = classify(wb)
    print(f"Classified as lane={result.profile.lane_id} confidence={result.confidence:.2f} breakdown={result.breakdown}")

    parser_cls = result.profile.parser_cls
    row_sets = run_parser(parser_cls(), wb, MappingProfile())

    write_opus_workbook_multi(
        row_sets,
        args.out,
        sheet_name_overrides=parser_cls.SHEET_NAME_OVERRIDES,
        scoped_sheet_name_overrides=parser_cls.SCOPED_SHEET_NAME_OVERRIDES,
    )
    for suffix, row_set in row_sets.items():
        label = suffix or "(default)"
        print(
            f"[{label}] {len(row_set.rates)} RATES, {len(row_set.rates_port_port)} RATES PORT-PORT, "
            f"{len(row_set.arbs)} ARBS, {len(row_set.cmdt_notes)} CMDT NOTE, "
            f"{len(row_set.special_notes)} SPECIAL NOTE, "
            f"{len(row_set.route_notes)} ROUTE NOTE, {len(row_set.vertical_rates)} VERTICAL RATES rows"
        )
    for (suffix, cmdt_seq), count in vertical_rates_over_cap(row_sets).items():
        print(
            f"WARNING: [{suffix or '(default)'}] VERTICAL RATES {cmdt_seq} has {count:,} rows, over the "
            f"{VERTICAL_RATES_ROW_CAP:,}-row OPUS upload limit. It is already one sheet per CMDT Seq, so "
            "this commodity group is too big to upload as vertical rates - split it or use RATES instead."
        )
    print(f"Wrote {args.out}")


def main() -> None:
    ap = argparse.ArgumentParser(prog="mrg2opus")
    sub = ap.add_subparsers(dest="command", required=True)

    parse_ap = sub.add_parser("parse", help="Convert one or more raw MRG Excel files to OPUS format")
    parse_ap.add_argument("input", type=Path, nargs="+", help="One or more raw MRG .xlsx files to merge and parse")
    parse_ap.add_argument("--out", type=Path, required=True)
    parse_ap.set_defaults(func=parse_command)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
