"""Standalone MRG-vs-reference-OPUS comparison mode - separate from the
5-file-then-4-step wizard. Upload one or more MRG files, upload a
reference OPUS-format Excel file, choose which RATES form(s) to check,
and see where the parser's own output diverges from the reference.

See docs/superpowers/specs/2026-08-23-mrg-opus-comparison-design.md.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import Any

import openpyxl
import streamlit as st
from openpyxl.workbook import Workbook

from mrg2opus.audit.compare import (
    CMDT_NOTE_IGNORE_FIELDS_BY_LANE,
    RATES_IGNORE_FIELDS_BY_LANE,
    RATES_PORT_PORT_IGNORE_FIELDS_BY_LANE,
    SPECIAL_NOTE_IGNORE_FIELDS_BY_LANE,
    arbs_row_key,
    diff_by_key,
    diff_cmdt_blocks,
    rates_row_key,
    read_arbs_sheet,
    read_cmdt_note_sheet,
    read_rates_sheet,
    read_special_note_sheet,
    reconstruct_blocks,
)
from mrg2opus.excel_io.merge import DuplicateSheetError
from mrg2opus.parsers.registry import ClassificationResult, get_profile
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols
from mrg2opus.ui.mrg_upload import fingerprint_uploads, load_and_classify
from mrg2opus.ui.parsing import run_parser

RATES_MODE_OPTIONS = ["Grouped (RATES)", "Exploded (RATES PORT-PORT)", "Both"]


@dataclass
class CompareState:
    upload_key: str | None = None
    workbook: Workbook | None = None
    classification_results: list[ClassificationResult] = field(default_factory=list)
    selected_lane_id: str | None = None
    reference_workbook: Workbook | None = None
    rates_mode: str = "Both"
    apply_known_gaps: bool = True
    row_sets: dict[str, Any] | None = None
    compare_results: list[dict[str, Any]] | None = None


def _get_state() -> CompareState:
    if "compare" not in st.session_state:
        st.session_state.compare = CompareState()
    return st.session_state.compare


def _compare_keyed_sheet(sheet_type, suffix, sheet_name, generated, ref_wb, key_fn, fields, reader, ignore_fields=frozenset()) -> dict | None:
    sub_lane = suffix or "(default)"
    try:
        expected = reader(ref_wb, sheet_name)
    except KeyError:
        if not generated:
            return None
        return {
            "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
            "found_in_reference": False, "matched": 0,
            "missing": [], "extra": sorted(({"key": key_fn(r)} for r in generated), key=lambda d: str(d["key"])),
            "field_mismatches": [],
        }
    if not generated and not expected:
        return None
    result = diff_by_key(generated, expected, key_fn=key_fn, fields=fields, ignore_fields=ignore_fields)
    return {
        "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
        "found_in_reference": True, "matched": result.matched,
        "missing": sorted(({"key": k} for k in result.missing), key=lambda d: str(d["key"])),
        "extra": sorted(({"key": k} for k in result.extra), key=lambda d: str(d["key"])),
        "field_mismatches": sorted(
            (
                {"key": m[0], "field": m[1], "generated": m[2], "reference": m[3]}
                for m in result.field_mismatches
            ),
            key=lambda d: (str(d["key"]), d["field"]),
        ),
    }


def _compare_block_sheet(sheet_type, suffix, sheet_name, generated, ref_wb, fields, reader, ignore_fields=frozenset()) -> dict | None:
    sub_lane = suffix or "(default)"
    try:
        expected = reader(ref_wb, sheet_name)
    except KeyError:
        extra_keys = sorted({b.key for b in reconstruct_blocks(generated)})
        if not extra_keys:
            return None
        return {
            "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
            "found_in_reference": False, "matched": None,
            "missing": [], "extra": [{"contents": k} for k in extra_keys],
            "field_mismatches": [],
        }
    gen_has_blocks = bool(reconstruct_blocks(generated))
    exp_has_blocks = bool(reconstruct_blocks(expected))
    if not gen_has_blocks and not exp_has_blocks:
        return None
    result = diff_cmdt_blocks(generated, expected, fields, ignore_fields=ignore_fields)
    return {
        "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
        "found_in_reference": True, "matched": None,
        "missing": [{"contents": k} for k in result.missing_blocks],
        "extra": [{"contents": k} for k in result.extra_blocks],
        "field_mismatches": sorted(
            (
                {"key": m[0], "child_index": m[1], "field": m[2], "generated": m[3], "reference": m[4]}
                for m in result.field_mismatches
            ),
            key=lambda d: (d["key"], d["child_index"], d["field"]),
        ),
    }


def _run_comparison(row_sets: dict, ref_wb: Workbook, rates_mode: str, lane_id: str, apply_known_gaps: bool = True) -> list[dict]:
    want_grouped = rates_mode in ("Grouped (RATES)", "Both")
    want_exploded = rates_mode in ("Exploded (RATES PORT-PORT)", "Both")

    rates_ignore = RATES_IGNORE_FIELDS_BY_LANE.get(lane_id, frozenset()) if apply_known_gaps else frozenset()
    port_port_ignore = RATES_PORT_PORT_IGNORE_FIELDS_BY_LANE.get(lane_id, frozenset()) if apply_known_gaps else frozenset()
    cmdt_ignore = CMDT_NOTE_IGNORE_FIELDS_BY_LANE.get(lane_id, frozenset()) if apply_known_gaps else frozenset()
    special_ignore = SPECIAL_NOTE_IGNORE_FIELDS_BY_LANE.get(lane_id, frozenset()) if apply_known_gaps else frozenset()

    results: list[dict] = []
    for suffix, row_set in row_sets.items():
        tag = f"-{suffix}" if suffix else ""
        if want_grouped:
            r = _compare_keyed_sheet(
                "RATES", suffix, f"OPUS RATES{tag}",
                [x.model_dump() for x in row_set.rates], ref_wb,
                rates_row_key, cols.RATES_ROW_FIELDS, read_rates_sheet, ignore_fields=rates_ignore,
            )
            if r is not None:
                results.append(r)
        if want_exploded:
            r = _compare_keyed_sheet(
                "RATES PORT-PORT", suffix, f"OPUS RATES{tag} PORT-PORT",
                [x.model_dump() for x in row_set.rates_port_port], ref_wb,
                rates_row_key, cols.RATES_ROW_FIELDS, read_rates_sheet, ignore_fields=port_port_ignore,
            )
            if r is not None:
                results.append(r)
        r = _compare_keyed_sheet(
            "ARBS", suffix, f"OPUS ARBS{tag}",
            [x.model_dump() for x in row_set.arbs], ref_wb,
            arbs_row_key, cols.ARBS_ROW_FIELDS, read_arbs_sheet,
        )
        if r is not None:
            results.append(r)
        r = _compare_block_sheet(
            "CMDT NOTE", suffix, f"OPUS CMDT NOTE{tag}",
            [x.model_dump() for x in row_set.cmdt_notes], ref_wb,
            cols.CMDT_NOTE_ROW_FIELDS, read_cmdt_note_sheet, ignore_fields=cmdt_ignore,
        )
        if r is not None:
            results.append(r)
        r = _compare_block_sheet(
            "SPECIAL NOTE", suffix, f"OPUS SPECIAL NOTE{tag}",
            [x.model_dump() for x in row_set.special_notes], ref_wb,
            cols.SPECIAL_NOTE_ROW_FIELDS, read_special_note_sheet, ignore_fields=special_ignore,
        )
        if r is not None:
            results.append(r)
    return results


def _render_results(results: list[dict]) -> None:
    if not results:
        st.info("Nothing to compare - the parsed MRG produced no rows for the sheet type(s) selected.")
        return

    st.markdown("#### Comparison summary")
    st.dataframe(
        [
            {
                "Sheet": r["sheet_type"],
                "Sub-lane": r["sub_lane"],
                "In reference?": "Yes" if r["found_in_reference"] else "No - sheet not found",
                "Matched": r["matched"] if r["matched"] is not None else "-",
                "Missing": len(r["missing"]),
                "Extra": len(r["extra"]),
                "Field mismatches": len(r["field_mismatches"]),
            }
            for r in results
        ],
        hide_index=True,
        width="stretch",
    )

    st.markdown("#### Details")
    for r in results:
        label = f"{r['sheet_type']} — {r['sub_lane']} ({r['sheet_name']})"
        with st.expander(label):
            if not r["found_in_reference"]:
                st.warning(
                    f"Reference workbook has no sheet matching **{r['sheet_name']}** - "
                    "every generated row is listed as extra."
                )
            if r["missing"]:
                st.markdown(f"**Missing** ({len(r['missing'])}, in reference but not generated)")
                st.dataframe(r["missing"][:50], hide_index=True, width="stretch")
            if r["extra"]:
                st.markdown(f"**Extra** ({len(r['extra'])}, generated but not in reference)")
                st.dataframe(r["extra"][:50], hide_index=True, width="stretch")
            if r["field_mismatches"]:
                st.markdown(f"**Field mismatches** ({len(r['field_mismatches'])})")
                st.dataframe(r["field_mismatches"][:50], hide_index=True, width="stretch")
            if r["found_in_reference"] and not r["missing"] and not r["extra"] and not r["field_mismatches"]:
                st.success("No differences found.")


def render() -> None:
    state = _get_state()
    st.subheader("Compare: MRG vs. reference OPUS file")
    st.caption(
        "Upload the raw MRG rate sheet(s) and an existing OPUS-format Excel file "
        "(e.g. a filing someone already produced) to see where they diverge."
    )

    mrg_files = st.file_uploader(
        "Raw MRG rate sheet(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="compare_mrg_upload"
    )
    reference_file = st.file_uploader(
        "Reference OPUS-format Excel file (.xlsx)", type=["xlsx"], key="compare_reference_upload"
    )

    if not mrg_files or reference_file is None:
        st.info("Upload both the MRG file(s) and a reference OPUS file to continue.")
        return

    mrg_names = [f.name for f in mrg_files]
    mrg_payloads = [f.getvalue() for f in mrg_files]
    reference_payload = reference_file.getvalue()

    upload_key = fingerprint_uploads(mrg_names + [reference_file.name], mrg_payloads + [reference_payload])
    if upload_key != state.upload_key:
        state.upload_key = upload_key
        state.workbook = None
        state.classification_results = []
        state.selected_lane_id = None
        state.reference_workbook = None
        state.row_sets = None
        state.compare_results = None

    if state.workbook is None:
        try:
            state.workbook, state.classification_results = load_and_classify(mrg_payloads)
        except DuplicateSheetError as exc:
            st.error(str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - surfaced directly to the user, not swallowed
            st.error(f"Couldn't open one of the MRG files: {exc}")
            return

    if state.reference_workbook is None:
        try:
            state.reference_workbook = openpyxl.load_workbook(io.BytesIO(reference_payload), data_only=True)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Couldn't open the reference OPUS file: {exc}")
            return

    results = state.classification_results
    best = results[0] if results else None
    if best is None:
        st.error("No lane parsers are registered - nothing to classify against.")
        return

    lane_ids = [r.profile.lane_id for r in results]
    default_lane = state.selected_lane_id or best.profile.lane_id
    selected = st.selectbox(
        "Lane", options=lane_ids, index=lane_ids.index(default_lane) if default_lane in lane_ids else 0,
        help="Auto-selected from the best classification match; override if it's wrong.",
    )
    state.selected_lane_id = selected

    state.rates_mode = st.radio(
        "Generate MRG as:", options=RATES_MODE_OPTIONS,
        index=RATES_MODE_OPTIONS.index(state.rates_mode),
        help="Controls which of the two derived RATES forms gets compared - both come from the same parse.",
    )

    state.apply_known_gaps = st.checkbox(
        "Ignore known non-derivable columns",
        value=state.apply_known_gaps,
        help="Skip flagging columns already documented as never matching a fresh parse "
             "(e.g. `type`, externally-assigned sequence numbers) - see MIGRATION_NOTES.md.",
    )

    if st.button("Run Comparison", type="primary"):
        parser_cls = get_profile(state.selected_lane_id).parser_cls
        parser = parser_cls()
        with st.spinner("Parsing MRG..."):
            state.row_sets = run_parser(parser, state.workbook, MappingProfile())
        state.compare_results = _run_comparison(state.row_sets, state.reference_workbook, state.rates_mode, state.selected_lane_id, state.apply_known_gaps)

    if state.compare_results is not None:
        _render_results(state.compare_results)
