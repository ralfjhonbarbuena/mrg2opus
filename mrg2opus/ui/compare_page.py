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
import pandas as pd
import streamlit as st
from openpyxl.workbook import Workbook

from mrg2opus.audit.compare import (
    AUDIT_KEY_FIELDS,
    CMDT_NOTE_IGNORE_FIELDS_BY_LANE,
    NOTE_PRESENTATION_FIELDS,
    RATES_IGNORE_FIELDS_BY_LANE,
    RATES_PORT_PORT_IGNORE_FIELDS_BY_LANE,
    RATES_PRESENTATION_FIELDS,
    SPECIAL_NOTE_IGNORE_FIELDS_BY_LANE,
    arbs_row_key,
    audit_concat,
    audit_row_key,
    diff_by_key,
    diff_cmdt_blocks,
    diff_vertical_blocks,
    explain_profile_overrides,
    find_duplicate_filings,
    find_sheet,
    freetime_compared_fields,
    freetime_row_key,
    profile_skips_rows,
    profile_without_row_skips,
    read_arbs_sheet,
    read_cmdt_note_sheet,
    read_freetime_sheet,
    read_rates_sheet,
    read_route_note_sheet,
    read_special_note_sheet,
    read_vertical_rates_sheet,
    reconstruct_blocks,
    reconstruct_vertical_blocks,
    route_note_counts,
    split_mismatches_by_tier,
)
from mrg2opus.audit.side_by_side import build_side_by_side_workbook
from mrg2opus.excel_io.merge import DuplicateSheetError
from mrg2opus.excel_io.writer import resolve_sheet_names
from mrg2opus.parsers.registry import ClassificationResult, get_profile
from mrg2opus.presets.models import MappingProfile
from mrg2opus.schema import opus_columns as cols
from mrg2opus.ui.errors import show_error
from mrg2opus.ui.mrg_upload import fingerprint_uploads, load_and_classify
from mrg2opus.ui.parsing import run_parser
from mrg2opus.ui.state import get_state as get_wizard_state

RATES_MODE_OPTIONS = ["Grouped (RATES)", "Exploded (RATES PORT-PORT)", "Both"]

# Sheets that can be left out of a comparison. The two RATES forms have
# their own control above (they are the point of the screen); these are
# the ones an audit may not need every time - and RN and VERTICAL RATES
# are the slowest, since one is counted per note and the other rebuilt
# block by block.
SKIPPABLE_SHEETS = ["ORIGIN ARBS", "CMDT NOTE", "SPECIAL NOTE", "ROUTE NOTE", "VERTICAL RATES", "FREETIME"]


@dataclass
class CompareState:
    upload_key: str | None = None
    workbook: Workbook | None = None
    classification_results: list[ClassificationResult] = field(default_factory=list)
    selected_lane_id: str | None = None
    reference_workbook: Workbook | None = None
    rates_mode: str = "Both"
    apply_known_gaps: bool = True
    # Whether to parse with the settings configured in Convert, or with a
    # default profile. Compare used to be hardwired to defaults, so it
    # silently answered "what does the tool produce unaided?" rather than
    # "is the file I am about to upload right?".
    use_wizard_profile: bool = True
    skip_sheets: list[str] = field(default_factory=list)
    row_sets: dict[str, Any] | None = None
    compare_results: list[dict[str, Any]] | None = None
    # Set when the comparison ran against a customized profile: the output
    # columns that will differ BECAUSE of a setting, and why.
    explained_overrides: dict[str, str] = field(default_factory=dict)
    # Rows filed twice - OPUS rejects them, so the audit confirms absence.
    duplicate_filings: list[dict[str, Any]] = field(default_factory=list)
    # (sheet label, our rows, the reference's rows) for the paired workbook.
    side_by_side: list[tuple] = field(default_factory=list)
    # (lane_id, rates_mode, apply_known_gaps, use_wizard_profile) at the moment compare_results
    # was computed - lets render() detect when the visible results no
    # longer match the current control settings, instead of silently
    # showing a stale comparison after the user changes lane/mode/toggle
    # without re-clicking Run Comparison.
    results_computed_for: tuple | None = None


def _get_state() -> CompareState:
    if "compare" not in st.session_state:
        st.session_state.compare = CompareState()
    return st.session_state.compare


def _compare_keyed_sheet(sheet_type, suffix, sheet_name, generated, ref_wb, key_fn, fields, reader,
                         ignore_fields=frozenset(), presentation_fields=frozenset(), skipped_keys=frozenset()) -> dict | None:
    sub_lane = suffix or "(default)"
    try:
        expected = reader(ref_wb, sheet_name)
    except KeyError:
        if not generated:
            return None
        return {
            "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
            "found_in_reference": False, "matched": 0,
            "routes_matched": 0, "substance_ok": 0,
            "missing": [], "intentionally_absent": [],
            "extra": sorted(({"key": key_fn(r)} for r in generated), key=lambda d: str(d["key"])),
            "field_mismatches": [], "substance_mismatches": [], "presentation_mismatches": [],
        }
    if not generated and not expected:
        return None
    result = diff_by_key(generated, expected, key_fn=key_fn, fields=fields, ignore_fields=ignore_fields)
    mismatches = sorted(
        (
            {"key": m[0], "field": m[1], "generated": m[2], "reference": m[3]}
            for m in result.field_mismatches
        ),
        key=lambda d: (str(d["key"]), d["field"]),
    )
    substance, presentation = split_mismatches_by_tier(mismatches, presentation_fields)

    # `matched` counts rows with NO difference at all, so renaming one
    # commodity code drops it through the floor even though every rate is
    # still right. These two say what people actually came to find out:
    # how many routes line up, and how many of those are correct on
    # substance (a row differing only on presentation still counts).
    common_keys = {key_fn(r) for r in generated} & {key_fn(r) for r in expected}
    routes_matched = len(common_keys)
    substance_ok = routes_matched - len({m["key"] for m in substance})

    # Rows the user chose not to file read as "missing" otherwise, which is
    # exactly how a parser that failed to produce them would read.
    intentionally_absent = [
        {"key": k} for k in sorted(result.missing, key=str) if k in skipped_keys
    ]
    unexplained_missing = [
        {"key": k} for k in sorted(result.missing, key=str) if k not in skipped_keys
    ]
    return {
        "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
        "found_in_reference": True, "matched": result.matched,
        "routes_matched": routes_matched, "substance_ok": substance_ok,
        "missing": unexplained_missing,
        "intentionally_absent": intentionally_absent,
        "extra": sorted(({"key": k} for k in result.extra), key=lambda d: str(d["key"])),
        "field_mismatches": mismatches,
        "substance_mismatches": substance,
        "presentation_mismatches": presentation,
    }


def _compare_block_sheet(sheet_type, suffix, sheet_name, generated, ref_wb, fields, reader,
                         ignore_fields=frozenset(), presentation_fields=frozenset()) -> dict | None:
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
            "routes_matched": None, "substance_ok": None,
            "missing": [], "intentionally_absent": [],
            "extra": [{"contents": k} for k in extra_keys],
            "field_mismatches": [], "substance_mismatches": [], "presentation_mismatches": [],
        }
    gen_has_blocks = bool(reconstruct_blocks(generated))
    exp_has_blocks = bool(reconstruct_blocks(expected))
    if not gen_has_blocks and not exp_has_blocks:
        return None
    result = diff_cmdt_blocks(generated, expected, fields, ignore_fields=ignore_fields)
    mismatches = sorted(
        (
            {"key": m[0], "child_index": m[1], "field": m[2], "generated": m[3], "reference": m[4]}
            for m in result.field_mismatches
        ),
        key=lambda d: (d["key"], d["child_index"], d["field"]),
    )
    substance, presentation = split_mismatches_by_tier(mismatches, presentation_fields)
    return {
        "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
        "found_in_reference": True, "matched": None,
        "routes_matched": None, "substance_ok": None,
        "missing": [{"contents": k} for k in result.missing_blocks],
        "intentionally_absent": [],
        "extra": [{"contents": k} for k in result.extra_blocks],
        "field_mismatches": mismatches,
        "substance_mismatches": substance,
        "presentation_mismatches": presentation,
    }


def _skipped_row_keys(parser, workbook, profile) -> frozenset[tuple]:
    """The row keys the user's own Skip Filing / Skip DG settings removed.

    Worked out by re-parsing with just those two settings off and taking
    the difference, rather than by guessing which reference rows belong to
    a skipped group - the reference files a group under its own name, which
    need not be the name the profile keys on. Costs a second parse, so it
    only runs when a skip is actually set.
    """
    if not profile_skips_rows(profile):
        return frozenset()

    def keys_of(row_sets: dict) -> set[tuple]:
        keys: set[tuple] = set()
        for row_set in row_sets.values():
            keys |= {audit_row_key(r.model_dump()) for r in row_set.rates}
            keys |= {audit_row_key(r.model_dump()) for r in row_set.rates_port_port}
        return keys

    kept = keys_of(run_parser(parser, workbook, profile))
    everything = keys_of(run_parser(parser, workbook, profile_without_row_skips(profile)))
    return frozenset(everything - kept)



def _empty_result(sheet_type, sub_lane, sheet_name, **over):
    base = {
        "sheet_type": sheet_type, "sub_lane": sub_lane, "sheet_name": sheet_name,
        "found_in_reference": True, "matched": None, "routes_matched": None, "substance_ok": None,
        "missing": [], "intentionally_absent": [], "extra": [],
        "field_mismatches": [], "substance_mismatches": [], "presentation_mismatches": [],
    }
    base.update(over)
    return base


def _compare_route_note_sheet(suffix, sheet_name, generated, ref_wb) -> dict | None:
    """RN can't be matched row-for-row: it is addressed by (Header Seq,
    Route Seq) and OPUS assigns both itself - LAWC's real filing numbers
    its headers from 1015, which no fresh parse reproduces. What compares
    is how many times each note text appears and which lane it names."""
    sub_lane = suffix or "(default)"
    try:
        expected = read_route_note_sheet(ref_wb, sheet_name)
    except KeyError:
        if not generated:
            return None
        return _empty_result("ROUTE NOTE", sub_lane, sheet_name, found_in_reference=False,
                             extra=[{"contents": c, "lane": lane} for c, lane in route_note_counts(
                                 [r.model_dump() for r in generated])])
    if not generated and not expected:
        return None

    ours = route_note_counts([r.model_dump() for r in generated])
    theirs = route_note_counts(expected)
    missing = [{"contents": c, "lane": lane, "reference count": n}
               for (c, lane), n in sorted(theirs.items()) if (c, lane) not in ours]
    extra = [{"contents": c, "lane": lane, "your count": n}
             for (c, lane), n in sorted(ours.items()) if (c, lane) not in theirs]
    differing = [
        {"key": c, "field": f"lane {lane}" if lane else "count",
         "generated": ours[(c, lane)], "reference": theirs[(c, lane)]}
        for (c, lane) in sorted(set(ours) & set(theirs))
        if ours[(c, lane)] != theirs[(c, lane)]
    ]
    return _empty_result(
        "ROUTE NOTE", sub_lane, sheet_name,
        matched=len(set(ours) & set(theirs)) - len(differing),
        missing=missing, extra=extra,
        field_mismatches=differing, substance_mismatches=differing,
    )


def _compare_vertical_rates_sheet(suffix, sheet_name, generated, ref_wb) -> dict | None:
    """Compared as BLOCKS, not rows: the sheet is a columnar encoding, so
    a row on its own ("VNCMP", no destination, no rate) is the second
    origin of the block above rather than a route of its own."""
    sub_lane = suffix or "(default)"
    rows = [r.model_dump() for r in generated]
    try:
        expected = read_vertical_rates_sheet(ref_wb, sheet_name)
    except KeyError:
        if not rows:
            return None
        return _empty_result(
            "VERTICAL RATES", sub_lane, sheet_name, found_in_reference=False,
            extra=[{"Row": str(b.key)} for b in reconstruct_vertical_blocks(rows)],
        )
    if not rows and not expected:
        return None

    missing, extra, differing = diff_vertical_blocks(rows, expected)
    mismatches = [
        {"key": key, "field": "rates", "generated": str(ours), "reference": str(theirs)}
        for key, ours, theirs in differing
    ]
    ours_blocks = reconstruct_vertical_blocks(rows)
    matched_blocks = len(ours_blocks) - len(extra) - len(differing)
    return _empty_result(
        "VERTICAL RATES", sub_lane, sheet_name,
        matched=matched_blocks,
        missing=[{"Row": str(k)} for k in missing],
        extra=[{"Row": str(k)} for k in extra],
        field_mismatches=mismatches, substance_mismatches=mismatches,
    )


def _compare_freetime_sheet(suffix, sheet_name, generated, ref_wb) -> dict | None:
    sub_lane = suffix or "(default)"
    rows = [r.model_dump() for r in generated]
    try:
        expected = read_freetime_sheet(ref_wb, sheet_name)
    except KeyError:
        if not rows:
            return None
        return _empty_result("FREETIME", sub_lane, sheet_name, found_in_reference=False,
                             extra=[{"Row": str(freetime_row_key(r))} for r in rows])
    if not rows and not expected:
        return None

    result = diff_by_key(rows, expected, key_fn=freetime_row_key, fields=freetime_compared_fields())
    mismatches = [
        {"key": m[0], "field": m[1], "generated": m[2], "reference": m[3]}
        for m in result.field_mismatches
    ]
    return _empty_result(
        "FREETIME", sub_lane, sheet_name, matched=result.matched,
        missing=[{"Row": str(k)} for k in sorted(result.missing, key=str)],
        extra=[{"Row": str(k)} for k in sorted(result.extra, key=str)],
        field_mismatches=mismatches, substance_mismatches=mismatches,
    )


def _read_or_empty(ref_wb: Workbook, sheet_name: str) -> list[dict]:
    """The reference may simply not carry a sheet (LAWC files no VERTICAL
    RATES); an empty side is still worth writing, so the pairing is
    visible rather than the sheet silently absent."""
    try:
        return read_rates_sheet(ref_wb, sheet_name)
    except KeyError:
        return []


def tag_for(suffix: str) -> str:
    return f"-{suffix}" if suffix else ""


def _pick_sheet_name(ref_wb: Workbook, scoped: str, plain: str) -> str:
    """Which name to look for in the reference.

    A multi-scope lane writes RATES-OEW and RATES-OMW into ONE workbook,
    but the reference filings are delivered one file per scope, with
    plain names - so the scoped name finds nothing and every row reports
    as extra. Prefer the scoped name when the reference really is a
    combined workbook, and fall back to the plain one when it isn't.
    Returns the scoped name when neither exists, so the "sheet not found"
    message names what was actually looked for.
    """
    for candidate in (scoped, plain):
        try:
            find_sheet(ref_wb, candidate)
            return candidate
        except KeyError:
            continue
    return scoped


def _run_comparison(
    row_sets: dict,
    ref_wb: Workbook,
    rates_mode: str,
    lane_id: str,
    apply_known_gaps: bool = True,
    skipped_keys: frozenset = frozenset(),
    skip_sheets: frozenset = frozenset(),
    sheet_name_overrides: dict | None = None,
    scoped_sheet_name_overrides: dict | None = None,
) -> tuple[list[dict], list[dict]]:
    want_grouped = rates_mode in ("Grouped (RATES)", "Both")
    want_exploded = rates_mode in ("Exploded (RATES PORT-PORT)", "Both")

    rates_ignore = RATES_IGNORE_FIELDS_BY_LANE.get(lane_id, frozenset()) if apply_known_gaps else frozenset()
    port_port_ignore = RATES_PORT_PORT_IGNORE_FIELDS_BY_LANE.get(lane_id, frozenset()) if apply_known_gaps else frozenset()
    cmdt_ignore = CMDT_NOTE_IGNORE_FIELDS_BY_LANE.get(lane_id, frozenset()) if apply_known_gaps else frozenset()
    special_ignore = SPECIAL_NOTE_IGNORE_FIELDS_BY_LANE.get(lane_id, frozenset()) if apply_known_gaps else frozenset()

    results: list[dict] = []
    duplicates: list[dict] = []
    side_by_side: list[tuple] = []
    for suffix, row_set in row_sets.items():
        # Part of the audit in its own right: OPUS rejects a filing that
        # carries the same route twice, so it has to be confirmed absent
        # rather than merely not mentioned.
        for label, rows in (("RATES", row_set.rates), ("RATES PORT-PORT", row_set.rates_port_port)):
            for concat, count in find_duplicate_filings([r.model_dump() for r in rows]):
                duplicates.append({
                    "sheet": f"{label}{tag_for(suffix)}", "count": count, "row": concat,
                })
        scoped = resolve_sheet_names(suffix, sheet_name_overrides, scoped_sheet_name_overrides)
        plain = resolve_sheet_names("", sheet_name_overrides, scoped_sheet_name_overrides)
        names = {k: _pick_sheet_name(ref_wb, scoped[k], plain[k]) for k in scoped}
        if want_grouped:
            side_by_side.append((
                names["rates"],
                [x.model_dump() for x in row_set.rates],
                _read_or_empty(ref_wb, names["rates"]),
            ))
            r = _compare_keyed_sheet(
                "RATES", suffix, names["rates"],
                [x.model_dump() for x in row_set.rates], ref_wb,
                audit_row_key, cols.RATES_ROW_FIELDS, read_rates_sheet, ignore_fields=rates_ignore,
                presentation_fields=RATES_PRESENTATION_FIELDS, skipped_keys=skipped_keys,
            )
            if r is not None:
                results.append(r)
        if want_exploded:
            side_by_side.append((
                names["rates_port_port"],
                [x.model_dump() for x in row_set.rates_port_port],
                _read_or_empty(ref_wb, names["rates_port_port"]),
            ))
            r = _compare_keyed_sheet(
                "RATES PORT-PORT", suffix, names["rates_port_port"],
                [x.model_dump() for x in row_set.rates_port_port], ref_wb,
                audit_row_key, cols.RATES_ROW_FIELDS, read_rates_sheet, ignore_fields=port_port_ignore,
                presentation_fields=RATES_PRESENTATION_FIELDS, skipped_keys=skipped_keys,
            )
            if r is not None:
                results.append(r)
        r = None if "ORIGIN ARBS" in skip_sheets else _compare_keyed_sheet(
            "ARBS", suffix, names["arbs"],
            [x.model_dump() for x in row_set.arbs], ref_wb,
            arbs_row_key, cols.ARBS_ROW_FIELDS, read_arbs_sheet,
        )
        if r is not None:
            results.append(r)
        r = None if "CMDT NOTE" in skip_sheets else _compare_block_sheet(
            "CMDT NOTE", suffix, names["cmdt_notes"],
            [x.model_dump() for x in row_set.cmdt_notes], ref_wb,
            cols.CMDT_NOTE_ROW_FIELDS, read_cmdt_note_sheet, ignore_fields=cmdt_ignore,
            presentation_fields=NOTE_PRESENTATION_FIELDS,
        )
        if r is not None:
            results.append(r)
        r = None if "SPECIAL NOTE" in skip_sheets else _compare_block_sheet(
            "SPECIAL NOTE", suffix, names["special_notes"],
            [x.model_dump() for x in row_set.special_notes], ref_wb,
            cols.SPECIAL_NOTE_ROW_FIELDS, read_special_note_sheet, ignore_fields=special_ignore,
            presentation_fields=NOTE_PRESENTATION_FIELDS,
        )
        if r is not None:
            results.append(r)
        if "ROUTE NOTE" not in skip_sheets:
            r = _compare_route_note_sheet(suffix, names["route_notes"], row_set.route_notes, ref_wb)
            if r is not None:
                results.append(r)
        if "VERTICAL RATES" not in skip_sheets:
            r = _compare_vertical_rates_sheet(suffix, names["vertical_rates"], row_set.vertical_rates, ref_wb)
            if r is not None:
                results.append(r)
        if "FREETIME" not in skip_sheets:
            r = _compare_freetime_sheet(suffix, names["freetime"], row_set.freetime, ref_wb)
            if r is not None:
                results.append(r)
    return results, duplicates, side_by_side


_DETAIL_ROW_LIMIT = 50


_KEY_COLUMN_LABELS = {
    "origin_code": "Origin", "origin_term": "O.Term", "origin_transmode": "O.Transmode",
    "o_via_code": "O.Via", "d_via_code": "D.Via",
    "destination_code": "Destination", "destination_term": "D.Term",
    "destination_transmode": "D.Transmode",
    "prefix": "Prefix", "cgo_type": "CGO Type", "route_note": "Route Note",
}


def _as_text(value):
    """None stays None; everything else becomes a string - see _flatten."""
    return None if value is None else str(value)


def _flatten(entries: list[dict]) -> list[dict]:
    """Spread the match key into its own named columns.

    The CSV used to carry the key as one stringified Python tuple -
    "('MYPKG', 'MXZLO', 'DG', 'D', None, None)" - which can't be sorted,
    filtered or VLOOKUP'd against the other draft. These are the same
    columns the audit concatenates by hand, so a row lines up with the
    auditor's own sheet.
    """
    flat: list[dict] = []
    for entry in entries:
        row: dict = {}
        key = entry.get("key")
        if isinstance(key, tuple) and len(key) == len(AUDIT_KEY_FIELDS):
            row.update({
                _KEY_COLUMN_LABELS[name]: value
                for name, value in zip(AUDIT_KEY_FIELDS, key)
            })
        elif isinstance(key, str):
            row["Note contents"] = key
        elif key is not None:
            row["Row"] = str(key)
        if "contents" in entry:
            row["Note contents"] = entry["contents"]
        if "child_index" in entry:
            row["Row in block"] = "parent" if entry["child_index"] == 0 else str(entry["child_index"])
        if "field" in entry:
            # Stringified deliberately. One column carries values from
            # every OPUS column at once - "G0001" from a commodity code
            # sits beside 5800 from a rate - and Arrow, which Streamlit
            # serialises the grid through, can't type a mixed column: it
            # raised ArrowInvalid and dumped a traceback per render.
            # None is left as None so a blank still reads as blank.
            row["Column"] = entry["field"]
            row["Your draft"] = _as_text(entry.get("generated"))
            row["Reference"] = _as_text(entry.get("reference"))
        flat.append(row)
    return flat


def _render_detail_table(label: str, rows: list[dict], key: str) -> None:
    """One bucket of differences, on screen and as a CSV.

    The on-screen grid is capped at _DETAIL_ROW_LIMIT rows (a real
    mismatch count would otherwise make Streamlit's grid unwieldy); the
    CSV always holds every row and is always offered, since it is the
    thing the auditor actually works from beside their own draft."""
    flat = _flatten(rows)
    st.markdown(f"**{label}** ({len(flat)})")
    st.dataframe(flat[:_DETAIL_ROW_LIMIT], hide_index=True, width="stretch")
    # The CSV is always offered, not only past the on-screen cap: this is
    # the artefact the auditor works from, alongside their own draft.
    st.download_button(
        f"⬇ Download as CSV ({len(flat)} rows)",
        data=pd.DataFrame(flat).to_csv(index=False).encode("utf-8"),
        file_name=f"{key}.csv",
        mime="text/csv",
        key=f"download_{key}",
    )
    if len(flat) > _DETAIL_ROW_LIMIT:
        st.caption(f"Showing {_DETAIL_ROW_LIMIT} of {len(flat)} on screen - the CSV has all of them.")


def _render_results(results: list[dict], explained_overrides: dict[str, str],
                    duplicate_filings: list[dict] | None = None,
                    side_by_side: list[tuple] | None = None) -> None:
    if not results:
        st.info("Nothing to compare - the parsed MRG produced no rows for the sheet type(s) selected.")
        return

    # The headline answers the question people actually bring here: are the
    # routes and the rates right? A renamed commodity code produces one
    # difference per row, so left in the same column as a wrong rate it
    # buries it - the two are counted separately.
    substance = sum(len(r["substance_mismatches"]) for r in results)
    presentation = sum(len(r["presentation_mismatches"]) for r in results)
    unexplained_missing = sum(len(r["missing"]) for r in results)
    extra = sum(len(r["extra"]) for r in results)
    intentional = sum(len(r["intentionally_absent"]) for r in results)

    routes = sum(r["routes_matched"] or 0 for r in results)
    substance_ok = sum(r["substance_ok"] or 0 for r in results)

    cols_ = st.columns(5)
    cols_[0].metric("Routes matched", routes, help="Rows present in both, matched on origin, destination, cargo type and routing.")
    cols_[1].metric(
        "…correct on substance", substance_ok,
        delta=None if substance_ok == routes else -(routes - substance_ok),
        help="Of the matched routes, how many agree on every rate, currency, term and port name. "
             "A row differing only on a field you chose still counts here.",
    )
    cols_[2].metric("Rows unaccounted for", unexplained_missing + extra, help="Missing from your output or extra in it, with no setting explaining either.")
    cols_[3].metric("Presentation differences", presentation, help="Commodity code/description, sequence numbers, note text - what you choose or the writer assigns.")
    cols_[4].metric("Not filed on purpose", intentional, help="Rows your Skip Filing / Skip DG settings removed.")

    if substance == 0 and unexplained_missing == 0 and extra == 0:
        st.success(
            "Every route matched and every rate agrees. "
            + (f"The {presentation} remaining differences are all in fields you control or the writer assigns."
               if presentation else "No differences at all.")
        )

    if side_by_side:
        st.download_button(
            "⬇ Download both drafts side by side (.xlsx)",
            data=build_side_by_side_workbook(side_by_side),
            file_name="mrg2opus_side_by_side.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help=(
                "Your draft and the reference in one workbook, a sheet each, with the "
                "CONCATENATE already built in column A and a live XLOOKUP against the other "
                "sheet in column B - to reconcile them by hand the way the audit already does."
            ),
        )

    if duplicate_filings:
        st.error(
            f"**{len(duplicate_filings)} route(s) filed more than once.** OPUS rejects a duplicate "
            "filing, so these have to be resolved before the filing is approved."
        )
        _render_detail_table("Duplicate filings", duplicate_filings, "duplicate_filings")
    else:
        st.caption("✓ No duplicate route filings — checked on the same columns the audit concatenates.")

    if explained_overrides:
        st.info(
            "**Differences expected from your settings** — these columns will not match the reference, "
            "because you changed them:\n\n"
            + "\n".join(f"- `{field_name}` — {why}" for field_name, why in sorted(explained_overrides.items()))
        )

    st.markdown("#### Comparison summary")
    st.dataframe(
        [
            {
                "Sheet": r["sheet_type"],
                "Sub-lane": r["sub_lane"],
                "In reference?": "Yes" if r["found_in_reference"] else "No - sheet not found",
                # None, not "-": the note sheets have no row-level matched
                # count, and mixing a placeholder string into a column of
                # integers is what Arrow refuses to type - the same fault
                # the detail grids had. A blank cell reads the same and
                # keeps the column sortable as a number.
                "Matched": r["matched"],
                "Missing": len(r["missing"]),
                "Not filed on purpose": len(r["intentionally_absent"]),
                "Extra": len(r["extra"]),
                "Substance": len(r["substance_mismatches"]),
                "Presentation": len(r["presentation_mismatches"]),
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
            row_key = f"{r['sheet_type']}_{r['sub_lane']}".replace(" ", "_")
            # Substance first: it is the only bucket that means something
            # is wrong, so it should not sit below hundreds of expected
            # presentation rows.
            if r["substance_mismatches"]:
                _render_detail_table(
                    "Substance differences — rates, currencies, terms, port names",
                    r["substance_mismatches"], f"{row_key}_substance",
                )
            if r["missing"]:
                _render_detail_table("Missing (in reference but not generated)", r["missing"], f"{row_key}_missing")
            if r["extra"]:
                _render_detail_table("Extra (generated but not in reference)", r["extra"], f"{row_key}_extra")
            if r["intentionally_absent"]:
                _render_detail_table(
                    "Not filed on purpose (your Skip Filing / Skip DG settings)",
                    r["intentionally_absent"], f"{row_key}_intentional",
                )
            if r["presentation_mismatches"]:
                _render_detail_table(
                    "Presentation differences — fields you choose, or the writer assigns",
                    r["presentation_mismatches"], f"{row_key}_presentation",
                )
            if (
                r["found_in_reference"]
                and not r["missing"] and not r["extra"] and not r["field_mismatches"]
            ):
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
            state.workbook, state.classification_results = load_and_classify(mrg_payloads, mrg_names)
        except DuplicateSheetError as exc:
            st.error(str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - surfaced directly to the user, not swallowed
            show_error(
                "Couldn't open one of the MRG files. It may not be a valid .xlsx file, "
                "or the file could be corrupted.",
                exc,
            )
            return

    if state.reference_workbook is None:
        try:
            state.reference_workbook = openpyxl.load_workbook(io.BytesIO(reference_payload), data_only=True)
        except Exception as exc:  # noqa: BLE001
            show_error(
                "Couldn't open the reference OPUS file. It may not be a valid .xlsx file, "
                "or the file could be corrupted.",
                exc,
            )
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

    # Which settings the MRG is parsed with. Compare used to be hardwired
    # to a default profile, so it answered "what does the tool produce
    # unaided?" rather than "is the file I am about to upload right?" -
    # and anything customized in Convert was invisible to it.
    wizard_profile = get_wizard_state().profile
    customized = bool(explain_profile_overrides(wizard_profile)) or profile_skips_rows(wizard_profile)
    state.use_wizard_profile = st.checkbox(
        "Use the settings from Convert",
        value=state.use_wizard_profile,
        help=(
            "Compare the output you actually configured — your commodity codes, descriptions, "
            "sequence numbers and any groups you chose not to file. Uncheck to compare the tool's "
            "unaided output against the reference instead."
        ),
    )
    if state.use_wizard_profile and not customized:
        st.caption("Nothing is customized in Convert yet, so this is the same as the tool's default output.")

    state.skip_sheets = st.multiselect(
        "Skip these sheets",
        options=SKIPPABLE_SHEETS,
        default=state.skip_sheets,
        help="Leave a sheet out of the comparison. The two RATES forms are chosen above; "
             "ROUTE NOTE and VERTICAL RATES are the slowest to check.",
    )

    profile = wizard_profile if state.use_wizard_profile else MappingProfile()
    current_settings = (
        state.selected_lane_id, state.rates_mode, state.apply_known_gaps,
        state.use_wizard_profile, tuple(sorted(state.skip_sheets)),
    )

    if st.button("Run Comparison", type="primary"):
        parser_cls = get_profile(state.selected_lane_id).parser_cls
        parser = parser_cls()
        with st.spinner("Parsing MRG..."):
            state.row_sets = run_parser(parser, state.workbook, profile)
            skipped_keys = _skipped_row_keys(parser, state.workbook, profile)
        state.explained_overrides = explain_profile_overrides(profile)
        state.compare_results, state.duplicate_filings, state.side_by_side = _run_comparison(
            state.row_sets, state.reference_workbook, state.rates_mode,
            state.selected_lane_id, state.apply_known_gaps, skipped_keys,
            frozenset(state.skip_sheets),
            parser_cls.SHEET_NAME_OVERRIDES, parser_cls.SCOPED_SHEET_NAME_OVERRIDES,
        )
        state.results_computed_for = current_settings

    if state.compare_results is not None:
        if state.results_computed_for != current_settings:
            st.warning(
                "⚠️ Lane, RATES mode, the known-gaps toggle, the profile choice or the skipped sheets "
                "changed since this comparison ran - click **Run Comparison** to refresh before trusting these results."
            )
        _render_results(state.compare_results, state.explained_overrides,
                        state.duplicate_filings, state.side_by_side)
