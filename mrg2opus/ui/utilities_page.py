"""Utilities: four tools that start from a filing you already have.

Convert goes raw MRG -> OPUS and Compare checks one against the other.
Everything here takes a finished OPUS workbook in.

The tool is chosen BEFORE a file is asked for, so each one can describe
itself first and then ask for exactly what it needs - which is one file
for three of them and two for the delta.
"""
from __future__ import annotations

import io
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import streamlit as st
from openpyxl.workbook import Workbook

from mrg2opus.excel_io.writer import write_opus_workbook_multi
from mrg2opus.schema.opus_rows import OpusRowSet
from mrg2opus.ui.errors import show_error
from mrg2opus.utilities.checks import run_all
from mrg2opus.utilities.delta import compare_filings
from mrg2opus.utilities.reshape import from_vertical, regroup_port_port, to_port_port, to_vertical
from mrg2opus.utilities.summary import summarize
from mrg2opus.utilities.workbook import LoadedFiling, load_filing

TOOLS = {
    "Reshape a filing": "Move the same rates between OPUS's three shapes.",
    "Check a filing": "What to look at before you submit it.",
    "What's in this filing": "Sheets, groups, ports and the validity window.",
    "What changed since last time": "Two finished filings, side by side.",
}

# (which sheet it reads, which OpusRowSet field it writes, the function
# between them, whether it is exact). Each says plainly which it is.
CONVERSIONS = {
    "RATES → RATES PORT-PORT": ("rates", "rates_port_port", to_port_port, True),
    "RATES → VERTICAL RATES": ("rates", "vertical_rates", to_vertical, True),
    "VERTICAL RATES → RATES": ("vertical_rates", "rates", from_vertical, True),
    "RATES PORT-PORT → RATES (grouped)": ("rates_port_port", "rates", regroup_port_port, False),
}


@dataclass
class UtilitiesState:
    # Parsed filings by upload, so moving between tools doesn't re-read
    # the same 7,400-row workbook.
    cache: dict[str, LoadedFiling] = field(default_factory=dict)


def _get_state() -> UtilitiesState:
    if "utilities_state" not in st.session_state:
        st.session_state["utilities_state"] = UtilitiesState()
    return st.session_state["utilities_state"]


def _open(uploaded) -> Workbook | None:
    try:
        return openpyxl.load_workbook(io.BytesIO(uploaded.getvalue()), data_only=True)
    except Exception as exc:  # noqa: BLE001 - surfaced, not swallowed
        show_error(
            "Couldn't open that file. It may not be a valid .xlsx, or it could be corrupted.", exc
        )
        return None


def _ask_for_filing(label: str, key: str) -> LoadedFiling | None:
    """One uploader plus everything that has to be true before a tool can
    work: the file opens, and it holds sheets we recognize."""
    uploaded = st.file_uploader(label, type=["xlsx"], key=key)
    if uploaded is None:
        return None

    state = _get_state()
    cache_key = f"{uploaded.name}:{uploaded.size}"
    if cache_key not in state.cache:
        wb = _open(uploaded)
        if wb is None:
            return None
        state.cache[cache_key] = load_filing(wb)
    filing = state.cache[cache_key]

    if filing.is_empty:
        st.error(
            "No OPUS sheets recognized in that workbook. Sheets are matched by name - RATES, CMDT NOTE "
            "or SRCHG, VERTICAL RATES, and so on."
        )
        st.caption(f"It holds: {', '.join(filing.unread)}.")
        return None

    read = ", ".join(f"**{name}**" for name in filing.found_as.values())
    st.caption(f"Read {len(filing.found_as)} sheet(s): {read}.")
    if filing.unread:
        st.caption(f"Left alone: {', '.join(filing.unread)}.")
    return filing


def _download(row_set: OpusRowSet, field_name: str, filename: str) -> None:
    """Write one sheet as a real OPUS workbook - same headers, merges and
    column widths the converter's own export uses."""
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "out.xlsx"
        write_opus_workbook_multi({"": row_set}, path)
        st.download_button(
            f"Download {filename}", data=path.read_bytes(), file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", key=f"dl_{field_name}",
        )


# --- the tools ---------------------------------------------------------------

def _reshape() -> None:
    st.caption(
        "OPUS takes the same rates in three shapes. **RATES** is one row per route with four rate "
        "columns across; **RATES PORT-PORT** is the same rows with every \";\"-joined port split out; "
        "**VERTICAL RATES** is one row per container size, with ports running downwards."
    )
    filing = _ask_for_filing("The filing to reshape (.xlsx)", "utilities_reshape_upload")
    if filing is None:
        return

    available = [
        label for label, (source, _f, _fn, _exact) in CONVERSIONS.items() if filing.rows(source)
    ]
    if not available:
        st.info("This workbook has no RATES, RATES PORT-PORT or VERTICAL RATES sheet to reshape.")
        return

    choice = st.radio("Convert", options=available, key="utilities_conversion")
    source, target, convert, exact = CONVERSIONS[choice]
    if exact:
        st.caption("Exact - the same derivation the converter makes on every run.")
    else:
        st.warning(
            "**This groups, it does not restore.** Nothing in a port-port sheet records which ports "
            "shared a row, so every origin agreeing on all the rest - destination, both terms, both "
            "vias, all four rates, both notes - is folded into one. That is the smallest equivalent "
            "filing, not the one you started with: a real LAEC week goes in at 2,528 rows and comes "
            "out at 884, covering the same routes at the same rates. Fewer rows is the safe direction "
            "- OPUS rejects a duplicate filing and accepts a group - but read it before you file it."
        )

    rows = filing.rows(source)
    st.write(f"Reading **{filing.found_as[source]}** — {len(rows):,} rows.")
    if st.button("Reshape", type="primary"):
        out = convert(rows)
        if not out:
            st.warning("That sheet produced no rows. It may be empty, or shaped differently than expected.")
            return
        folded = len(rows) - len(out)
        note = f" — {folded:,} folded into groups" if not exact and folded > 0 else ""
        st.success(f"{len(rows):,} rows in, {len(out):,} rows out{note}.")
        st.dataframe([r.model_dump() for r in out[:200]], hide_index=True, width="stretch")
        if len(out) > 200:
            st.caption(f"Showing the first 200 of {len(out):,}. The download has all of them.")
        _download(OpusRowSet(**{target: out}), target, f"{target}.xlsx")


def _check() -> None:
    st.caption(
        "Each check reports its all-clear as loudly as its findings - \"no duplicate filings\" is "
        "something to confirm before approving, not just the absence of a warning."
    )
    filing = _ask_for_filing("The filing to check (.xlsx)", "utilities_check_upload")
    if filing is None:
        return

    with st.spinner("Checking..."):
        results = run_all(filing)

    st.dataframe(
        [
            {
                "Check": r.name,
                "Result": "OK" if r.passed else "Look at this",
                "Detail": r.headline,
                "Findings": len(r.findings),
            }
            for r in results
        ],
        hide_index=True, width="stretch",
    )
    for result in results:
        if not result.findings:
            continue
        errors = sum(1 for f in result.findings if f.severity == "error")
        label = f"{result.name} — {len(result.findings)} finding(s)"
        with st.expander(f"{label}{f', {errors} that OPUS rejects' if errors else ''}"):
            st.dataframe(
                [{"Severity": f.severity, "What": f.summary, "Where": f.detail} for f in result.findings],
                hide_index=True, width="stretch",
            )


def _summary() -> None:
    st.caption("The questions you ask of an unfamiliar workbook, in the order you ask them.")
    filing = _ask_for_filing("The filing to read (.xlsx)", "utilities_summary_upload")
    if filing is None:
        return

    s = summarize(filing)
    start, end = s.validity
    cols = st.columns(4)
    cols[0].metric("Rate rows", f"{s.rate_rows:,}")
    cols[1].metric("Origins", f"{s.origins:,}")
    cols[2].metric("Destinations", f"{s.destinations:,}")
    cols[3].metric("Commodity groups", f"{len(s.groups):,}")
    if start or end:
        st.caption(f"Validity, from the note sheets' own application dates: **{start or '?'} to {end or '?'}**.")
    else:
        st.caption("No validity window - this filing carries no note sheet with application dates.")

    st.markdown("**Sheets**")
    st.dataframe(s.sheets, hide_index=True, width="stretch")
    if s.groups:
        st.markdown("**Commodity groups**")
        st.dataframe(s.groups, hide_index=True, width="stretch")
    if s.cargo_types:
        st.markdown("**Cargo types**")
        st.dataframe(s.cargo_types, hide_index=True, width="stretch")
    if s.unread_sheets:
        st.caption(f"Sheets this tool doesn't read: {', '.join(s.unread_sheets)}.")


def _delta() -> None:
    st.caption(
        "Routes are matched on their locations, terms, transmodes and vias, with the rates left out of "
        "the match - so a route whose rate moved reports as a change rather than vanishing from one "
        "side and reappearing as new in the other."
    )
    col_before, col_after = st.columns(2)
    with col_before:
        earlier = _ask_for_filing("The EARLIER filing (.xlsx)", "utilities_delta_earlier")
    with col_after:
        later = _ask_for_filing("The LATER filing (.xlsx)", "utilities_delta_later")
    if earlier is None or later is None:
        st.info("Upload both filings to compare them.")
        return

    before, after = earlier.rows("rates"), later.rows("rates")
    if not before or not after:
        st.warning("Both filings need a RATES sheet for this.")
        return

    d = compare_filings(before, after)
    cols = st.columns(4)
    cols[0].metric("Rates changed", f"{len({c.route for c in d.changes}):,}")
    cols[1].metric("Routes added", f"{len(d.added):,}")
    cols[2].metric("Routes removed", f"{len(d.removed):,}")
    cols[3].metric("Unchanged", f"{d.unchanged:,}")

    if d.changes:
        st.markdown("**Rates that moved**, biggest first")
        st.dataframe(
            [
                {
                    "Route": c.route, "Commodity": c.commodity, "Size": c.size,
                    "Was": float(c.was) if c.was is not None else None,
                    "Now": float(c.now) if c.now is not None else None,
                    "Change": float(c.difference) if c.difference is not None else None,
                    "%": round(c.percent, 1) if c.percent is not None else None,
                }
                for c in d.changes
            ],
            hide_index=True, width="stretch",
        )
    else:
        st.success("Every route the two filings share is at the same rate.")

    for label, rows in (("Routes only in the later filing", d.added),
                        ("Routes only in the earlier filing", d.removed)):
        if rows:
            with st.expander(f"{label} — {len(rows):,}"):
                st.dataframe(
                    [
                        {
                            "Origin": r.get("origin_code"), "Destination": r.get("destination_code"),
                            "Prefix": r.get("prefix"), "CGO Type": r.get("cgo_type"),
                            "Commodity": r.get("commodity_group_description"),
                        }
                        for r in rows
                    ],
                    hide_index=True, width="stretch",
                )


RENDERERS = {
    "Reshape a filing": _reshape,
    "Check a filing": _check,
    "What's in this filing": _summary,
    "What changed since last time": _delta,
}


def render() -> None:
    st.subheader("Utilities")
    st.caption(
        "Tools for a filing you already have. Convert builds one from a raw MRG; these reshape, check "
        "and read the finished workbook."
    )

    tool = st.radio("Tool", options=list(TOOLS), key="utilities_tool", captions=list(TOOLS.values()))
    st.divider()
    RENDERERS[tool]()
