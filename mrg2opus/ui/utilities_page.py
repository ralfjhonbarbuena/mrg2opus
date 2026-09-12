"""Utilities: four tools that start from a filing you already have.

Convert goes raw MRG -> OPUS and Compare checks one against the other.
Everything here takes a finished OPUS workbook in.
"""
from __future__ import annotations

import io
import tempfile
from dataclasses import dataclass
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

TOOLS = [
    "Reshape a filing",
    "Check a filing",
    "What's in this filing",
    "What changed since last time",
]

# (label, which sheet it reads, which OpusRowSet field it writes, the
# function between them). Each says plainly whether it is exact.
CONVERSIONS = {
    "RATES → RATES PORT-PORT": ("rates", "rates_port_port", to_port_port, True),
    "RATES → VERTICAL RATES": ("rates", "vertical_rates", to_vertical, True),
    "VERTICAL RATES → RATES": ("vertical_rates", "rates", from_vertical, True),
    "RATES PORT-PORT → RATES (grouped)": ("rates_port_port", "rates", regroup_port_port, False),
}


@dataclass
class UtilitiesState:
    upload_key: str | None = None
    filing: LoadedFiling | None = None


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


def _load(uploaded, state: UtilitiesState) -> LoadedFiling | None:
    key = f"{uploaded.name}:{uploaded.size}"
    if state.upload_key != key:
        wb = _open(uploaded)
        if wb is None:
            return None
        state.filing, state.upload_key = load_filing(wb), key
    return state.filing


def _found_as_caption(filing: LoadedFiling) -> None:
    read = ", ".join(f"**{name}**" for name in filing.found_as.values())
    st.caption(f"Read {len(filing.found_as)} sheet(s): {read}." if read else "No OPUS sheets recognized.")
    if filing.unread:
        st.caption(f"Left alone: {', '.join(filing.unread)}.")


def _download(row_set: OpusRowSet, field: str, filename: str) -> None:
    """Write one sheet as a real OPUS workbook - same headers, merges and
    column widths the converter's own export uses."""
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "out.xlsx"
        write_opus_workbook_multi({"": row_set}, path)
        st.download_button(
            f"Download {filename}", data=path.read_bytes(), file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", key=f"dl_{field}",
        )


# --- the tools ---------------------------------------------------------------

def _reshape(filing: LoadedFiling) -> None:
    st.markdown("#### Reshape a filing")
    st.caption(
        "OPUS takes the same rates in three shapes. **RATES** is one row per route with four rate "
        "columns across; **RATES PORT-PORT** is the same rows with every \";\"-joined port split out; "
        "**VERTICAL RATES** is one row per container size, with ports running downwards."
    )
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
        note = f" &mdash; {folded:,} folded into groups" if not exact and folded > 0 else ""
        st.success(f"{len(rows):,} rows in, {len(out):,} rows out{note}.")
        st.dataframe(
            [r.model_dump() for r in out[:200]], hide_index=True, width="stretch"
        )
        if len(out) > 200:
            st.caption(f"Showing the first 200 of {len(out):,}. The download has all of them.")
        _download(OpusRowSet(**{target: out}), target, f"{target}.xlsx")


def _check(filing: LoadedFiling) -> None:
    st.markdown("#### Check a filing")
    st.caption(
        "What to look at before submitting. Each check reports its all-clear as loudly as its findings - "
        "\"no duplicate filings\" is something to confirm, not just the absence of a warning."
    )
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


def _summary(filing: LoadedFiling) -> None:
    st.markdown("#### What's in this filing")
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


def _delta(filing: LoadedFiling) -> None:
    st.markdown("#### What changed since last time")
    st.caption(
        "Two finished filings of the same lane. Routes are matched on their locations, terms, transmodes "
        "and vias, with the rates left out of the match - so a route whose rate moved reports as a change "
        "rather than vanishing from one side and reappearing as new in the other."
    )
    earlier_file = st.file_uploader(
        "The EARLIER filing (.xlsx)", type=["xlsx"], key="utilities_delta_earlier"
    )
    if earlier_file is None:
        st.info("The file above is treated as the later one. Upload the earlier filing to compare against.")
        return

    wb = _open(earlier_file)
    if wb is None:
        return
    earlier = load_filing(wb)
    before, after = earlier.rows("rates"), filing.rows("rates")
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


def render() -> None:
    state = _get_state()
    st.subheader("Utilities")
    st.caption(
        "Tools for a filing you already have. Convert builds one from a raw MRG; these reshape, check "
        "and read the finished workbook."
    )

    uploaded = st.file_uploader("OPUS-format Excel file (.xlsx)", type=["xlsx"], key="utilities_upload")
    if uploaded is None:
        st.info("Upload an OPUS filing to begin.")
        return

    filing = _load(uploaded, state)
    if filing is None:
        return
    if filing.is_empty:
        st.error(
            "No OPUS sheets recognized in that workbook. Sheets are matched by name - RATES, CMDT NOTE "
            "or SRCHG, VERTICAL RATES, and so on."
        )
        st.caption(f"It holds: {', '.join(filing.unread)}.")
        return
    _found_as_caption(filing)

    st.divider()
    tool = st.radio("Tool", options=TOOLS, horizontal=True, key="utilities_tool")
    if tool == TOOLS[0]:
        _reshape(filing)
    elif tool == TOOLS[1]:
        _check(filing)
    elif tool == TOOLS[2]:
        _summary(filing)
    else:
        _delta(filing)
