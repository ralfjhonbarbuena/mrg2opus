from __future__ import annotations

import io
import tempfile
from pathlib import Path

import streamlit as st

from mrg2opus.excel_io.writer import write_opus_workbook_multi
from mrg2opus.parsers.registry import get_profile
from mrg2opus.schema.opus_rows import OpusRowSet
from mrg2opus.ui.errors import show_error
from mrg2opus.ui.state import WizardState, reset_state


def _apply_skips(row_sets: dict[str, OpusRowSet], skip_output_sheets: dict[str, bool]) -> dict[str, OpusRowSet]:
    """skip_output_sheets is keyed by the final OPUS sheet name (as shown in
    Step 3's checkboxes) - translate that back into which OpusRowSet field
    to blank out per sub-lane before writing."""
    if not skip_output_sheets:
        return row_sets
    filtered: dict[str, OpusRowSet] = {}
    for suffix, row_set in row_sets.items():
        tag = f"-{suffix}" if suffix else ""
        data = row_set.model_dump()
        if skip_output_sheets.get(f"OPUS RATES{tag}"):
            data["rates"] = []
        if skip_output_sheets.get(f"OPUS RATES{tag} PORT-PORT"):
            data["rates_port_port"] = []
        if skip_output_sheets.get(f"OPUS ARBS{tag}"):
            data["arbs"] = []
        if skip_output_sheets.get(f"OPUS CMDT NOTE{tag}"):
            data["cmdt_notes"] = []
        if skip_output_sheets.get(f"OPUS SPECIAL NOTE{tag}"):
            data["special_notes"] = []
        filtered[suffix] = OpusRowSet.model_validate(data)
    return filtered


def _build_workbook_bytes(state: WizardState) -> bytes:
    row_sets = _apply_skips(state.row_sets, state.profile.skip_output_sheets)
    overrides = get_profile(state.selected_lane_id).parser_cls.SHEET_NAME_OVERRIDES if state.selected_lane_id else None
    with tempfile.TemporaryDirectory() as tmp_dir:
        out_path = Path(tmp_dir) / "opus_output.xlsx"
        write_opus_workbook_multi(row_sets, out_path, sheet_name_overrides=overrides)
        return out_path.read_bytes()


def render(state: WizardState) -> None:
    st.subheader("Export")

    if not state.row_sets:
        st.warning("Nothing to export - go back to Preview.")
        if st.button("← Back to Preview"):
            state.step = 2
            st.rerun()
        return

    st.markdown("#### Final row counts")
    summary = []
    for suffix, row_set in state.row_sets.items():
        label = suffix or "(default)"
        summary.append(
            {
                "sub-lane": label,
                "RATES": len(row_set.rates),
                "RATES PORT-PORT": len(row_set.rates_port_port),
                "ARBS": len(row_set.arbs),
                "CMDT NOTE": len(row_set.cmdt_notes),
                "SPECIAL NOTE": len(row_set.special_notes),
            }
        )
    st.dataframe(summary, hide_index=True, width="stretch")

    if state.profile.skip_output_sheets:
        skipped = [name for name, skip in state.profile.skip_output_sheets.items() if skip]
        if skipped:
            st.caption(f"Skipping from export: {', '.join(skipped)}")

    if state.output_bytes is None:
        with st.spinner("Building output workbook..."):
            try:
                state.output_bytes = _build_workbook_bytes(state)
            except Exception as exc:  # noqa: BLE001 - a real write failure, surface it directly
                show_error("Couldn't build the output workbook.", exc)
                return

    if len(state.upload_names) == 1:
        base_name = Path(state.upload_names[0]).stem
    elif state.upload_names:
        base_name = f"{Path(state.upload_names[0]).stem}_combined"
    else:
        base_name = "opus_output"
    out_filename = f"{base_name}_opus.xlsx"
    state.output_filename = out_filename

    st.download_button(
        "⬇ Download OPUS workbook",
        data=io.BytesIO(state.output_bytes),
        file_name=out_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    col_back, col_restart = st.columns(2)
    with col_back:
        if st.button("← Back to Customize"):
            state.step = 3
            st.rerun()
    with col_restart:
        if st.button("Start over with a new file"):
            reset_state()
            st.rerun()
