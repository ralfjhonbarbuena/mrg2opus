from __future__ import annotations

import io
import tempfile
from pathlib import Path

import streamlit as st

from mrg2opus.excel_io.writer import write_opus_workbook_multi
from mrg2opus.parsers.registry import get_profile
from mrg2opus.ui.errors import show_error
from mrg2opus.ui.sheets import apply_skips as _apply_skips, output_sheets
from mrg2opus.ui.state import WizardState, reset_state


def _build_workbook_bytes(state: WizardState) -> bytes:
    parser_cls = get_profile(state.selected_lane_id).parser_cls if state.selected_lane_id else None
    # parser_cls matters here: skips are keyed by the REAL sheet name, and
    # a lane with sheet-name overrides (the TAD trades) resolves different
    # names - without it the keys wouldn't match and nothing would skip.
    row_sets = _apply_skips(state.row_sets, state.profile.skip_output_sheets, parser_cls)
    overrides = parser_cls.SHEET_NAME_OVERRIDES if parser_cls else None
    scoped_overrides = parser_cls.SCOPED_SHEET_NAME_OVERRIDES if parser_cls else None
    with tempfile.TemporaryDirectory() as tmp_dir:
        out_path = Path(tmp_dir) / "opus_output.xlsx"
        write_opus_workbook_multi(
            row_sets, out_path, sheet_name_overrides=overrides, scoped_sheet_name_overrides=scoped_overrides
        )
        return out_path.read_bytes()


def render(state: WizardState) -> None:
    st.subheader("Export")

    if not state.row_sets:
        st.warning("Nothing to export - go back to Preview.")
        if st.button("← Back to Preview"):
            state.step = 2
            st.rerun()
        return

    st.markdown("#### Sheets in this workbook")
    parser_cls = get_profile(state.selected_lane_id).parser_cls if state.selected_lane_id else None
    skipped = [n for n, skip in state.profile.skip_output_sheets.items() if skip]
    kept = _apply_skips(state.row_sets, state.profile.skip_output_sheets, parser_cls)
    st.dataframe(
        [{"sub-lane": s.scope_label, "sheet": s.name, "rows": s.rows} for s in output_sheets(kept, parser_cls)],
        hide_index=True,
        width="stretch",
    )
    if skipped:
        st.caption(f"Left out at your request: {', '.join(skipped)}")

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
