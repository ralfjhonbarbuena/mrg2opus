"""Phase 2 Streamlit wizard entrypoint.

    ./.venv/Scripts/python.exe -m streamlit run mrg2opus/ui/app.py

Two modes, selected at the top: "Convert" (the 4-step wizard:
upload+classify -> preview -> customize -> export) and "Compare"
(standalone: upload an MRG plus a reference OPUS file, see where they
diverge - see docs/superpowers/specs/2026-08-23-mrg-opus-comparison-design.md).
"""
from __future__ import annotations

import streamlit as st

# Importing the lane modules registers their LayoutProfile as a side effect -
# same requirement as cli.py.
from mrg2opus.parsers import auec, auwc, cse, eaf, laec, lawc, saf, tad_oew_omw, waf  # noqa: F401
from mrg2opus.ui import compare_page
from mrg2opus.ui.state import get_state
from mrg2opus.ui.steps import step1_upload, step2_preview, step3_customize, step4_export

STEP_LABELS = {
    1: "1. Upload & Classify",
    2: "2. Preview",
    3: "3. Customize",
    4: "4. Export",
}

STEP_RENDERERS = {
    1: step1_upload.render,
    2: step2_preview.render,
    3: step3_customize.render,
    4: step4_export.render,
}


def main() -> None:
    st.set_page_config(page_title="mrg2opus", layout="wide")
    st.title("MRG → OPUS Converter")

    mode = st.radio("Mode", options=["Convert", "Compare"], horizontal=True)
    st.divider()

    if mode == "Compare":
        compare_page.render()
        return

    state = get_state()

    cols = st.columns(len(STEP_LABELS))
    for col, (step_num, label) in zip(cols, STEP_LABELS.items()):
        with col:
            if step_num == state.step:
                st.markdown(f"**➤ {label}**")
            else:
                prefix = "✅ " if step_num < state.step else ""
                # Every step's own render() already checks its prerequisites
                # (state.workbook/row_sets) and shows a "go back" prompt if
                # they're missing, so jumping directly to any step - not
                # just completed ones - degrades gracefully rather than
                # erroring.
                if st.button(f"{prefix}{label}", key=f"step_nav_{step_num}", width="stretch"):
                    state.step = step_num
                    st.rerun()
    st.divider()

    STEP_RENDERERS[state.step](state)


if __name__ == "__main__":
    main()
