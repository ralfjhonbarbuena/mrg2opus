"""Phase 2 Streamlit wizard entrypoint.

    ./.venv/Scripts/python.exe -m streamlit run mrg2opus/ui/app.py

4-step linear wizard: upload+classify -> preview -> customize -> export.
Each step is a module in mrg2opus/ui/steps/ with a render(state) function.
"""
from __future__ import annotations

import streamlit as st

# Importing the lane modules registers their LayoutProfile as a side effect -
# same requirement as cli.py.
from mrg2opus.parsers import cse, eaf, laec, lawc, saf  # noqa: F401
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

    state = get_state()

    cols = st.columns(len(STEP_LABELS))
    for col, (step_num, label) in zip(cols, STEP_LABELS.items()):
        with col:
            if step_num == state.step:
                st.markdown(f"**➤ {label}**")
            elif step_num < state.step:
                st.markdown(f"✅ {label}")
            else:
                st.markdown(f"{label}")
    st.divider()

    STEP_RENDERERS[state.step](state)


if __name__ == "__main__":
    main()
