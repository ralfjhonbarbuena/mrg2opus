"""Shared error-display helper for the wizard and Compare mode: keeps raw
exception text out of the user-facing message - someone using this tool
to file freight rates shouldn't have to parse a Python/openpyxl traceback
to know what went wrong - while still making it available, one click
away, for anyone who does want it (e.g. to report a bug).
"""
from __future__ import annotations

import streamlit as st


def show_error(message: str, exc: Exception) -> None:
    st.error(message)
    with st.expander("Technical details"):
        st.code(str(exc))
