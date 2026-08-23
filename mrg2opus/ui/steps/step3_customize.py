from __future__ import annotations

import streamlit as st

from mrg2opus.parsers.registry import get_profile
from mrg2opus.presets.store import list_presets, load_preset, save_preset
from mrg2opus.ui.parsing import run_parser
from mrg2opus.ui.state import WizardState

SHEET_LABELS = {
    "rates": "OPUS RATES",
    "rates_port_port": "OPUS RATES PORT-PORT",
    "arbs": "OPUS ARBS",
    "cmdt_notes": "OPUS CMDT NOTE",
    "special_notes": "OPUS SPECIAL NOTE",
}


def _output_sheet_names(row_sets: dict) -> list[str]:
    names: list[str] = []
    for suffix, row_set in row_sets.items():
        tag = f"-{suffix}" if suffix else ""
        if row_set.rates:
            names.append(f"OPUS RATES{tag}")
        if row_set.rates_port_port:
            names.append(f"OPUS RATES{tag} PORT-PORT")
        if row_set.arbs:
            names.append(f"OPUS ARBS{tag}")
        if row_set.cmdt_notes:
            names.append(f"OPUS CMDT NOTE{tag}")
        if row_set.special_notes:
            names.append(f"OPUS SPECIAL NOTE{tag}")
    return names


def render(state: WizardState) -> None:
    st.subheader("Customize")

    if not state.row_sets:
        st.warning("Nothing parsed yet - go back to Preview.")
        if st.button("← Back to Preview"):
            state.step = 2
            st.rerun()
        return

    with st.expander("Load / save a named preset"):
        existing = list_presets()
        col_load, col_save = st.columns(2)
        with col_load:
            if existing:
                pick = st.selectbox("Existing presets", options=existing)
                if st.button("Load preset"):
                    state.profile = load_preset(pick)
                    st.success(f"Loaded preset '{pick}'.")
                    st.rerun()
            else:
                st.caption("No saved presets yet.")
        with col_save:
            name = st.text_input("Save current settings as", value=state.profile.name)
            if st.button("Save preset"):
                path = save_preset(state.profile.model_copy(update={"name": name}))
                st.success(f"Saved to {path.name}.")

    groups = state.default_commodity_groups
    st.markdown("#### Commodity groups")
    st.caption(
        "Codes shown here are only starting suggestions from the parsed file - there is no commodity code "
        "registry, so the code that ends up in the OPUS output should mainly come from you. Descriptions ship "
        "with a sensible default but are always yours to change too. This table doesn't support mouse dragging "
        "to reorder rows, but the **Order** column does the same job: give groups numbers to set the sequence "
        "they'll appear in on the generated OPUS RATES / RATES PORT-PORT / CMDT NOTE sheets - lower numbers "
        "come first. Two (or more) groups given the exact same Description merge into one CMDT NOTE block. "
        "**Skip DG** stops that group's base Dry rows from also filing an identical D/DG duplicate - check it "
        "for a group that shouldn't carry a Dangerous Goods rate at all."
    )
    if groups:
        existing_order = state.profile.commodity_group_order
        editor_rows = [
            {
                "order": (existing_order.index(desc) + 1) if desc in existing_order else len(existing_order) + i + 1,
                "default_code": code,
                "default_description": desc,
                # All three overrides are keyed by the group's DEFAULT
                # description (desc), not its code - see
                # parsers/common/commodity.py's module docstring. Several
                # groups can share one default code (e.g. LAWC's main dry
                # grid/"Reefer"/"LAWC NOR" all default to G0001), so code
                # can't serve as a unique key here.
                "code": state.profile.commodity_code_overrides.get(desc, code),
                "description": state.profile.commodity_description_overrides.get(desc, desc),
                "override_cmdt_seq": state.profile.commodity_sequence_overrides.get(desc),
                "skip_dg": state.profile.skip_dg_generation.get(desc, False),
            }
            for i, (code, desc) in enumerate(groups)
        ]
        editor_rows.sort(key=lambda r: r["order"])
        edited = st.data_editor(
            editor_rows,
            hide_index=True,
            width="stretch",
            disabled=["default_code", "default_description"],
            column_order=["order", "default_code", "default_description", "code", "description", "override_cmdt_seq", "skip_dg"],
            column_config={
                "order": st.column_config.NumberColumn("Order", step=1, required=True, help="Lower numbers appear first in the output."),
                "default_code": st.column_config.TextColumn("Parsed default code"),
                "default_description": st.column_config.TextColumn("Parsed default description"),
                "code": st.column_config.TextColumn("Code (yours)", required=True),
                "description": st.column_config.TextColumn("Description (yours)", required=True),
                "override_cmdt_seq": st.column_config.NumberColumn("Override CMDT Seq", step=1),
                "skip_dg": st.column_config.CheckboxColumn("Skip DG", help="Don't file a D/DG duplicate for this group's base Dry rows."),
            },
            key="commodity_overrides_editor",
        )
    else:
        edited = []
        st.caption("No commodity groups found in the parsed output.")

    st.markdown("#### Skip output sheets")
    sheet_names = _output_sheet_names(state.row_sets)
    skip_choices: dict[str, bool] = {}
    if sheet_names:
        cols = st.columns(min(3, len(sheet_names)))
        for i, name in enumerate(sheet_names):
            with cols[i % len(cols)]:
                skip_choices[name] = st.checkbox(
                    name, value=state.profile.skip_output_sheets.get(name, False), key=f"skip_{name}"
                )
    else:
        st.caption("No output sheets to skip.")

    col_back, col_next = st.columns(2)
    with col_back:
        if st.button("← Back to Preview"):
            state.step = 2
            st.rerun()
    with col_next:
        if st.button("Apply & Continue to Export →", type="primary"):
            # Keyed by each row's default_description - the stable identity
            # every override dict uses (see the editor_rows comment above).
            code_overrides = {
                r["default_description"]: r["code"].strip()
                for r in edited
                if r.get("code") and r["code"].strip() and r["code"].strip() != r["default_code"]
            }
            description_overrides = {
                r["default_description"]: r["description"].strip()
                for r in edited
                if r.get("description") and r["description"].strip() and r["description"].strip() != r["default_description"]
            }
            sequence_overrides = {
                r["default_description"]: int(r["override_cmdt_seq"])
                for r in edited
                if r.get("override_cmdt_seq") not in (None, "")
            }
            skip_output_sheets = {name: skip for name, skip in skip_choices.items() if skip}
            skip_dg_generation = {
                r["default_description"]: True for r in edited if r.get("skip_dg")
            }
            # The FINAL description (post-override) of each row, sorted by
            # its "Order" value - this is what actually ends up on the
            # output rows, so it's what group_order needs to match against
            # (see parsers/common/ordering.py::reorder_row_set()).
            commodity_group_order = [
                (r["description"].strip() if r.get("description") else r["default_description"])
                for r in sorted(edited, key=lambda r: r.get("order", 0))
            ]

            state.profile = state.profile.model_copy(
                update={
                    "commodity_code_overrides": code_overrides,
                    "commodity_description_overrides": description_overrides,
                    "commodity_sequence_overrides": sequence_overrides,
                    "commodity_group_order": commodity_group_order,
                    "skip_output_sheets": skip_output_sheets,
                    "skip_dg_generation": skip_dg_generation,
                }
            )

            parser_cls = get_profile(state.selected_lane_id).parser_cls
            parser = parser_cls()
            with st.spinner("Re-running with overrides..."):
                state.row_sets = run_parser(parser, state.workbook, state.profile)
            state.output_bytes = None  # invalidate any previously-built export
            state.step = 4
            st.rerun()
