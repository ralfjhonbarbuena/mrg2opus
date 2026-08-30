from __future__ import annotations

from decimal import Decimal

import streamlit as st

from mrg2opus.parsers.registry import get_profile
from mrg2opus.presets.store import list_presets, load_preset, save_preset
from mrg2opus.ui.parsing import run_parser
from mrg2opus.ui.sheets import output_sheets
from mrg2opus.ui.state import WizardState

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
        "**Skip DG** stops just that group's base Dry rows from also filing an identical D/DG duplicate - to "
        "turn Dangerous Goods off for the whole filing instead, use the Dangerous Goods checkbox below."
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

    st.markdown("#### Special instructions")
    excluded_charge_codes_input = st.text_input(
        "Exclude charge codes from filing (comma-separated)",
        value=", ".join(state.profile.excluded_charge_codes),
        help=(
            "Every charge code the raw MRG's own \"Includes\" line names is filed by default - this is where "
            "you drop the ones your account shouldn't file. Applies to the whole filing, every commodity "
            "group. Use it for rules the MRG text doesn't reflect - e.g. excluding \"BRS\", or a Hong Kong "
            "account excluding \"BAF\" because it duplicates OBS and isn't applicable for their RFAs."
        ),
    )

    col_rfa_eff, col_rfa_exp = st.columns(2)
    with col_rfa_eff:
        rfa_effective_date = st.date_input(
            "RFA effective date (optional)",
            value=state.profile.rfa_effective_date,
            help=(
                "Each individual charge code's own CMDT NOTE Application Effective date normally just mirrors "
                "this filing's rate validity start - but the real-world RFA (Rate Filing Agreement) window is "
                "usually a separate, longer-lived date a human filer enters instead. Leave blank to keep using "
                "the rate validity start."
            ),
        )
    with col_rfa_exp:
        rfa_expiry_date = st.date_input(
            "RFA expiry date (optional)",
            value=state.profile.rfa_expiry_date,
            help="Same idea as RFA effective date, for Application Expires. Leave blank to keep using the rate validity end.",
        )

    include_vertical_rates = st.checkbox(
        "Include Vertical Rates (alternate OPUS upload format)",
        value=state.profile.include_vertical_rates,
        help=(
            "The same rates as the RATES sheet, reshaped one row per container size instead of 4 rate "
            "columns per row - a general OPUS upload option, faster to upload. Generated for every lane by "
            "default, as one sheet. Uncheck to leave it out entirely."
        ),
    )

    # One DG control for EVERY lane. The underlying field and the default
    # differ by lane family - TAD mirrors its own export tool's opt-IN
    # "Include Dry Dangerous" setting (off by default), every other lane
    # generates the duplicate by default and opts OUT - but the user sees
    # the same checkbox either way, and non-TAD lanes keep the per-group
    # "Skip DG" column above for finer control.
    st.markdown("#### Dangerous Goods (DG)")
    is_tad_lane = bool(state.selected_lane_id and state.selected_lane_id.startswith("TAD-"))
    group_descriptions = [desc for _code, desc in groups]
    if is_tad_lane:
        dg_currently_on = state.profile.generate_tad_dg_duplicate
    else:
        # On unless every group is currently skipped.
        dg_currently_on = not (
            bool(group_descriptions)
            and all(state.profile.skip_dg_generation.get(d, False) for d in group_descriptions)
        )
    generate_dg = st.checkbox(
        "File D/DG duplicate rows",
        value=dg_currently_on,
        help=(
            "Files a second, identical row for each base Dry (D/DR) row with CGO TYPE flipped to DG, at the "
            "same rate - a standing filing convention, not something the raw MRG states. "
            + (
                "Off by default for TAD filings, matching the team's own export tool."
                if is_tad_lane
                else "On by default for this lane. Unchecking it drops DG everywhere; to drop it for only "
                "some commodity groups, leave this checked and use the Skip DG column above."
            )
        ),
    )

    generate_tad_dg_duplicate = generate_dg if is_tad_lane else state.profile.generate_tad_dg_duplicate
    include_tad_d7 = state.profile.include_tad_d7
    tad_d7_addon = state.profile.tad_d7_addon
    if is_tad_lane:
        is_aew_amw = state.selected_lane_id == "TAD-AEW-AMW"
        if is_aew_amw:
            st.markdown("#### TAD AEW/AMW filing options")
            col_d7, col_d7_amt = st.columns(2)
            with col_d7:
                include_tad_d7 = st.checkbox(
                    "Include D7 (OFT 45)",
                    value=include_tad_d7,
                    help=(
                        "The raw MRG carries no OFT 45 column for AEW/AMW - this derives it as OFT 40HC plus "
                        "the add-on beside. Applies to D/DR rows only; a generated D/DG duplicate copies the "
                        "same value. AEW/AMW only (the Japan scope never gets one)."
                    ),
                )
            with col_d7_amt:
                tad_d7_addon = Decimal(
                    str(
                        st.number_input(
                            "D7 add-on (added to OFT 40HC)",
                            value=float(tad_d7_addon),
                            step=50.0,
                            help="Per this filing's own Surcharges reference sheet the standard add-on is 700.",
                        )
                    )
                )

    st.markdown("#### Skip output sheets")
    st.caption("Named exactly as they'll appear in the exported workbook. Every sheet the export would contain is listed.")
    sheets = output_sheets(state.row_sets, get_profile(state.selected_lane_id).parser_cls)
    skip_choices: dict[str, bool] = {}
    if sheets:
        cols = st.columns(min(3, len(sheets)))
        for i, sheet in enumerate(sheets):
            with cols[i % len(cols)]:
                skip_choices[sheet.name] = st.checkbox(
                    f"{sheet.name}  ({sheet.rows:,})",
                    value=state.profile.skip_output_sheets.get(sheet.name, False),
                    key=f"skip_{sheet.scope}_{sheet.name}",
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
            # Master DG toggle wins when it's off (skip every group);
            # when on, the per-group Skip DG column decides. TAD lanes
            # don't use this dict at all - their toggle is the bool below.
            if not generate_dg and not is_tad_lane:
                skip_dg_generation = {r["default_description"]: True for r in edited}
            else:
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

            excluded_charge_codes = [
                c.strip().upper() for c in excluded_charge_codes_input.split(",") if c.strip()
            ]

            state.profile = state.profile.model_copy(
                update={
                    "commodity_code_overrides": code_overrides,
                    "commodity_description_overrides": description_overrides,
                    "commodity_sequence_overrides": sequence_overrides,
                    "commodity_group_order": commodity_group_order,
                    "skip_output_sheets": skip_output_sheets,
                    "skip_dg_generation": skip_dg_generation,
                    "excluded_charge_codes": excluded_charge_codes,
                    "rfa_effective_date": rfa_effective_date,
                    "rfa_expiry_date": rfa_expiry_date,
                    "include_vertical_rates": include_vertical_rates,
                    "generate_tad_dg_duplicate": generate_tad_dg_duplicate,
                    "include_tad_d7": include_tad_d7,
                    "tad_d7_addon": tad_d7_addon,
                }
            )

            parser_cls = get_profile(state.selected_lane_id).parser_cls
            parser = parser_cls()
            with st.spinner("Re-running with overrides..."):
                state.row_sets = run_parser(parser, state.workbook, state.profile)
            state.output_bytes = None  # invalidate any previously-built export
            state.step = 4
            st.rerun()
