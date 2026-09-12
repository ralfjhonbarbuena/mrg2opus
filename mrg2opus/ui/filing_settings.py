"""The filing settings, rendered identically wherever they are needed.

Convert uses them to shape the workbook it exports. Compare uses them to
shape the draft it checks against a reference - the auditor writes their
own draft rather than inheriting the processor's, so they need the same
controls, not a pointer at somebody else's.

One implementation on purpose: two copies of a commodity-group editor
would drift, and a setting that behaved differently on the two screens
would be worse than not having it on one of them.
"""
from __future__ import annotations

from decimal import Decimal

import streamlit as st

from mrg2opus.presets.models import MappingProfile, ScopeOverrides
from mrg2opus.presets.store import export_profile, import_profile, preset_filename

# The scope picker's "not one scope, the shared settings" option.
_ALL_SCOPES = "All scopes"

# Where a just-imported preset waits for the next run to pick it up, and
# where a failed import leaves its reason. Both named without the
# "<prefix>_" shape on purpose: reset_filing_settings() wipes every key
# with that prefix, and these have to survive it - importing a preset is
# exactly when the widgets need clearing.
_PENDING_PRESET = "pending_preset_profile"
_IMPORT_ERROR = "preset_import_error"
# A parsed preset that belongs to another lane, held back for the
# "Import anyway" button rather than thrown away.
_IMPORT_MISMATCH = "preset_import_mismatch"

_EDITOR_KEY_BASE = "commodity_overrides_editor"
_EDITOR_NONCE = "commodity_overrides_editor_nonce"


def _editor_key(prefix: str, scope: str = "") -> str:
    """The grid's widget key, carrying a nonce that Refresh bumps.

    Changing the KEY is what actually resets st.data_editor. Popping
    st.session_state[key] - the obvious move, and what this did at first -
    looks right and does nothing: the grid's edits live in frontend state
    tied to the widget's identity, so it just repopulates the key and the
    edits stay on screen. A key it has never seen is a brand-new widget,
    drawn from the DataFrame we pass rather than from any retained state.

    The prefix keeps Convert's grid and Compare's apart, and the scope
    keeps one sub-lane's grid from showing the previous one's edits: two
    widgets sharing a key would share their edits.
    """
    return f"{prefix}_{scope}_{_EDITOR_KEY_BASE}_{st.session_state.get(prefix + _EDITOR_NONCE, 0)}"


def _refresh_editor(prefix: str) -> None:
    """Throw away the grid's in-progress edits so it redraws from the last
    applied values - see _editor_key for why the key changes rather than
    the state being cleared."""
    st.session_state[prefix + _EDITOR_NONCE] = st.session_state.get(prefix + _EDITOR_NONCE, 0) + 1


def reset_filing_settings(prefix: str) -> None:
    """Forget every setting made under `prefix`, for when the thing being
    configured changes out from under them.

    Streamlit keeps a keyed widget's value in session_state, so it
    outlives the data it was chosen against: switch lane and the grid
    would replay the previous lane's edits onto the new lane's groups by
    ROW POSITION, quietly renaming unrelated commodity groups. Dropping
    the keys makes each widget draw from the values passed to it again.
    """
    for key in [k for k in st.session_state if k.startswith(prefix + "_")]:
        del st.session_state[key]
    _refresh_editor(prefix)


def _stage_preset(prefix: str, profile: MappingProfile) -> None:
    """Put a preset where the next run will pick it up, and clear the
    widgets so it shows.

    Both halves are needed. Replacing the profile alone changes nothing
    on screen: Streamlit renders a keyed widget from its stored value, not
    from the `value=` a caller passes, so the settings would keep showing
    whatever was there while the returned profile said otherwise - the
    same way the Refresh button did nothing until it changed the grid's
    key rather than its contents.
    """
    st.session_state[prefix + _PENDING_PRESET] = profile
    st.session_state.pop(prefix + _IMPORT_ERROR, None)
    st.session_state.pop(prefix + _IMPORT_MISMATCH, None)
    # Clears the uploader along with everything else, so the file doesn't
    # sit there afterwards looking like it still needs importing.
    reset_filing_settings(prefix)


def _import_anyway(prefix: str) -> None:
    """Take the other lane's preset that _import_preset held back."""
    held = st.session_state.get(prefix + _IMPORT_MISMATCH)
    if held is not None:
        _stage_preset(prefix, held)


def _import_preset(prefix: str, upload_key: str, lane_id: str | None = None) -> None:
    """Read an uploaded preset, check it belongs here, and stage it.

    Both halves are needed. Replacing the profile alone changes nothing
    on screen: Streamlit renders a keyed widget from its stored value, not
    from the `value=` a caller passes, so the settings would keep showing
    whatever was there while the returned profile said otherwise - the
    same way the Refresh button did nothing until it changed the grid's
    key rather than its contents.

    Runs as a button's callback rather than on the upload itself, so
    dropping a file never silently overwrites settings someone is partway
    through entering.

    A preset from another lane is REFUSED rather than applied, because
    the wrong lane is not the harmless case it looks like. Every commodity
    setting is keyed by a group's default description, and those repeat
    across lanes - "FAK" is a group in EAF and in all three TAD lanes - so
    the other lane's overrides land on a group that merely shares a name,
    silently renaming and recoding it. It is held for "Import anyway"
    rather than discarded: some of what a preset carries (the RFA dates,
    the excluded charge codes) is lane-agnostic and worth reusing on
    purpose.
    """
    uploaded = st.session_state.get(upload_key)
    if uploaded is None:
        st.session_state[prefix + _IMPORT_ERROR] = "Choose a file first."
        return
    try:
        profile = import_profile(uploaded.getvalue())
    except Exception as exc:  # noqa: BLE001 - the reason is shown to the user
        st.session_state[prefix + _IMPORT_ERROR] = (
            f"{uploaded.name} isn't a settings file this can read. ({type(exc).__name__})"
        )
        st.session_state.pop(prefix + _IMPORT_MISMATCH, None)
        return

    # A file with no lane predates the stamp, or was built before a lane
    # was picked. Nothing to check it against, so it goes through.
    if profile.lane_id and lane_id and profile.lane_id != lane_id:
        st.session_state[prefix + _IMPORT_MISMATCH] = profile
        st.session_state[prefix + _IMPORT_ERROR] = (
            f"These settings were made for **{profile.lane_id}**, and this filing is **{lane_id}**. "
            "Commodity groups are matched by description, and the same description turns up in more "
            "than one lane - so importing this would rename and recode whichever group happens to "
            "share a name."
        )
        return
    _stage_preset(prefix, profile)


def _render_presets(profile: MappingProfile, key_prefix: str, lane_id: str | None = None) -> None:
    """Export the whole settings sheet as a file, or import one back.

    A file rather than a named folder entry, so the settings travel with
    the person who made them - onto another machine, to the auditor, into
    the repo - instead of living beside whichever copy of the app wrote
    them. The format is unchanged, so anything already in data/presets
    imports here as-is.

    Takes the profile the settings CURRENTLY describe, not the one last
    applied - which is why the caller renders this into a container held
    open from the top of the page and filled at the bottom. Exporting what
    was applied rather than what is on screen would quietly write a
    different sheet than the one being looked at.

    The file is stamped with the lane it was made for, and importing one
    from another lane is refused - see _import_preset.
    """
    with st.expander("Export / import these settings"):
        col_export, col_import = st.columns(2)
        with col_export:
            st.markdown("**Export**")
            name = st.text_input(
                "Name this settings file", value=profile.name, key=f"{key_prefix}_preset_name"
            )
            st.download_button(
                "Export settings",
                data=export_profile(profile.model_copy(update={"name": name, "lane_id": lane_id})),
                file_name=preset_filename(name),
                mime="application/json",
                key=f"{key_prefix}_preset_export",
                help="Everything on this page, as one file you can keep, send on, or import later.",
            )
            if lane_id:
                st.caption(f"Stamped **{lane_id}** — it will refuse to import into another lane.")
        with col_import:
            st.markdown("**Import**")
            upload_key = f"{key_prefix}_preset_import"
            st.file_uploader("A settings file (.json)", type=["json"], key=upload_key)
            st.button(
                "Import settings", key=f"{key_prefix}_preset_import_go",
                on_click=_import_preset, args=(key_prefix, upload_key, lane_id),
                help="Replaces every setting below with the file's own.",
            )
            if (problem := st.session_state.get(key_prefix + _IMPORT_ERROR)):
                st.error(problem)
            # Only offered once a mismatch has been refused, so the guard
            # can't be walked past without reading what it said.
            if st.session_state.get(key_prefix + _IMPORT_MISMATCH) is not None:
                st.button(
                    "Import anyway", key=f"{key_prefix}_preset_import_force",
                    on_click=_import_anyway, args=(key_prefix,),
                    help="Use it regardless - worth it for the dates and charge codes, which are "
                         "the same whatever the lane. Check the commodity table afterwards.",
                )


def _with_scope_overrides(
    profile: MappingProfile, scope: str, settings: dict
) -> dict[str, ScopeOverrides]:
    """`profile.by_scope` with `scope` set to just the settings that
    differ from the filing-wide ones, and dropped entirely when none do."""
    kept: dict = {}
    for field_name, value in settings.items():
        shared = getattr(profile, field_name)
        if isinstance(value, list):
            if value != shared:
                kept[field_name] = value
        else:
            differing = {k: v for k, v in value.items() if shared.get(k) != v}
            # A group the shared settings answer but this scope's grid
            # does not is an override too - it says "not set here".
            if differing:
                kept[field_name] = differing
    by_scope = {k: v for k, v in profile.by_scope.items() if k != scope}
    if kept:
        by_scope[scope] = ScopeOverrides(**kept)
    return by_scope


def render_filing_settings(
    profile: MappingProfile,
    groups: list[tuple[str, str]],
    groups_by_scope: dict[str, list[tuple[str, str]]],
    dg_twin_groups: frozenset[str],
    reefer_nor_groups: frozenset[str],
    scopes: list[str],
    lane_id: str | None,
    key_prefix: str,
) -> MappingProfile:
    """Draw the settings and return the profile they describe.

    `groups` is the parser's own (code, description) pairs from an
    override-free parse - the identities every override dict is keyed by.
    `dg_twin_groups` names the subset of those the lane files a DG twin
    for by default (commodity_utils.groups_offering_dg_twins), and
    `reefer_nor_groups` the reefer and NOR groups, which can be asked for
    one whether or not the lane files it (parsers.common.dg_twins).
    `scopes` is the parse's own sub-lane keys and `groups_by_scope` which
    groups each of them has. A lane with sub-lanes files each as its own
    OPUS workbook, so the commodity settings are asked per scope as well
    as filing-wide - see the scope picker below and
    MappingProfile.for_scope.

    Returns a NEW profile built from what is on screen rather than
    mutating the one passed in, so the caller decides when it takes
    effect: Convert holds it until "Apply & Continue", Compare adopts it
    straight away and re-checks on "Run Comparison".
    """
    # A preset imported on the previous run replaces the caller's profile
    # for this one, and is handed back so the caller adopts it too.
    pending = st.session_state.pop(key_prefix + _PENDING_PRESET, None)
    if pending is not None:
        profile = pending
    # Held open here and filled at the very bottom, so the preset panel
    # sits above the settings while saving what they currently say.
    preset_slot = st.container()

    st.markdown("#### Commodity groups")
    # Which scope these settings are for. "All scopes" edits the
    # filing-wide values; picking one edits only what that scope answers
    # differently, so a shared setting still reaches every scope that
    # hasn't overridden that particular group.
    editing_scope = None
    if len(scopes) > 1:
        choice = st.selectbox(
            "Settings for",
            options=[_ALL_SCOPES, *scopes],
            help=(
                "Each sub-lane is filed as its own OPUS workbook and can answer these differently - one "
                "scope's own commodity code, description, order or skips. Anything you leave alone here "
                "follows the All scopes value."
            ),
            key=f"{key_prefix}_scope_choice",
        )
        editing_scope = None if choice == _ALL_SCOPES else choice
    # What the grid shows: the filing-wide settings, or one scope's own
    # view of them (its overrides merged over the shared ones).
    view = profile if editing_scope is None else profile.for_scope(editing_scope)
    if editing_scope is not None:
        groups = groups_by_scope.get(editing_scope, [])
        st.caption(f"Editing **{editing_scope}** only. Blank cells here fall back to the All scopes value.")
    st.caption(
        "Every column here is yours to change. **CMDT Code** is numbered G0001, G0002, ... in the order the "
        "groups were found - a placeholder, never a code read out of your file, since there is no commodity "
        "code registry; edit it to whatever your account files under. **CMDT Description** starts as the "
        "description the tool derived, and two groups given the exact same description merge into one CMDT "
        "NOTE block. This table doesn't support mouse dragging to reorder rows, but **Order** does the same "
        "job: lower numbers appear first on the generated OPUS RATES / RATES PORT-PORT / CMDT NOTE sheets. "
        "Leave **CMDT Seq** blank to let the tool number the group itself. **Skip Filing** leaves the group "
        "out of the filing entirely - to keep a group but drop only its DG rows, see Dangerous Goods below."
    )
    if groups:
        existing_order = view.commodity_group_order
        # Every override dict is keyed by the group's DEFAULT description
        # (desc) - see parsers/common/commodity.py's module docstring. The
        # parser's own structural code can't serve as that key (several
        # groups share one, e.g. LAWC's main dry grid / "Reefer" / "LAWC
        # NOR" are all internally G0001) and isn't shown at all: the code
        # the user sees is the sequential G0001, G0002, ... placeholder
        # seeded into commodity_code_overrides right after the first parse
        # by commodity_utils.assign_sequential_default_codes().
        editor_rows = [
            {
                "order": (existing_order.index(desc) + 1) if desc in existing_order else len(existing_order) + i + 1,
                "code": view.commodity_code_overrides.get(desc, code),
                "description": view.commodity_description_overrides.get(desc, desc),
                "override_cmdt_seq": view.commodity_sequence_overrides.get(desc),
                "skip_filing": view.skip_commodity_filing.get(desc, False),
            }
            for i, (code, desc) in enumerate(groups)
        ]
        # The default description is the key every override is stored
        # under, but showing it beside the editable one is what made the
        # table confusing ("which of these two do I change?"), so it's
        # kept out of the editor entirely and recovered positionally when
        # the edits come back - see the zip() below. Sorting both lists
        # the same way is what makes that safe; data_editor never adds,
        # deletes or moves rows (num_rows is "fixed" by default).
        row_keys = [desc for _code, desc in groups]
        paired = sorted(zip(editor_rows, row_keys), key=lambda pair: pair[0]["order"])
        editor_rows = [row for row, _key in paired]
        row_keys = [key for _row, key in paired]
        st.button(
            "↺ Refresh table",
            key=f"{key_prefix}_refresh_commodity_editor",
            on_click=_refresh_editor,
            args=(key_prefix,),
            help=(
                "Redraw the table from the last applied values - use it if you've cleared a cell and want "
                "to see its default again. Discards any edits made since you last hit Apply."
            ),
        )
        edited = st.data_editor(
            editor_rows,
            hide_index=True,
            width="stretch",
            column_order=["order", "override_cmdt_seq", "code", "description", "skip_filing"],
            column_config={
                "order": st.column_config.NumberColumn(
                    "Order", step=1, required=True,
                    help="Lower numbers appear first in the output.",
                ),
                "override_cmdt_seq": st.column_config.NumberColumn(
                    "CMDT Seq", step=1,
                    help="Leave blank to let the tool number this group automatically.",
                ),
                "code": st.column_config.TextColumn(
                    "CMDT Code", required=True,
                    help="The code that lands in the OPUS output. A G0001-onwards placeholder by default.",
                ),
                "description": st.column_config.TextColumn(
                    "CMDT Description", required=True,
                    help="Two groups given the same description merge into one CMDT NOTE block.",
                ),
                "skip_filing": st.column_config.CheckboxColumn(
                    "Skip Filing",
                    help="Leave this commodity group out of the filing altogether.",
                ),
            },
            key=_editor_key(key_prefix, editing_scope or ""),
        )
    else:
        edited, row_keys = [], []
        st.caption("No commodity groups found in the parsed output.")

    st.markdown("#### Special instructions")
    excluded_charge_codes_input = st.text_input(
        "Exclude charge codes from filing (comma-separated)",
        value=", ".join(profile.excluded_charge_codes),
        help=(
            "Every charge code the raw MRG's own \"Includes\" line names is filed by default - this is where "
            "you drop the ones your account shouldn't file. Applies to the whole filing, every commodity "
            "group. Use it for rules the MRG text doesn't reflect - e.g. excluding \"BRS\", or a Hong Kong "
            "account excluding \"BAF\" because it duplicates OBS and isn't applicable for their RFAs."
        ),
        key=f"{key_prefix}_excluded_charge_codes",
    )

    col_rfa_eff, col_rfa_exp = st.columns(2)
    with col_rfa_eff:
        rfa_effective_date = st.date_input(
            "RFA effective date (optional)",
            value=profile.rfa_effective_date,
            help=(
                "Each individual charge code's own CMDT NOTE Application Effective date normally just mirrors "
                "this filing's rate validity start - but the real-world RFA (Rate Filing Agreement) window is "
                "usually a separate, longer-lived date a human filer enters instead. Leave blank to keep using "
                "the rate validity start."
            ),
            key=f"{key_prefix}_rfa_effective_date",
        )
    with col_rfa_exp:
        rfa_expiry_date = st.date_input(
            "RFA expiry date (optional)",
            value=profile.rfa_expiry_date,
            help="Same idea as RFA effective date, for Application Expires. Leave blank to keep using the rate validity end.",
            key=f"{key_prefix}_rfa_expiry_date",
        )

    include_vertical_rates = st.checkbox(
        "Include Vertical Rates (alternate OPUS upload format)",
        value=profile.include_vertical_rates,
        help=(
            "The same rates as the RATES sheet, reshaped one row per container size instead of 4 rate "
            "columns per row - a general OPUS upload option, faster to upload. Generated for every lane by "
            "default, as one sheet. Uncheck to leave it out entirely."
        ),
        key=f"{key_prefix}_include_vertical_rates",
    )

    # DG on every lane, asked the way that lane is actually filed. TAD
    # mirrors its own export tool's opt-IN "Include Dry Dangerous" setting
    # (off by default) and answers it per service scope, since each scope
    # is a separate filed workbook; every other lane generates the
    # duplicate by default and opts out, per commodity group.
    st.markdown("#### Dangerous Goods (DG)")
    is_tad_lane = bool(lane_id and lane_id.startswith("TAD-"))
    group_descriptions = [desc for _code, desc in groups]

    tad_dg_by_scope = dict(profile.tad_dg_by_scope)
    if is_tad_lane:
        # One answer per SERVICE SCOPE rather than one for the filing.
        # Each scope is filed as its own workbook, and a round can differ
        # between them: in reference/2_OPUS/23 the AEW and AMW filings
        # carry their D/DG rows and the Japan ones, cut from the same
        # source workbook that same week, carry none. With a single
        # switch, reproducing that took two runs.
        st.caption(
            "Files a second, identical row for each base Dry (D/DR) row with CGO TYPE flipped to DG, at "
            "the same rate - a standing filing convention, not something the raw MRG states. Off by "
            "default for TAD filings, matching the team's own export tool. Each service scope is filed "
            "as its own workbook, so each answers for itself."
        )
        scope_cols = st.columns(min(3, max(len(scopes), 1)))
        for i, scope in enumerate(scopes):
            with scope_cols[i % len(scope_cols)]:
                tad_dg_by_scope[scope] = st.checkbox(
                    scope or "This filing",
                    value=profile.files_tad_dg(scope),
                    key=f"{key_prefix}_tad_dg_{scope}",
                )
        generate_dg = any(tad_dg_by_scope.get(s, False) for s in scopes)
    else:
        # On unless every group is currently skipped.
        dg_currently_on = not (
            bool(group_descriptions)
            and all(view.skip_dg_generation.get(d, False) for d in group_descriptions)
        )
        generate_dg = st.checkbox(
            "File D/DG duplicate rows",
            value=dg_currently_on,
            help=(
                "Files a second, identical row for each base Dry (D/DR) row with CGO TYPE flipped to DG, at "
                "the same rate - a standing filing convention, not something the raw MRG states. On by "
                "default for this lane. Unchecking it drops DG everywhere; to drop it for only some "
                "commodity groups, leave this checked and untick them below."
            ),
            key=f"{key_prefix}_generate_dg",
        )

    # Per group, and only for groups DG can mean something for: the ones
    # this lane files a twin for by default, plus every reefer and NOR
    # group, which can be asked for one even where the lane files none -
    # another MRG may well carry them (user, 2026-09-06), and needing a
    # code change to file them would be worse than a checkbox that starts
    # off. Groups that are neither - LAWC's OOG, an in-gauge group - have
    # no DG concept at all and get no checkbox, since a tick that cannot
    # do anything reads as a setting the tool ignored.
    skip_dg_choices: dict[str, bool] = {}
    if not is_tad_lane and generate_dg:
        dg_able = [desc for _code, desc in groups if desc in dg_twin_groups or desc in reefer_nor_groups]
        if dg_able:
            st.caption(
                "Ticked groups get DG rows. Reefer and NOR groups start unticked unless this lane files "
                "them as dangerous already - tick one to file its DG rows too (R/RF becomes a matching "
                "R/DG row; R/DR becomes a D/DG row noted REEFER DRY AS DANGEROUS)."
            )
            # Labelled with whatever the table says right now, not with
            # the last applied description - on Convert those differ until
            # Apply, and a group renamed a line above should be findable
            # here by its new name.
            labels = {key: (r.get("description") or key) for r, key in zip(edited, row_keys)}
            cols_dg = st.columns(min(3, len(dg_able)))
            for i, desc in enumerate(dg_able):
                with cols_dg[i % len(cols_dg)]:
                    # Off unless the lane files this group's twin itself -
                    # which is also the default the pipeline reads, so an
                    # untouched profile and a ticked-through one agree.
                    on_by_default = desc in dg_twin_groups
                    skip_dg_choices[desc] = not st.checkbox(
                        labels.get(desc, desc),
                        value=not view.skip_dg_generation.get(desc, not on_by_default),
                        key=f"{key_prefix}_dg_group_{desc}",
                    )
        else:
            st.caption("No commodity group in this file can carry DG rows, so there is nothing to choose here.")

    # The filing-wide flag stays the fallback for a scope nobody answered
    # for - a preset made here always answers for every scope it saw, but
    # the CLI and an older preset can still be reading just this one.
    generate_tad_dg_duplicate = generate_dg if is_tad_lane else profile.generate_tad_dg_duplicate
    include_tad_d7 = profile.include_tad_d7
    tad_d7_addon = profile.tad_d7_addon
    if is_tad_lane:
        is_aew_amw = lane_id == "TAD-AEW-AMW"
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
                    key=f"{key_prefix}_include_tad_d7",
                )
            with col_d7_amt:
                tad_d7_addon = Decimal(
                    str(
                        st.number_input(
                            "D7 add-on (added to OFT 40HC)",
                            value=float(tad_d7_addon),
                            step=50.0,
                            help="Per this filing's own Surcharges reference sheet the standard add-on is 700.",
                            key=f"{key_prefix}_tad_d7_addon",
                        )
                    )
                )

    # --- fold the edits back into a profile ---------------------------
    rows = list(zip(edited, row_keys))
    # The code column IS the override now - there is no separate
    # parsed code to compare against, and its own default was
    # already seeded as an override at Step 2.
    code_overrides = {
        key: r["code"].strip()
        for r, key in rows
        if r.get("code") and r["code"].strip()
    }
    description_overrides = {
        key: r["description"].strip()
        for r, key in rows
        if r.get("description") and r["description"].strip() and r["description"].strip() != key
    }
    sequence_overrides = {
        key: int(r["override_cmdt_seq"])
        for r, key in rows
        if r.get("override_cmdt_seq") not in (None, "")
    }
    skip_commodity_filing = {key: True for r, key in rows if r.get("skip_filing")}
    # Master DG toggle wins when it's off (skip every group, so the
    # "is DG on?" reading at the top of the section still says off on the
    # next run); when on, the per-group checkboxes decide. TAD lanes don't
    # use this dict at all - their toggle is the bool below.
    if not generate_dg and not is_tad_lane:
        skip_dg_generation = {key: True for _r, key in rows}
    else:
        # Both answers are written, not just the skips: an absent flag
        # means "off" for a reefer/NOR group and "on" for a dry one, so
        # only an explicit False says "file this group's twin" for the
        # first kind - see pipeline.run_parser.
        skip_dg_generation = dict(skip_dg_choices)
    # The FINAL description (post-override) of each row, sorted by
    # its "Order" value - this is what actually ends up on the
    # output rows, so it's what group_order needs to match against
    # (see parsers/common/ordering.py::reorder_row_set()).
    commodity_group_order = [
        (r["description"].strip() if r.get("description") else key)
        for r, key in sorted(rows, key=lambda pair: pair[0].get("order", 0))
    ]

    excluded_charge_codes = [
        c.strip().upper() for c in excluded_charge_codes_input.split(",") if c.strip()
    ]


    commodity_settings = {
        "commodity_code_overrides": code_overrides,
        "commodity_description_overrides": description_overrides,
        "commodity_sequence_overrides": sequence_overrides,
        "commodity_group_order": commodity_group_order,
        "skip_dg_generation": skip_dg_generation,
        "skip_commodity_filing": skip_commodity_filing,
    }
    if editing_scope is not None:
        # Only what this scope answers DIFFERENTLY is kept - an entry
        # equal to the filing-wide answer would freeze that group here,
        # so a later change to the shared setting would silently stop
        # reaching this scope.
        scoped = profile.model_copy(update={
            "lane_id": lane_id,
            "by_scope": _with_scope_overrides(profile, editing_scope, commodity_settings),
            "excluded_charge_codes": excluded_charge_codes,
            "rfa_effective_date": rfa_effective_date,
            "rfa_expiry_date": rfa_expiry_date,
            "include_vertical_rates": include_vertical_rates,
            "generate_tad_dg_duplicate": generate_tad_dg_duplicate,
            "tad_dg_by_scope": tad_dg_by_scope,
            "include_tad_d7": include_tad_d7,
            "tad_d7_addon": tad_d7_addon,
        })
        with preset_slot:
            _render_presets(scoped, key_prefix, lane_id)
        return scoped

    built = profile.model_copy(
        update={
            "lane_id": lane_id,
            **commodity_settings,
            "excluded_charge_codes": excluded_charge_codes,
            "rfa_effective_date": rfa_effective_date,
            "rfa_expiry_date": rfa_expiry_date,
            "include_vertical_rates": include_vertical_rates,
            "generate_tad_dg_duplicate": generate_tad_dg_duplicate,
            "tad_dg_by_scope": tad_dg_by_scope,
            "include_tad_d7": include_tad_d7,
            "tad_d7_addon": tad_d7_addon,
        }
    )
    with preset_slot:
        _render_presets(built, key_prefix, lane_id)
    return built
