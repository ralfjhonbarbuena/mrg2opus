"""Ocean Network Express house colours for the Streamlit UI.

Every value below was read off ONE's own live stylesheets
(www.one-line.com/sites/oneglobal/files/css/*.css, September 2026) rather
than eyeballed from screenshots or a logo, so the magenta is the real one
rather than a near-miss. Frequency in that CSS, which is roughly how
load-bearing each colour is: MAGENTA 157 uses, PETROL 97, PLUM 94,
INK 69, ORANGE 68.

How ONE actually uses them:
  a            { color: #bd0f72 }              -> links are magenta
  a:hover      { color: #5d0035 }              -> and darken to plum
  .btn         { background: #bd0f72; color: #fff; text-transform: uppercase }
  h1,h2,h3     { color: #004d6c; font-weight: bold; text-transform: uppercase }
  #site-footer { background: #004d6c }
  .o-c-orange  { color: #f99d21 }              -> a sparing utility accent

So: magenta carries ACTION (links, primary buttons), petrol carries
STRUCTURE (headings, chrome), orange is a sparing highlight. Deliberately
no logo, wordmark or ship imagery - palette and type only, per request.

LIGHT AND DARK. Both are real Streamlit themes declared in
.streamlit/config.toml, so the built-in appearance switch toggles the
whole app - backgrounds, text, links, borders. ONE's own site is light,
so light mode is close to a transcription; dark mode needs translating.
Two of their colours are already dark and do that work, which is why the
dark app isn't built on invented greys: INK (#002636) is the page ground
and PETROL (#004D6C) the raised panel. What can't survive the flip is
anything used as TEXT - petrol as heading text scores 1.71:1 on the ink
ground and magenta as link text 2.61:1, both under the 4.5:1 AA floor -
so each gets a lightened tint, computed rather than eyeballed. Those two
tints are the only colours here that aren't ONE's verbatim. The true
magenta is kept where it matters most, as the FILL of primary buttons,
where white-on-magenta scores 6.04:1 in either mode.
Measured on the ink ground: text 12.53:1, headings 10.11:1, links
4.59:1. On white: text 15.9:1, headings 8.6:1, links 6.04:1.

This module only covers what config.toml can't express - uppercase
headings and buttons, the magenta title rule, and the "you can edit
this" marking. Keep the two in step.
"""
from __future__ import annotations

import streamlit as st

# --- ONE's palette, verbatim ------------------------------------------------
MAGENTA = "#BD0F72"  # primary: button fills, and links on light. THE ONE colour.
PLUM = "#5D0035"  # magenta's dark partner: hover, light mode only
PETROL = "#004D6C"  # structural blue: headings on light, panels on dark
INK = "#002636"  # deepest navy: text on light, the ground on dark
ORANGE = "#F99D21"  # sparing highlight accent
TEXT_DARK = "#E5E5E5"
WHITE = "#FFFFFF"

# Captions need a different neutral per mode, and ONE's palette happens to
# carry one for each: neither works on both grounds. #666666 scores 5.74:1
# on white but only 2.75:1 on the ink ground; #999999 is the mirror image
# at 2.85:1 and 5.54:1. Using one for both is what made light-mode caption
# text hard to read - it was #999999 everywhere, under the 4.5:1 floor.
MUTED_LIGHT = "#666666"
MUTED_DARK = "#999999"

# --- The two derived tints, required by contrast on the dark ground ---------
# MAGENTA lightened 35% toward white: 4.59:1 on INK, clears AA for text.
MAGENTA_TEXT_DARK = "#D463A3"
# MAGENTA lightened 15%: a hover FILL that still holds white text (4.94:1).
MAGENTA_HOVER_DARK = "#C73387"
# PETROL lightened 75%: heading text at 10.11:1, keeping the petrol hue.
PETROL_TEXT_DARK = "#BFD2DA"

# ProximaNova is licensed and not web-served by us; ONE's own stack falls
# back to Helvetica/Arial, so we do the same rather than ship a lookalike.
FONT_STACK = "'Proxima Nova', ProximaNova, Helvetica, Arial, sans-serif"


def _is_dark() -> bool:
    """Whether the dark theme is active.

    Streamlit documents st.context.theme.type as possibly stale for one
    run right after a switch (and on a session's first load), so this is
    used ONLY for the handful of accents that have no value working on
    both grounds - heading colour and the two magenta text/hover tints.
    Everything structural (backgrounds, body text, links, borders) comes
    from config.toml's own [theme.light]/[theme.dark], which Streamlit
    always gets right. A stale read is therefore a briefly off-tone
    heading that corrects itself on the next interaction, never an
    unreadable page.
    """
    try:
        return st.context.theme.type == "dark"
    except Exception:  # noqa: BLE001 - older/newer Streamlit, or no context
        return False


def _css(dark: bool) -> str:
    heading = PETROL_TEXT_DARK if dark else PETROL
    accent_text = MAGENTA_TEXT_DARK if dark else MAGENTA
    hover_fill = MAGENTA_HOVER_DARK if dark else PLUM
    # ONE darkens to plum on hover; on a navy ground that would vanish, so
    # dark mode lightens instead - the same gesture, inverted.
    link_hover = WHITE if dark else PLUM
    focus_ring = "rgba(212, 99, 163, 0.45)" if dark else "rgba(189, 15, 114, 0.35)"
    muted = MUTED_DARK if dark else MUTED_LIGHT

    return f"""
<style>
  :root {{
    --one-magenta: {MAGENTA};
    --one-accent-text: {accent_text};
    --one-hover-fill: {hover_fill};
    --one-heading: {heading};
    --one-orange: {ORANGE};
    --one-petrol: {PETROL};
    --one-muted: {muted};
  }}

  /* Form elements don't inherit font, so they're named explicitly; the
     rest comes down from body. Deliberately NOT a broad [class*="st-"]
     sweep - that catches Streamlit's icon spans, and since those icons
     are LIGATURES of a symbol font, overriding the family renders the
     ligature's literal name ("upload" printed beside the Upload label).
     Same trap for code blocks, so both faces are pinned back below. */
  html, body, button, input, textarea, select {{
    font-family: {FONT_STACK};
  }}
  [data-testid="stIconMaterial"], [class*="material-symbols"], .material-icons {{
    font-family: "Material Symbols Rounded" !important;
  }}
  code, pre, kbd, samp {{
    font-family: "Source Code Pro", ui-monospace, SFMono-Regular, Menlo, monospace !important;
  }}

  /* Headings: bold and uppercase, which is ONE's own h1/h2/h3 rule.
     h4+ keeps normal casing so long helper captions stay readable. */
  h1, h2, h3,
  [data-testid="stMarkdownContainer"] h1,
  [data-testid="stMarkdownContainer"] h2,
  [data-testid="stMarkdownContainer"] h3 {{
    color: var(--one-heading) !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.01em;
  }}
  h4, h5, h6 {{ color: var(--one-heading) !important; font-weight: 700 !important; }}

  /* The page title sits on a magenta rule - the one place the brand
     colour is used structurally, standing in for the logo we're not
     using. Magenta holds up as a RULE on either ground. */
  h1 {{ border-bottom: 3px solid var(--one-magenta); padding-bottom: 0.3rem; }}

  a:hover {{ color: {link_hover} !important; text-decoration: underline; }}

  /* Buttons. Streamlit renders primary/secondary via kind=, which is
     stabler across versions than the generated class names. One height
     and one shape for all of them, so a row lines up however long the
     labels are - the step nav looked ragged because it mixed plain
     markdown text in among real buttons. */
  .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {{
    font-weight: 700;
    text-transform: uppercase;
    transition: 0.25s;
    min-height: 2.75rem;
    white-space: nowrap;
    padding-left: 0.75rem;
    padding-right: 0.75rem;
  }}
  /* Labels are trusted to fit; if one ever doesn't, it ellipsises rather
     than reflowing the row. */
  .stButton > button p, .stDownloadButton > button p {{
    overflow: hidden;
    text-overflow: ellipsis;
  }}
  .stButton > button[kind="primary"],
  .stDownloadButton > button[kind="primary"],
  .stFormSubmitButton > button[kind="primary"] {{
    background-color: var(--one-magenta);
    border: 1px solid var(--one-magenta);
    color: {WHITE};
  }}
  .stButton > button[kind="primary"]:hover,
  .stDownloadButton > button[kind="primary"]:hover,
  .stFormSubmitButton > button[kind="primary"]:hover {{
    background-color: var(--one-hover-fill);
    border-color: var(--one-hover-fill);
    color: {WHITE};
  }}
  /* Secondary reads as ONE's outlined .trans-btn. */
  .stButton > button[kind="secondary"], .stDownloadButton > button[kind="secondary"] {{
    background-color: transparent;
    border: 1px solid var(--one-accent-text);
    color: var(--one-accent-text);
  }}
  .stButton > button[kind="secondary"]:hover, .stDownloadButton > button[kind="secondary"]:hover {{
    background-color: var(--one-magenta);
    border-color: var(--one-magenta);
    color: {WHITE};
  }}
  .stButton > button:focus:not(:active) {{ box-shadow: 0 0 0 2px {focus_ring}; }}

  [data-testid="stProgress"] > div > div > div > div {{ background-color: var(--one-magenta) !important; }}
  .stTabs [aria-selected="true"] {{ color: var(--one-accent-text) !important; }}
  .stTabs [data-baseweb="tab-highlight"] {{ background-color: var(--one-magenta); }}

  [data-testid="stDataFrame"] thead tr th, [data-testid="stTable"] thead tr th {{
    text-transform: uppercase;
    font-weight: 700;
  }}

  /* Callouts get a coloured rule on the left, the way ONE's own notices
     do. Only the two whose default tint fights the palette are changed:
     info picks up petrol, warning ONE's own orange accent. Success and
     error keep Streamlit's green and red - those colours carry meaning
     that outranks the brand, and a filing that failed should not look
     on-brand. */
  [data-testid="stAlertContentInfo"] {{ border-left: 4px solid var(--one-heading); }}
  [data-testid="stAlertContentWarning"] {{ border-left: 4px solid var(--one-orange); }}

  /* --- "You can change this" -------------------------------------------
     Anything the user can actually edit gets a magenta left edge that
     brightens on focus, so the editable controls read as a group against
     the surrounding explanatory text. Streamlit's data_editor draws its
     cells to a <canvas>, so per-cell styling is impossible from CSS -
     that grid instead marks its editable COLUMNS in their own headers
     (see steps/step3_customize.py) and gets the same edge here. */
  [data-testid="stTextInput"], [data-testid="stNumberInput"],
  [data-testid="stDateInput"], [data-testid="stSelectbox"],
  [data-testid="stTextArea"], [data-testid="stMultiSelect"],
  [data-testid="stDataEditor"], [data-testid="stFileUploader"] {{
    border-left: 3px solid var(--one-magenta);
    padding-left: 0.75rem;
    transition: border-color 0.25s;
  }}
  [data-testid="stTextInput"]:focus-within, [data-testid="stNumberInput"]:focus-within,
  [data-testid="stDateInput"]:focus-within, [data-testid="stSelectbox"]:focus-within,
  [data-testid="stTextArea"]:focus-within, [data-testid="stMultiSelect"]:focus-within,
  [data-testid="stDataEditor"]:focus-within, [data-testid="stFileUploader"]:focus-within {{
    border-left-color: var(--one-accent-text);
  }}
  /* Their labels carry the same signal, so the edge isn't the only cue. */
  [data-testid="stWidgetLabel"] p {{ font-weight: 600; }}

  [data-baseweb="input"]:focus-within, [data-baseweb="select"]:focus-within,
  [data-baseweb="textarea"]:focus-within {{
    border-color: var(--one-accent-text) !important;
    box-shadow: 0 0 0 1px var(--one-accent-text);
  }}

  [data-testid="stCaptionContainer"], .stCaption {{ color: var(--one-muted); }}
</style>
"""


def apply_theme() -> None:
    """Inject the house styling. Call once, right after set_page_config()."""
    st.markdown(_css(_is_dark()), unsafe_allow_html=True)
