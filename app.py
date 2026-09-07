"""
QC Pipeline — Nielsen Branded Streamlit UI
Run: streamlit run app.py
"""

import re
import html as html_lib

import io
import time
import streamlit as st
import pandas as pd

from qc_check import (
    build_rate_check,
    build_exposure_check,
    build_sample,
    build_extrap_check,
    SHEETS,
)
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="QC Pipeline · Nielsen",
    page_icon="🟣",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def _flatten(content: str) -> str:
    """Collapse any HTML fragment onto a single line — see _html() below
    for exactly why this is necessary, not optional."""
    return re.sub(r"\s*\n\s*", " ", content).strip()


def _html(content: str) -> None:
    """
    Render raw HTML safely, immune to Streamlit's Markdown pre-pass.

    Streamlit runs every st.markdown() call through a CommonMark parser
    before it ever looks at unsafe_allow_html. Two CommonMark rules bite
    hand-built, multi-line HTML f-strings:

      1. A line indented 4+ spaces is read as a literal code block.
      2. A blank line (even one that's only whitespace) ends an "HTML
         block" early, kicking whatever comes after back out into plain
         Markdown — which is exactly how rule 1 gets triggered even in
         code that looks correctly indented, because nesting one f-string
         inside another (e.g. a `log_content` variable built on its own
         then interpolated into a parent template) creates a line that's
         nothing but the outer template's leading spaces followed
         immediately by the inner string's own leading newline. That
         reads as a blank line to the parser, even though it isn't blank
         in the source.

    textwrap.dedent() only strips whitespace common to every line, so it
    cannot fix nested indentation like this, and doesn't touch blank
    lines at all — collapsing the whole fragment onto a single line is
    the only fix that removes both failure modes at once, regardless of
    how deeply the call is nested or how the string was assembled.
    """
    st.markdown(_flatten(content), unsafe_allow_html=True)


def _html_into(placeholder, content: str) -> None:
    """Same safety as _html(), but writes into an st.empty() placeholder
    instead of appending to the normal page flow. This is what makes
    live, mid-run updates possible: a placeholder reserves its position
    in the page the moment st.empty() is called, and any later
    .markdown() call on it updates just that spot — including calls made
    from inside run_pipeline() while it's still executing, which is how
    stage results and log lines appear progressively instead of all at
    once at the end.
    """
    placeholder.markdown(_flatten(content), unsafe_allow_html=True)



# ── CSS ───────────────────────────────────────────────────────────────────────
_html("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,400&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,400;0,9..144,500;0,9..144,600;0,9..144,700;1,9..144,500&display=swap');

/* ═══ NIELSEN TOKENS (2021 brand refresh: violet/midnight, not red) ═══ */
:root {
    --n-violet:       #6E37FA;
    --n-violet-dark:  #5324D9;
    --n-violet-wash:  #F3EFFE;
    --n-midnight:     #002041;
    --n-midnight-2:   #0B2A4D;
    --n-paper:        #F7F6F3;
    --n-panel:        #FFFFFF;
    --n-ink:          #201C2C;
    --n-ink-soft:     #6B6779;
    --n-ink-faint:    #A6A2B3;
    --n-line:         #E9E5F2;
    --n-line-soft:    #F1EEF8;
    --n-good:         #1F9D6C;
    --n-good-wash:    #EFFAF4;
    --n-fail:         #D6455A;
    --n-fail-wash:    #FDF1F2;
    --n-warn:         #C9821E;
    --n-warn-wash:    #FDF6EA;
    --n-maxw:         1140px;
}

/* ═══ RESET ═══ */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

/* Forces every native browser control (buttons, checkboxes, file
   inputs) to render with LIGHT default chrome regardless of the user's
   OS/browser dark-mode setting. Without this, a browser in dark mode
   renders unstyled or partially-styled native controls — like the file
   uploader's "Browse files" button — using dark default colors (black
   background, white text) until a CSS state like :hover happens to
   override it. That's what "black until hover" is: not our CSS being
   wrong, but the browser's own dark-mode default filling the gap our
   CSS didn't reach. This one declaration removes that gap entirely,
   for every native control in the app, not just the ones we've styled
   so far.
*/
:root { color-scheme: light !important; }
html { color-scheme: light !important; }

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] > section.main,
.block-container {
    background: var(--n-paper) !important;
    font-family: 'Inter', -apple-system, sans-serif !important;
    color: var(--n-ink) !important;
    max-width: 100% !important;
    padding: 0 !important;
    color-scheme: light !important;
}

[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
footer, #MainMenu { display: none !important; }

[data-testid="stHorizontalBlock"] { gap: 20px !important; align-items: stretch !important; }
[data-testid="column"] { padding: 0 !important; }
[data-testid="stVerticalBlock"] { gap: 0 !important; }

[data-testid="stAppViewContainer"] > .main .block-container {
    padding: 0 0 56px !important;
}

/* Every direct content section is centered and capped, like a real product page */
.n-shell { max-width: var(--n-maxw); margin: 0 auto; padding: 0 40px; }

/* ═══ TOPBAR ═══ */
.n-topbar-outer {
    background: radial-gradient(ellipse 900px 200px at 15% 0%, rgba(110,55,250,0.28), transparent 60%),
                linear-gradient(100deg, var(--n-midnight) 0%, var(--n-midnight-2) 100%);
    border-bottom: 3px solid transparent;
    border-image: linear-gradient(90deg, var(--n-violet) 0%, #9B7CFA 45%, #4ECDC4 100%) 1;
    margin-bottom: 44px;
}

.n-topbar {
    max-width: var(--n-maxw);
    margin: 0 auto;
    height: 80px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 40px;
}

.n-topbar-left { display: flex; align-items: center; gap: 16px; }

.n-mark {
    width: 38px; height: 38px;
    border-radius: 11px;
    background: linear-gradient(135deg, var(--n-violet) 0%, var(--n-violet-dark) 100%);
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
    box-shadow: 0 4px 12px rgba(110,55,250,0.35), inset 0 1px 0 rgba(255,255,255,0.2);
}

.n-mark-play {
    width: 0; height: 0;
    border-top: 6px solid transparent;
    border-bottom: 6px solid transparent;
    border-left: 10px solid #FFFFFF;
    margin-left: 3px;
}

.n-topbar-text { display: flex; flex-direction: column; gap: 2px; }

.n-wordmark {
    font-family: 'Inter', sans-serif;
    font-size: 19px;
    font-weight: 700;
    color: #FFFFFF;
    letter-spacing: -0.3px;
    line-height: 1;
}

.n-app-name { font-size: 12px; font-weight: 500; color: rgba(255,255,255,0.6); letter-spacing: 0.01em; }

.n-version {
    font-size: 11px;
    font-weight: 600;
    color: #FFFFFF;
    font-family: 'Inter', monospace;
    background: rgba(110,55,250,0.35);
    border: 1px solid rgba(110,55,250,0.6);
    padding: 5px 12px;
    border-radius: 20px;
}

/* ═══ SECTION HEADER (numbered step) ═══ */
.n-step-head {
    display: flex;
    align-items: baseline;
    gap: 14px;
    margin-bottom: 18px;
}

.n-step-num {
    font-family: 'Fraunces', serif;
    font-style: italic;
    font-size: 13px;
    font-weight: 600;
    color: var(--n-violet);
    background: var(--n-violet-wash);
    width: 26px; height: 26px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
    transform: translateY(1px);
}

.n-step-title {
    font-family: 'Fraunces', serif;
    font-size: 18px;
    font-weight: 500;
    color: var(--n-ink);
}

.n-step-sub { font-size: 12.5px; color: var(--n-ink-faint); margin-left: 40px; margin-top: -12px; margin-bottom: 20px; }

.n-section { margin-bottom: 40px; }

/* ═══ UPLOAD GRID ═══ */
/* Each card is a real st.container(border=True) — an actual DOM parent
   Streamlit nests the header, uploader, and status line inside — not the
   old "open a <div>, render a widget, close the <div> in a separate
   st.markdown() call" trick, which produced three disconnected sibling
   nodes instead of one box. */
[data-testid="column"] [data-testid="stVerticalBlockBorderWrapper"] {
    border: 1px solid var(--n-line) !important;
    border-radius: 14px !important;
    background: var(--n-panel) !important;
    transition: border-color 0.2s, box-shadow 0.2s;
}

[data-testid="column"] [data-testid="stVerticalBlockBorderWrapper"]:hover {
    border-color: #D8D2EE !important;
    box-shadow: 0 4px 14px rgba(32,28,44,0.05) !important;
}

[data-testid="column"] [data-testid="stVerticalBlockBorderWrapper"] > div {
    padding: 16px 18px 18px !important;
    gap: 10px !important;
}

.n-upload-card-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 2px;
}

.n-upload-name { font-size: 13.5px; font-weight: 700; color: var(--n-ink); letter-spacing: -0.1px; }
.n-upload-hint { font-size: 10.5px; color: var(--n-ink-faint); font-family: 'Inter', monospace; margin-bottom: 4px; }

.n-upload-dot { width: 7px; height: 7px; border-radius: 50%; background: #E4E1EC; flex-shrink: 0; }
.n-upload-dot.loaded { background: var(--n-good); box-shadow: 0 0 0 3px rgba(31,157,108,0.15); }

.n-upload-status {
    margin-top: 6px;
    font-size: 11px;
    font-weight: 600;
    color: var(--n-ink-faint);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.n-upload-status.loaded { color: var(--n-good); }

/* The uploader widget itself — we do NOT hide its native content this
   time. Hiding the icon/text/button left a blank, undiscoverable box
   (that was the previous regression). config.toml now themes Streamlit's
   own components in Nielsen violet natively, so the native drag-drop
   text and Browse button are safe to just show — properly colored,
   guaranteed present, no CSS guessing about internal markup required. */
[data-testid="stFileUploader"] { background: transparent !important; margin: 0 !important; padding: 0 !important; color-scheme: light !important; }

/* Everything below anchors on [data-testid="stFileUploader"] — the one
   testid we have direct evidence actually matches, since other rules
   using it render correctly in your screenshots — and then selects by
   real HTML tag (section, button) rather than guessing deeper testid
   names like "stFileUploaderDropzone", which has failed to match
   several times in a row now regardless of what was written for it. */

[data-testid="stFileUploader"] section {
    background: #FBFAF8 !important;
    border: 1.5px dashed #DDD7EE !important;
    border-radius: 8px !important;
    padding: 10px 12px !important;
    min-height: unset !important;
    color-scheme: light !important;
    transition: border-color 0.2s, background 0.2s !important;
}

[data-testid="stFileUploader"] section:hover {
    border-color: var(--n-violet) !important;
    background: var(--n-violet-wash) !important;
}

[data-testid="stFileUploader"] section small {
    font-size: 9.5px !important;
    color: var(--n-ink-faint) !important;
    opacity: 0.7;
}

/* The button by TAG, not by testid — a <button> element inside the
   uploader will always be a <button>, regardless of what testid or
   internal class Streamlit's frontend build assigns it. */
[data-testid="stFileUploader"] button {
    color-scheme: light !important;
    background-color: #FFFFFF !important;
    border: 1px solid var(--n-line) !important;
    border-radius: 6px !important;
    transform: scale(0.82);
    transform-origin: left center;
}

[data-testid="stFileUploader"] button,
[data-testid="stFileUploader"] button * {
    color: var(--n-ink) !important;
}

[data-testid="stFileUploader"] button:hover {
    background-color: var(--n-violet-wash) !important;
    border-color: var(--n-violet) !important;
}

[data-testid="stFileUploader"] button:hover,
[data-testid="stFileUploader"] button:hover * {
    color: var(--n-violet-dark) !important;
}

/* Once a file is attached, Streamlit's own file chip is redundant with
   our status line below (which we build from uf.name / uf.size that we
   already control) — try both the modern and a plausible legacy testid
   for that chip so hiding it doesn't silently depend on guessing right. */
[data-testid="stFileUploaderFile"],
[data-testid="stFileUploader"] li {
    display: none !important;
}


/* ═══ ACTION ROW (progress + run button) ═══ */
.n-action-row {
    display: flex;
    align-items: center;
    gap: 28px;
    background: var(--n-panel);
    border: 1px solid var(--n-line);
    border-radius: 14px;
    padding: 22px 28px;
    margin-top: 16px;
    box-shadow: 0 2px 6px rgba(32,28,44,0.05), 0 8px 24px rgba(32,28,44,0.04);
}

.n-action-progress { flex: 1; }

.n-progress-track { height: 6px; background: var(--n-line-soft); border-radius: 4px; overflow: hidden; margin-bottom: 9px; }
.n-progress-fill { height: 100%; background: linear-gradient(90deg, var(--n-violet), var(--n-violet-dark)); border-radius: 4px; transition: width 0.4s ease; }
.n-progress-label { font-size: 12px; color: var(--n-ink-faint); font-weight: 500; }
.n-progress-label span { color: var(--n-ink); font-weight: 700; }

.n-action-row [data-testid="stButton"] { min-width: 260px; }

[data-testid="stButton"] > button {
    width: 100% !important;
    height: 50px !important;
    background: linear-gradient(135deg, var(--n-violet) 0%, var(--n-violet-dark) 100%) !important;
    color: #FFFFFF !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    letter-spacing: 0.03em !important;
    border: none !important;
    border-radius: 9px !important;
    cursor: pointer !important;
    transition: transform 0.15s ease, box-shadow 0.2s !important;
    box-shadow: 0 3px 10px rgba(110,55,250,0.32), inset 0 1px 0 rgba(255,255,255,0.18) !important;
}

[data-testid="stButton"] > button:hover:not(:disabled) {
    transform: translateY(-1px);
    box-shadow: 0 6px 18px rgba(110,55,250,0.4), inset 0 1px 0 rgba(255,255,255,0.2) !important;
}

[data-testid="stButton"] > button:disabled {
    background: #F1EFF7 !important;
    color: #BBB6CC !important;
    box-shadow: none !important;
    cursor: not-allowed !important;
    border: 1px solid var(--n-line) !important;
}

/* ═══ PIPELINE STEPPER ═══ */
.n-pipeline-bar { background: var(--n-panel); border: 1px solid var(--n-line); border-radius: 14px; padding: 30px 34px; box-shadow: 0 2px 6px rgba(32,28,44,0.05), 0 8px 24px rgba(32,28,44,0.04); }

.n-stages { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; position: relative; }

.n-stages::before {
    content: '';
    position: absolute;
    top: 18px;
    left: calc(12.5% + 18px);
    right: calc(12.5% + 18px);
    height: 2px;
    background: var(--n-line-soft);
    z-index: 0;
}

.n-stage-col { display: flex; flex-direction: column; align-items: center; gap: 12px; position: relative; z-index: 1; }

.n-stage-icon {
    width: 36px; height: 36px;
    border-radius: 50%;
    border: 2px solid #E4E1EC;
    background: #FFFFFF;
    display: flex; align-items: center; justify-content: center;
    font-size: 14px;
    font-weight: 700;
    color: #C9C4DB;
    transition: all 0.3s ease;
}

.n-stage-col.done   .n-stage-icon { border-color: var(--n-good); background: var(--n-good); color: #FFFFFF; }
.n-stage-col.active .n-stage-icon { border-color: var(--n-violet); background: var(--n-violet); color: #FFFFFF; box-shadow: 0 0 0 4px rgba(110,55,250,0.14); }
.n-stage-col.failed .n-stage-icon { border-color: var(--n-fail); background: var(--n-fail-wash); color: var(--n-fail); }
.n-stage-col.skipped .n-stage-icon { border-color: var(--n-warn); background: var(--n-warn-wash); color: var(--n-warn); }

.n-stage-label { font-size: 12.5px; font-weight: 600; color: var(--n-ink-faint); text-align: center; transition: color 0.3s; }
.n-stage-col.done   .n-stage-label { color: var(--n-good); }
.n-stage-col.active .n-stage-label { color: var(--n-violet); }
.n-stage-col.failed .n-stage-label { color: var(--n-fail); }
.n-stage-col.skipped .n-stage-label { color: var(--n-warn); }

.n-stage-sub { font-size: 10.5px; color: #C9C4DB; text-align: center; line-height: 1.4; }
.n-stage-col.done .n-stage-sub   { color: var(--n-ink-faint); }
.n-stage-col.active .n-stage-sub { color: var(--n-violet); opacity: 0.75; }

/* ═══ LOG + RESULTS ROW ═══ */
.n-row-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 18px;
}

.n-toggle-wrap [data-testid="stWidgetLabel"] p { font-size: 11.5px !important; font-weight: 600 !important; color: var(--n-ink-soft) !important; }

/* Show-logs checkbox — deliberately minimal CSS. Its color now comes
   from config.toml's primaryColor (Streamlit's own theme engine colors
   its native checkbox correctly by design), so we only touch layout —
   never the checkbox's own box/checkmark, which is what broke last
   time when transform/reset rules hit internals we couldn't verify. */
[data-testid="stCheckbox"] {
    display: flex !important;
    justify-content: flex-end !important;
    padding-top: 4px !important;
}
[data-testid="stCheckbox"] label {
    display: flex !important;
    flex-direction: row-reverse !important;
    align-items: center !important;
    gap: 8px !important;
    cursor: pointer !important;
}
[data-testid="stCheckbox"] label p {
    font-size: 12.5px !important;
    font-weight: 600 !important;
    color: var(--n-ink-soft) !important;
}

.n-log-card { background: var(--n-midnight); border-radius: 14px; overflow: hidden; box-shadow: 0 2px 6px rgba(32,28,44,0.05), 0 8px 24px rgba(32,28,44,0.04); }

.n-log-header {
    padding: 16px 24px;
    border-bottom: 1px solid rgba(255,255,255,0.08);
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: rgba(255,255,255,0.02);
}

.n-log-header-title { font-size: 11px; font-weight: 600; color: rgba(255,255,255,0.55); letter-spacing: 0.1em; text-transform: uppercase; }

.n-log-dot { width: 7px; height: 7px; border-radius: 50%; background: rgba(255,255,255,0.18); }
.n-log-dot.active { background: var(--n-violet); box-shadow: 0 0 6px rgba(110,55,250,0.6); }
.n-log-dot.done   { background: var(--n-good); }
.n-log-dot.warn   { background: var(--n-warn); }
.n-log-dot.fail   { background: var(--n-fail); box-shadow: 0 0 6px rgba(214,69,90,0.5); }

.n-log-body {
    padding: 22px 24px;
    font-family: 'SF Mono', 'Fira Code', 'Consolas', monospace;
    font-size: 12.5px;
    line-height: 1.95;
    min-height: 220px;
    max-height: 340px;
    overflow-y: auto;
}

.n-log-body::-webkit-scrollbar { width: 4px; }
.n-log-body::-webkit-scrollbar-track { background: rgba(255,255,255,0.02); }
.n-log-body::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.15); border-radius: 2px; }

.ll { display: flex; gap: 14px; padding: 3px 0; }
.ll-ts { color: rgba(255,255,255,0.22); flex-shrink: 0; }
.ll-ok   { color: #8FB8FF; }
.ll-good { color: #5CDDA0; }
.ll-fail { color: #FF7A8A; }
.ll-warn { color: #F0B84E; }
.ll-dim  { color: rgba(255,255,255,0.2); }
.ll-head { color: #B296FF; font-weight: 600; }
.ll-stage { color: rgba(255,255,255,0.35); }
.ll-div { border: none; border-top: 1px solid rgba(255,255,255,0.08); margin: 12px 0; }

.n-empty { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 10px; padding: 50px 40px; text-align: center; }
.n-empty-icon { font-size: 24px; opacity: 0.5; color: var(--n-violet); }
.n-empty-title { font-family: 'Fraunces', serif; font-size: 14px; font-weight: 500; font-style: italic; color: rgba(255,255,255,0.75); }
.n-empty-sub { font-size: 11.5px; color: rgba(255,255,255,0.35); max-width: 260px; line-height: 1.6; }

.n-log-compact {
    background: var(--n-panel);
    border: 1px solid var(--n-line);
    border-radius: 14px;
    padding: 22px 26px;
    display: flex;
    align-items: center;
    gap: 16px;
}

.n-log-compact-dot { width: 9px; height: 9px; border-radius: 50%; background: #E4E1EC; flex-shrink: 0; }
.n-log-compact-dot.active { background: var(--n-violet); box-shadow: 0 0 0 4px rgba(110,55,250,0.14); }
.n-log-compact-dot.done   { background: var(--n-good); box-shadow: 0 0 0 4px rgba(31,157,108,0.14); }
.n-log-compact-dot.fail   { background: var(--n-fail); box-shadow: 0 0 0 4px rgba(214,69,90,0.14); }

.n-log-compact-text { font-size: 13px; font-weight: 600; color: var(--n-ink); }
.n-log-compact-sub { font-size: 11px; color: var(--n-ink-faint); margin-top: 1px; }

/* ═══ RESULTS ═══ */
/* General card chrome for ANY st.container(border=True) in the app,
   not just ones sitting inside a column (the upload-card rule above
   only matches inside [data-testid="column"], so this results card —
   which lives directly in the page flow — needs its own base rule). */
[data-testid="stVerticalBlockBorderWrapper"] {
    border: 1px solid var(--n-line) !important;
    border-radius: 14px !important;
    background: var(--n-panel) !important;
}

/* Distinct, more generous padding for this specific card via
   Streamlit's container key= (adds a stable .st-key-<name> class) —
   the upload cards are small utility slots, this is a summary card
   and needs real breathing room, which is what "cramped" meant. */
.st-key-output_results {
    box-shadow: 0 2px 6px rgba(32,28,44,0.05), 0 8px 24px rgba(32,28,44,0.04) !important;
}
.st-key-output_results > div {
    padding: 10px 14px !important;
    gap: 0 !important;
}

.n-result-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 18px; }

.n-result-card { border-radius: 12px; padding: 22px 24px; border: 1px solid var(--n-line-soft); background: #FCFBFA; }
.n-result-card.ok   { border-color: rgba(31,157,108,0.28); background: var(--n-good-wash); }
.n-result-card.fail { border-color: rgba(214,69,90,0.25);  background: var(--n-fail-wash); }
.n-result-card.skip { border-color: rgba(201,130,30,0.25); background: var(--n-warn-wash); }

.n-rc-name { font-size: 10.5px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--n-ink-faint); margin-bottom: 16px; }
.n-rc-stat { font-size: 26px; font-weight: 800; line-height: 1; margin-bottom: 10px; }
.n-result-card.ok   .n-rc-stat { color: var(--n-good); }
.n-result-card.fail .n-rc-stat { color: var(--n-fail); }
.n-result-card.skip .n-rc-stat { color: var(--n-warn); }
.n-rc-sub { font-size: 11px; color: var(--n-ink-faint); line-height: 1.5; }
.n-rc-err {
    font-size: 10.5px;
    color: var(--n-fail);
    line-height: 1.5;
    margin-top: 4px;
    font-family: 'SF Mono', 'Fira Code', 'Consolas', monospace;
    word-break: break-word;
}

/* ═══ SUMMARY BANNER (top of page, after a run) ═══ */
.n-banner {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 14px 22px;
    border-radius: 12px;
    font-size: 14px;
    font-weight: 700;
    margin-bottom: 28px;
    border: 1px solid transparent;
}
.n-banner-icon { font-size: 16px; line-height: 1; }
.n-banner-active {
    background: var(--n-violet-wash);
    border-color: rgba(110,55,250,0.25);
    color: var(--n-violet-dark);
}
.n-banner-pass {
    background: var(--n-good-wash);
    border-color: rgba(31,157,108,0.28);
    color: var(--n-good);
}
.n-banner-partial {
    background: var(--n-warn-wash);
    border-color: rgba(201,130,30,0.28);
    color: var(--n-warn);
}
.n-banner-fail {
    background: var(--n-fail-wash);
    border-color: rgba(214,69,90,0.28);
    color: var(--n-fail);
}

/* ═══ RUN HISTORY ═══ */
/* Styled via the container's key= — same proven pattern as
   .st-key-output_results — rather than a manually opened/closed <div>,
   which would produce disconnected sibling nodes around the st.columns()
   rows exactly like the upload-card and output-results bugs fixed
   earlier in this app's history. */
.st-key-run_history {
    box-shadow: 0 2px 6px rgba(32,28,44,0.05), 0 8px 24px rgba(32,28,44,0.04) !important;
}
.st-key-run_history > div {
    padding: 6px 24px !important;
    gap: 0 !important;
}
.n-history-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 0;
    gap: 16px;
}
.n-history-left { display: flex; align-items: center; gap: 14px; min-width: 0; }
.n-history-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.n-history-dot.pass    { background: var(--n-good); }
.n-history-dot.partial { background: var(--n-warn); }
.n-history-dot.fail    { background: var(--n-fail); }
.n-history-time { font-size: 12.5px; font-weight: 600; color: var(--n-ink); font-family: 'Inter', monospace; }
.n-history-status { font-size: 11.5px; color: var(--n-ink-faint); }

[data-testid="stDownloadButton"] > button {
    height: 46px !important;
    background: var(--n-midnight) !important;
    color: #FFFFFF !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0 26px !important;
    letter-spacing: 0.02em !important;
    cursor: pointer !important;
    transition: background 0.2s !important;
    white-space: nowrap !important;
}

[data-testid="stDownloadButton"] > button:hover { background: var(--n-midnight-2) !important; }

/* Compact variant for the small per-row download buttons in Run History
   — scoped via the container's key= so it doesn't affect the main
   Step 4 download button. */
.st-key-run_history [data-testid="stDownloadButton"] > button {
    height: 32px !important;
    padding: 0 14px !important;
    font-size: 11px !important;
}

[data-testid="stSpinner"] > div { color: var(--n-violet) !important; }
.uploadedFileName { color: var(--n-ink) !important; font-size: 11px !important; }
</style>
""")

# ── Session state ─────────────────────────────────────────────────────────────
if "stage_states" not in st.session_state:
    st.session_state.stage_states = {1:"idle", 2:"idle", 3:"idle", 4:"idle"}
if "log_lines" not in st.session_state:
    st.session_state.log_lines = []
if "result_stats" not in st.session_state:
    st.session_state.result_stats = None
if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "show_logs" not in st.session_state:
    st.session_state.show_logs = True
if "pipeline_running" not in st.session_state:
    st.session_state.pipeline_running = False
if "run_history" not in st.session_state:
    st.session_state.run_history = []

# ── Shared definitions (used by both the live-update path inside
#    run_pipeline and the normal top-to-bottom render pass, so the two
#    can never drift into showing different markup for the same state) ──
STAGE_DEFS = [
    (1, "RateCheck",     "Adapt · BSR · YT"),
    (2, "ExposureCheck", "Sample vs Matex"),
    (3, "Sample",        "pID · matchdays"),
    (4, "ExtrapCheck",   "BT% / Msec%"),
]
ICON_MAP = {"idle":"—","done":"✓","active":"◎","failed":"✗","skipped":"⚠"}

# Live-update channel: run_pipeline writes into these placeholders (when
# given) as it executes, so stage results and log lines appear on screen
# progressively instead of all at once when the function returns. Plain
# module-level variables are safe here — Streamlit re-executes this
# whole file top-to-bottom on every interaction, so these are freshly
# None at the start of every run and only ever populated for the
# duration of the one run_pipeline() call that uses them.
_LIVE = {"stage_ph": None, "log_ph": None, "banner_ph": None}


def render_stage_stepper_into(ph):
    cols_html = ""
    for n, title, sub in STAGE_DEFS:
        state = st.session_state.stage_states[n]
        icon = ICON_MAP.get(state, "—")
        if state == "done" and st.session_state.result_stats:
            s = st.session_state.result_stats.get(n, {})
            rows = s.get("rows")
            if rows:
                sub = f"{rows:,} rows"
            if n == 3 and s.get("matchdays"):
                sub += f" · {s['matchdays']} matchdays"
        cols_html += f"""
        <div class="n-stage-col {state}">
          <div class="n-stage-icon">{icon}</div>
          <div class="n-stage-label">{title}</div>
          <div class="n-stage-sub">{sub}</div>
        </div>"""
    _html_into(ph, f"""
    <div class="n-pipeline-bar">
      <div class="n-stages">{cols_html}</div>
    </div>
    """)


def render_log_into(ph):
    done_count = sum(1 for s in (st.session_state.result_stats or {}).values() if s.get("ok"))
    if st.session_state.pipeline_running:
        dot_class = "active"
    elif st.session_state.result_stats:
        dot_class = "fail" if done_count == 0 else ("warn" if done_count < 4 else "done")
    elif st.session_state.log_lines:
        dot_class = "active"
    else:
        dot_class = ""

    if st.session_state.show_logs:
        if not st.session_state.log_lines:
            log_content = """
            <div class="n-empty">
              <div class="n-empty-icon">○</div>
              <div class="n-empty-title">No pipeline run yet</div>
              <div class="n-empty-sub">Upload all 5 source files and press Run QC Pipeline</div>
            </div>"""
        else:
            rows_html = ""
            for stamp, text, kind in st.session_state.log_lines:
                if text == "":
                    rows_html += '<div style="height:8px"></div>'
                elif text.startswith("──"):
                    rows_html += f'<hr class="ll-div"><div class="ll"><span class="ll-ts">{stamp}</span><span class="ll-stage">{text}</span></div>'
                else:
                    cls = {"ok":"ll-ok","good":"ll-good","fail":"ll-fail","warn":"ll-warn","dim":"ll-dim","head":"ll-head","stage":"ll-stage"}.get(kind,"ll-ok")
                    rows_html += f'<div class="ll"><span class="ll-ts">{stamp}</span><span class="{cls}">{text}</span></div>'
            log_content = f'<div class="n-log-body">{rows_html}</div>'

        _html_into(ph, f"""
        <div class="n-log-card">
          <div class="n-log-header">
            <span class="n-log-header-title">Live Output</span>
            <div class="n-log-dot {dot_class}"></div>
          </div>
          {log_content}
        </div>
        """)
    else:
        if st.session_state.pipeline_running:
            compact_dot, compact_text, compact_sub = "active", "Running…", "Toggle logs on for step-by-step detail"
        elif not st.session_state.log_lines:
            compact_dot, compact_text, compact_sub = "", "No pipeline run yet", "Toggle logs on for step-by-step detail"
        elif st.session_state.result_stats and done_count == 4:
            compact_dot, compact_text, compact_sub = "done", "Complete — all 4 sheets written", "Toggle logs on for step-by-step detail"
        elif st.session_state.result_stats and done_count > 0:
            compact_dot, compact_text, compact_sub = "fail", f"Partial — {done_count}/4 sheets written", "Toggle logs on to see what failed"
        elif st.session_state.result_stats:
            compact_dot, compact_text, compact_sub = "fail", "Failed — no output written", "Toggle logs on to see what failed"
        else:
            compact_dot, compact_text, compact_sub = "active", "Running…", "Toggle logs on for step-by-step detail"

        _html_into(ph, f"""
        <div class="n-log-compact">
          <div class="n-log-compact-dot {compact_dot}"></div>
          <div>
            <div class="n-log-compact-text">{compact_text}</div>
            <div class="n-log-compact-sub">{compact_sub}</div>
          </div>
        </div>
        """)


def render_banner_into(ph):
    if st.session_state.pipeline_running:
        _html_into(ph, '<div class="n-banner n-banner-active"><span class="n-banner-icon">⟳</span><span>Running pipeline…</span></div>')
        return
    if not st.session_state.result_stats:
        ph.markdown("", unsafe_allow_html=True)
        return
    done_count = sum(1 for s in st.session_state.result_stats.values() if s.get("ok"))
    if done_count == 4:
        cls, icon, text = "pass", "✓", "4/4 passed"
    elif done_count > 0:
        cls, icon, text = "partial", "◐", f"{done_count}/4 passed — partial"
    else:
        cls, icon, text = "fail", "✗", "Failed — 0/4 passed"
    _html_into(ph, f'<div class="n-banner n-banner-{cls}"><span class="n-banner-icon">{icon}</span><span>{text}</span></div>')


# ── Helpers ───────────────────────────────────────────────────────────────────
def ts():
    return time.strftime("%H:%M:%S")

def add_log(text, kind="ok"):
    st.session_state.log_lines.append((ts(), text, kind))
    if _LIVE["log_ph"] is not None:
        render_log_into(_LIVE["log_ph"])

def set_stage(n, state):
    st.session_state.stage_states[n] = state
    if _LIVE["stage_ph"] is not None:
        render_stage_stepper_into(_LIVE["stage_ph"])
    if _LIVE["banner_ph"] is not None:
        render_banner_into(_LIVE["banner_ph"])

# ── Excel writer ──────────────────────────────────────────────────────────────
def write_output_styled(sheets: dict) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    HDR_FILL  = PatternFill("solid", fgColor="1A1A1A")
    HDR_FONT  = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    CELL_FONT = Font(name="Calibri", size=10)
    ODD_FILL  = PatternFill("solid", fgColor="F9F9F9")
    EVN_FILL  = PatternFill("solid", fgColor="FFFFFF")
    T_FILL    = PatternFill("solid", fgColor="F0FBF6"); T_FONT = Font(name="Calibri", size=10, color="22A06B", bold=True)
    F_fill    = PatternFill("solid", fgColor="FFF5F5"); F_FONT = Font(name="Calibri", size=10, color="CC0000", bold=True)
    thin      = Side(style="thin", color="EEEEEE")
    BORDER    = Border(bottom=thin)

    for name, df in sheets.items():
        ws = wb.create_sheet(title=name)
        ws.sheet_properties.tabColor = "FFFF00"

        for ci, col in enumerate(df.columns, 1):
            c = ws.cell(row=1, column=ci, value=str(col))
            c.fill=HDR_FILL; c.font=HDR_FONT; c.alignment=HDR_ALIGN; c.border=BORDER

        for ri, row in enumerate(df.itertuples(index=False), 2):
            fill = ODD_FILL if ri % 2 == 0 else EVN_FILL
            for ci, val in enumerate(row, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                cname = df.columns[ci-1]
                if cname == "Check":
                    if val is True:
                        c.fill=T_FILL; c.font=T_FONT; c.value="TRUE"
                    elif val is False:
                        c.fill=F_fill; c.font=F_FONT; c.value="FALSE"
                    else:
                        c.fill=fill; c.font=CELL_FONT
                else:
                    c.fill=fill; c.font=CELL_FONT

        for col in ws.columns:
            ltr = get_column_letter(col[0].column)
            mx = max((len(str(c.value)) for c in col if c.value), default=8)
            ws.column_dimensions[ltr].width = min(max(mx + 2, 8), 36)

        ws.freeze_panes = "A2"
        if df.shape[1] > 0:
            ws.auto_filter.ref = f"A1:{get_column_letter(df.shape[1])}1"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── Pipeline runner ───────────────────────────────────────────────────────────
def run_pipeline(adapt_f, bsr_f, yt_f, matex_f, sample_f,
                  live_stage_ph=None, live_log_ph=None, live_banner_ph=None):
    # Wire up the live-update channel for the duration of this call only.
    # Every add_log()/set_stage() from here on will push straight into
    # these placeholders as it happens, which is what makes stage
    # results and log lines appear on screen progressively rather than
    # all at once when this function returns.
    _LIVE["stage_ph"]  = live_stage_ph
    _LIVE["log_ph"]    = live_log_ph
    _LIVE["banner_ph"] = live_banner_ph
    st.session_state.pipeline_running = True

    try:
        st.session_state.log_lines = []
        st.session_state.result_stats = None
        st.session_state.output_bytes = None
        for n in [1,2,3,4]: set_stage(n, "idle")
        if live_banner_ph is not None:
            render_banner_into(live_banner_ph)

        add_log("QC CHECK AUTOMATION PIPELINE", "head")
        add_log(f"Started at {time.strftime('%Y-%m-%d %H:%M:%S')}", "dim")
        add_log("", "dim")
        add_log("── LOAD ──────────────────────────────────────────────", "stage")

        load_ok = True
        frames  = {}

        for key, uf, sheet, hdr, label in [
            ("adapt",  adapt_f,  SHEETS["adapt"],  0, "Adapt"),
            ("bsr",    bsr_f,    SHEETS["bsr"],    5, "BSR"),
            ("yt",     yt_f,     SHEETS["yt"],     0, "YouTube MM"),
            ("matex",  matex_f,  SHEETS["matex"],  0, "Matex"),
            ("sample", sample_f, SHEETS["sample"], 0, "Sample"),
        ]:
            try:
                buf = io.BytesIO(uf.read())
                df  = pd.read_excel(buf, sheet_name=sheet, header=hdr, engine="openpyxl")
                df.columns = [str(c).strip() for c in df.columns]
                frames[key] = df
                add_log(f"  [✓]  {label:<14}  {len(df):,} rows · {len(df.columns)} cols", "good")
            except Exception as e:
                add_log(f"  [✗]  {label:<14}  {e}", "fail")
                load_ok = False

        if not load_ok:
            add_log("", "dim"); add_log("  PIPELINE ABORTED — fix load errors above", "fail")
            for n in [1, 2, 3, 4]:
                set_stage(n, "failed")
            st.session_state.result_stats = {
                n: {"ok": False, "error": "Load failed — see log above"} for n in [1, 2, 3, 4]
            }
            return

        add_log("", "dim")
        add_log("── STAGES ────────────────────────────────────────────", "stage")

        output_sheets  = {}
        sample_summary = None
        # stats IS st.session_state.result_stats from this point on (same
        # object, not a copy) — mutating stats[n] below is what the live
        # placeholders read mid-run, stage by stage, as each one lands.
        stats = st.session_state.result_stats = {}

        # Stage 1
        set_stage(1, "active")
        try:
            df = build_rate_check(frames["adapt"], frames["bsr"], frames["yt"])
            output_sheets["RateCheck"] = df
            stats[1] = {"rows": len(df), "ok": True}
            set_stage(1, "done")
            add_log(f"  [✓]  RateCheck        {len(df):,} rows", "good")
        except Exception as e:
            stats[1] = {"ok": False, "error": str(e)}
            set_stage(1, "failed")
            add_log(f"  [✗]  RateCheck        {e}", "fail")

        # Stage 2
        set_stage(2, "active")
        try:
            df = build_exposure_check(frames["sample"], frames["matex"])
            output_sheets["ExposureCheck"] = df
            stats[2] = {"rows": len(df), "ok": True}
            set_stage(2, "done")
            add_log(f"  [✓]  ExposureCheck    {len(df):,} rows", "good")
        except Exception as e:
            stats[2] = {"ok": False, "error": str(e)}
            set_stage(2, "failed")
            add_log(f"  [✗]  ExposureCheck    {e}", "fail")

        # Stage 3
        set_stage(3, "active")
        try:
            ss, sample_summary = build_sample(frames["sample"])
            output_sheets["Sample"] = ss
            stats[3] = {"rows": len(ss), "matchdays": len(sample_summary), "ok": True}
            set_stage(3, "done")
            add_log(f"  [✓]  Sample           {len(ss):,} rows · {len(sample_summary)} matchdays", "good")
        except Exception as e:
            stats[3] = {"ok": False, "error": str(e)}
            set_stage(3, "failed")
            add_log(f"  [✗]  Sample           {e}", "fail")

        # Stage 4
        set_stage(4, "active")
        try:
            if sample_summary is None:
                raise RuntimeError("Stage 3 must succeed before Stage 4 can run")
            df = build_extrap_check(frames["adapt"], sample_summary)
            output_sheets["ExtrapCheck"] = df
            stats[4] = {"rows": len(df), "ok": True}
            set_stage(4, "done")
            add_log(f"  [✓]  ExtrapCheck      {len(df):,} rows", "good")
        except RuntimeError as e:
            stats[4] = {"ok": False, "skipped": True, "error": str(e)}
            set_stage(4, "skipped")
            add_log(f"  [⚠]  ExtrapCheck      SKIPPED — {e}", "warn")
        except Exception as e:
            stats[4] = {"ok": False, "error": str(e)}
            set_stage(4, "failed")
            add_log(f"  [✗]  ExtrapCheck      {e}", "fail")

        add_log("", "dim")
        add_log("── OUTPUT ────────────────────────────────────────────", "stage")
        done = sum(1 for s in stats.values() if s.get("ok"))

        if output_sheets:
            try:
                st.session_state.output_bytes = write_output_styled(output_sheets)
                add_log(f"  [✓]  {done}/4 sheets written → QC_Output.xlsx", "good")
            except Exception as e:
                add_log(f"  [✗]  Write failed: {e}", "fail")

        add_log("", "dim")
        if done == 4:
            add_log("  ✓  COMPLETE — all 4 sheets written", "good")
        elif done > 0:
            add_log(f"  ⚠  PARTIAL — {done}/4 sheets written", "warn")
        else:
            add_log("  ✗  FAILED — no output written", "fail")

    finally:
        # This runs on every exit path — normal completion, the early
        # return on load failure, or (defensively) any exception that
        # somehow escapes the try/excepts above — so pipeline_running
        # can never get stuck True and the history entry is never
        # skipped or duplicated regardless of how the run ended.
        st.session_state.pipeline_running = False

        if st.session_state.result_stats is not None:
            done_final = sum(1 for s in st.session_state.result_stats.values() if s.get("ok"))
            st.session_state.run_history.insert(0, {
                "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "done": done_final,
                "total": 4,
                "output_bytes": st.session_state.output_bytes,
            })
            st.session_state.run_history = st.session_state.run_history[:5]

        if live_banner_ph is not None:
            render_banner_into(live_banner_ph)
        if live_stage_ph is not None:
            render_stage_stepper_into(live_stage_ph)
        if live_log_ph is not None:
            render_log_into(live_log_ph)

        _LIVE["stage_ph"] = None
        _LIVE["log_ph"] = None
        _LIVE["banner_ph"] = None


# ══════════════════════════════════════════════════════════════════════════════
# RENDER
# ══════════════════════════════════════════════════════════════════════════════

# ── TOPBAR ───────────────────────────────────────────────────────────────────
_html("""
<div class="n-topbar-outer">
  <div class="n-topbar">
    <div class="n-topbar-left">
      <div class="n-mark"><div class="n-mark-play"></div></div>
      <div class="n-topbar-text">
        <span class="n-wordmark">Nielsen</span>
        <span class="n-app-name">QC Automation Pipeline</span>
      </div>
    </div>
    <span class="n-version">v7.0 · Media Measurement</span>
  </div>
</div>
""")

_html('<div class="n-shell">')

# Banner placeholder — reserved right after the topbar so a completed
# run's pass/fail state is visible without scrolling. Populated below,
# after Step 1's button logic runs (or on every normal rerun from
# whatever session_state already holds).
banner_ph = st.empty()

# ═══════════════════════════════════════════════════════════════════════════
# STEP 1 — SOURCE FILES
# ═══════════════════════════════════════════════════════════════════════════
_html('<div class="n-section">')
_html("""
<div class="n-step-head">
  <div class="n-step-num">1</div>
  <div class="n-step-title">Source Files</div>
</div>
<div class="n-step-sub">Upload all five workbooks to unlock the pipeline run</div>
""")

FILE_DEFS = [
    ("adapt",  "Adapt",      "programs"),
    ("bsr",    "BSR",        "Database · row 6"),
    ("yt",     "YouTube MM", "MM"),
    ("matex",  "Matex",      "exposures"),
    ("sample", "Sample",     "programs"),
]

uploaded = {}
loaded_count = 0
upload_cols = st.columns(5, gap="small")

for col, (key, label, hint) in zip(upload_cols, FILE_DEFS):
    with col:
        with st.container(border=True):
            uf = st.session_state.get(f"upload_{key}")
            loaded_now = uf is not None
            _html(f"""
            <div class="n-upload-card-head">
              <span class="n-upload-name">{label}</span>
              <div class="n-upload-dot{' loaded' if loaded_now else ''}"></div>
            </div>
            <div class="n-upload-hint">sheet: {hint}</div>
            """)
            uf = st.file_uploader(
                label, type=["xlsx"], key=f"upload_{key}",
                help=f"Sheet: {hint}", label_visibility="collapsed",
            )
            uploaded[key] = uf
            if uf:
                loaded_count += 1
                status_txt = f"✓  {uf.name}  ·  {uf.size/1024:.0f} KB"
            else:
                status_txt = "Awaiting upload"
            _html(f'<div class="n-upload-status{" loaded" if uf else ""}">{status_txt}</div>')

# Action row — progress + run button
pct = int(loaded_count / 5 * 100)
all_ready = loaded_count == 5
btn_label = "Run QC Pipeline →" if all_ready else f"Waiting for {5 - loaded_count} more file{'s' if 5-loaded_count != 1 else ''}"

_html('<div class="n-action-row">')
action_progress_col, action_btn_col = st.columns([2.2, 1], gap="medium")
with action_progress_col:
    _html(f"""
    <div class="n-action-progress">
      <div class="n-progress-track">
        <div class="n-progress-fill" style="width:{pct}%"></div>
      </div>
      <div class="n-progress-label"><span>{loaded_count}</span> of 5 files ready</div>
    </div>
    """)
with action_btn_col:
    # The click is captured here, but run_pipeline() itself isn't called
    # until after stage_ph/log_ph below are declared — Streamlit reserves
    # each placeholder's position in the page the moment st.empty() runs,
    # so run_pipeline can write into them later in the script and the
    # updates still land in the right spot (Step 2 / Step 3), not here.
    run_clicked = st.button(btn_label, disabled=not all_ready, use_container_width=True)
_html('</div>')

_html('</div>')  # end n-section (step 1)

# ═══════════════════════════════════════════════════════════════════════════
# STEP 2 — PIPELINE STAGES
# ═══════════════════════════════════════════════════════════════════════════
_html('<div class="n-section">')
_html("""
<div class="n-step-head">
  <div class="n-step-num">2</div>
  <div class="n-step-title">Pipeline Stages</div>
</div>
""")
stage_ph = st.empty()
_html('</div>')  # end n-section (step 2)

# ═══════════════════════════════════════════════════════════════════════════
# STEP 3 — LOG
# ═══════════════════════════════════════════════════════════════════════════
_html('<div class="n-section">')

head_col1, head_col2 = st.columns([3, 1], gap="small")
with head_col1:
    _html("""
    <div class="n-step-head" style="margin-bottom:0;">
      <div class="n-step-num">3</div>
      <div class="n-step-title">Pipeline Log</div>
    </div>
    """)
with head_col2:
    st.session_state.show_logs = st.checkbox(
        "Show logs", value=st.session_state.show_logs, key="show_logs_checkbox",
    )

_html('<div style="height:18px"></div>')
log_ph = st.empty()
_html('</div>')  # end n-section (step 3)

# ═══════════════════════════════════════════════════════════════════════════
# Run the pipeline now, if the button was clicked — using the
# placeholders declared above so results appear live, stage by stage,
# in their correct positions on the page rather than all at once.
# ═══════════════════════════════════════════════════════════════════════════
if run_clicked:
    run_pipeline(
        uploaded["adapt"], uploaded["bsr"], uploaded["yt"],
        uploaded["matex"], uploaded["sample"],
        live_stage_ph=stage_ph, live_log_ph=log_ph, live_banner_ph=banner_ph,
    )
    st.rerun()

# Always render current state into the three placeholders above — this
# is what covers page load, and reruns triggered by something other than
# the Run button (e.g. toggling "Show logs"), where run_pipeline() never
# executes this script pass at all.
render_banner_into(banner_ph)
render_stage_stepper_into(stage_ph)
render_log_into(log_ph)

# ═══════════════════════════════════════════════════════════════════════════
# STEP 4 — OUTPUT
# ═══════════════════════════════════════════════════════════════════════════
if st.session_state.result_stats:
    _html('<div class="n-section">')
    with st.container(border=True, key="output_results"):
        res_head_col1, res_head_col2 = st.columns([3, 1], gap="small")
        with res_head_col1:
            _html("""
            <div class="n-step-head" style="margin-bottom:0;">
              <div class="n-step-num">4</div>
              <div class="n-step-title">Results</div>
            </div>
            """)
        with res_head_col2:
            if st.session_state.output_bytes:
                fname = f"QC_Output_{time.strftime('%Y%m%d_%H%M')}.xlsx"
                st.download_button(
                    label="↓  Download",
                    data=st.session_state.output_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        _html('<div style="height:26px"></div>')

        stats = st.session_state.result_stats
        names = ["RateCheck","ExposureCheck","Sample","ExtrapCheck"]

        cards_html = ""
        for i, sname in enumerate(names, 1):
            s = stats.get(i, {})
            ok      = s.get("ok", False)
            skipped = s.get("skipped", False)
            cls     = "ok" if ok else ("skip" if skipped else "fail")
            icon    = "✓" if ok else ("⚠" if skipped else "✗")
            rows    = s.get("rows")
            err_html = ""
            if ok:
                sub = f"{rows:,} rows" if rows else "complete"
                if i == 3 and s.get("matchdays"): sub += f"<br>{s['matchdays']} matchdays"
            elif skipped:
                sub = "Stage 3 required"
            else:
                sub = "See log above"
                raw_err = s.get("error", "")
                if raw_err:
                    safe_err = html_lib.escape(raw_err)[:90]
                    if len(raw_err) > 90:
                        safe_err += "…"
                    err_html = f'<div class="n-rc-err">{safe_err}</div>'

            cards_html += f"""
            <div class="n-result-card {cls}">
              <div class="n-rc-name">{sname}</div>
              <div class="n-rc-stat">{icon}</div>
              <div class="n-rc-sub">{sub}</div>
              {err_html}
            </div>"""

        _html(f'<div class="n-result-grid">{cards_html}</div>')
    _html('</div>')  # end n-section (step 4)

# ═══════════════════════════════════════════════════════════════════════════
# STEP 5 — RUN HISTORY
# ═══════════════════════════════════════════════════════════════════════════
if st.session_state.run_history:
    _html('<div class="n-section">')
    _html("""
    <div class="n-step-head">
      <div class="n-step-num">5</div>
      <div class="n-step-title">Run History</div>
    </div>
    <div class="n-step-sub">Last 5 runs this session</div>
    """)
    with st.container(key="run_history", border=True):
        for idx, run in enumerate(st.session_state.run_history):
            done, total = run["done"], run["total"]
            if done == total:
                dcls, dtext = "pass", f"{done}/{total} passed"
            elif done > 0:
                dcls, dtext = "partial", f"{done}/{total} passed — partial"
            else:
                dcls, dtext = "fail", f"{done}/{total} passed — failed"

            row_col1, row_col2 = st.columns([4, 1], gap="small")
            with row_col1:
                _html(f"""
                <div class="n-history-row" style="border-bottom:none; padding-bottom:0;">
                  <div class="n-history-left">
                    <div class="n-history-dot {dcls}"></div>
                    <div>
                      <div class="n-history-time">{run['time']}</div>
                      <div class="n-history-status">{dtext}</div>
                    </div>
                  </div>
                </div>
                """)
            with row_col2:
                if run.get("output_bytes"):
                    st.download_button(
                        "↓", data=run["output_bytes"],
                        file_name=f"QC_Output_{run['time'].replace(':','').replace(' ','_').replace('-','')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"hist_dl_{idx}", use_container_width=True,
                    )
            if idx < len(st.session_state.run_history) - 1:
                _html('<hr class="ll-div" style="border-top:1px solid var(--n-line-soft); margin:2px 0 10px;">')
    _html('</div>')  # end n-section (step 5)

_html('</div>')  # end n-shell

