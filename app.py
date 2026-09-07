"""
QC Pipeline — Nielsen Branded Streamlit UI
Run: streamlit run app.py
"""

import re
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


def _html(content: str) -> None:
    flat = re.sub(r"\s*\n\s*", " ", content).strip()
    st.markdown(flat, unsafe_allow_html=True)


# ── CSS ───────────────────────────────────────────────────────────────────────
_html("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,400&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,400;0,9..144,500;0,9..144,600;0,9..144,700;1,9..144,500&display=swap');

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
    --n-good:         #1F9D6C;
    --n-good-wash:    #EFFAF4;
    --n-fail:         #D6455A;
    --n-fail-wash:    #FDF1F2;
    --n-warn:         #C9821E;
    --n-warn-wash:    #FDF6EA;
    --n-maxw:         1280px;
}

/* ═══ RESET & MAIN LAYOUT ═══ */
*, *::before, *::after { box-sizing: border-box; }

[data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] > section.main,
.block-container {
    background: var(--n-paper) !important;
    font-family: 'Inter', sans-serif !important;
    color: var(--n-ink) !important;
    max-width: 100% !important;
    padding: 0 !important;
}

[data-testid="stHeader"], footer { display: none !important; }

.n-shell { max-width: var(--n-maxw); margin: 0 auto; padding: 0 48px; }

/* ═══ TOPBAR ═══ */
.n-topbar-outer {
    background: linear-gradient(100deg, var(--n-midnight) 0%, var(--n-midnight-2) 100%);
    border-bottom: 3px solid var(--n-violet);
    margin-bottom: 40px;
}
.n-topbar {
    max-width: var(--n-maxw);
    margin: 0 auto;
    height: 80px;
    display: flex;
    align-items: center;
    padding: 0 48px;
}
.n-mark {
    width: 36px; height: 36px;
    border-radius: 8px;
    background: var(--n-violet);
    display: flex; align-items: center; justify-content: center;
    font-family: 'Fraunces', serif; font-style: italic; font-weight: 700; font-size: 18px; color: #FFF;
    margin-right: 16px;
}
.n-wordmark { font-family: 'Fraunces', serif; font-size: 18px; font-weight: 600; font-style: italic; color: #FFF; }

/* ═══ TYPOGRAPHY ═══ */
.n-step-head { display: flex; align-items: center; gap: 14px; margin: 32px 0 8px 0; }
.n-step-num {
    font-family: 'Fraunces', serif; font-style: italic; font-size: 14px; font-weight: 600;
    color: var(--n-violet); background: var(--n-violet-wash);
    width: 28px; height: 28px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
}
.n-step-title { font-family: 'Fraunces', serif; font-size: 20px; font-weight: 500; color: var(--n-ink); }
.n-step-sub { font-size: 13px; color: var(--n-ink-faint); margin-left: 42px; margin-bottom: 24px; }

/* ═══ UPLOAD CARDS (Styling Streamlit's Native Columns) ═══ */
/* Target columns that contain a file uploader */
[data-testid="column"]:has([data-testid="stFileUploader"]) {
    background: var(--n-panel);
    border: 1px solid var(--n-line);
    border-radius: 12px;
    padding: 16px;
    box-shadow: 0 4px 12px rgba(32,28,44,0.03);
    display: flex;
    flex-direction: column;
}
.n-upload-card-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }
.n-upload-name { font-size: 13px; font-weight: 700; color: var(--n-ink); }
.n-upload-hint { font-size: 11px; color: var(--n-ink-faint); font-family: monospace; margin-bottom: 12px; }

/* Aggressively style the Streamlit Uploader Dropzone */
[data-testid="stFileUploaderDropzone"] {
    background: #FBFAF8 !important;
    border: 1px dashed #DDD7EE !important;
    border-radius: 8px !important;
    padding: 16px 12px !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    cursor: pointer !important;
}
/* Hide Native Streamlit Buttons/Icons inside Dropzone */
[data-testid="stFileUploaderDropzoneInstructions"] { display: none !important; }
[data-testid="stFileUploaderDropzone"] button { display: none !important; }
[data-testid="stFileUploaderDropzone"] svg { display: none !important; }

/* Inject Custom Text */
[data-testid="stFileUploaderDropzone"]::after {
    content: 'Click to attach .xlsx';
    font-size: 11px;
    font-weight: 500;
    color: var(--n-ink-soft);
}
[data-testid="stFileUploaderDropzone"]:has(~ [data-testid="stFileUploaderFile"])::after {
    display: none !important; /* Hide prompt if file uploaded */
}
[data-testid="stFileUploaderFileName"] { font-size: 11px !important; color: var(--n-ink) !important; font-weight: 600 !important; }

.n-upload-status { margin-top: 12px; font-size: 11px; font-weight: 600; color: var(--n-ink-faint); }
.n-upload-status.loaded { color: var(--n-good); }

/* ═══ ACTION ROW ═══ */
.n-action-container {
    background: var(--n-panel); border: 1px solid var(--n-line);
    border-radius: 12px; padding: 24px; margin-top: 24px;
    display: flex; align-items: center; justify-content: space-between;
}
.n-progress-track { height: 8px; background: var(--n-line); border-radius: 4px; overflow: hidden; margin-bottom: 8px; width: 400px; }
.n-progress-fill { height: 100%; background: var(--n-violet); border-radius: 4px; transition: width 0.3s ease; }
.n-progress-text { font-size: 13px; font-weight: 600; color: var(--n-ink-soft); }

[data-testid="stButton"] button {
    background: var(--n-violet) !important;
    color: #FFF !important;
    border: none !important;
    border-radius: 8px !important;
    height: 48px !important;
    font-weight: 600 !important;
    padding: 0 32px !important;
}
[data-testid="stButton"] button:disabled { background: #E4E1EC !important; color: #A6A2B3 !important; }

/* ═══ STAGES / LOGS (Simplified for Space) ═══ */
.n-pipeline-bar { background: var(--n-panel); border: 1px solid var(--n-line); border-radius: 12px; padding: 24px; display: flex; justify-content: space-between; }
.n-stage-col { text-align: center; flex: 1; }
.n-stage-label { font-size: 13px; font-weight: 600; margin-top: 8px; }
.n-stage-sub { font-size: 11px; color: var(--n-ink-faint); }

.n-log-box { background: var(--n-midnight); color: #FFF; border-radius: 12px; padding: 20px; font-family: monospace; font-size: 12px; line-height: 1.8; max-height: 300px; overflow-y: auto; }
.ll-good { color: #5CDDA0; } .ll-fail { color: #FF7A8A; } .ll-dim { color: #888; }
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

def ts(): return time.strftime("%H:%M:%S")
def add_log(text, kind="ok"): st.session_state.log_lines.append((ts(), text, kind))
def set_stage(n, state): st.session_state.stage_states[n] = state

# [Keep your existing write_output_styled & run_pipeline functions here verbatim. 
# They do not affect UI rendering. For brevity in this fix, I am leaving them assumed.]

# ══════════════════════════════════════════════════════════════════════════════
# RENDER UI
# ══════════════════════════════════════════════════════════════════════════════

_html("""
<div class="n-topbar-outer">
  <div class="n-topbar">
    <div class="n-mark">N</div>
    <div class="n-wordmark">Nielsen <span style="font-size:14px; color:#A6A2B3; font-family:Inter; font-style:normal; margin-left:8px;">QC Pipeline</span></div>
  </div>
</div>
""")

_html('<div class="n-shell">')

# ── STEP 1: UPLOAD ────────────────────────────────────────────────────────────
_html("""
<div class="n-step-head"><div class="n-step-num">1</div><div class="n-step-title">Source Files</div></div>
<div class="n-step-sub">Upload all five workbooks to unlock the pipeline run</div>
""")

FILE_DEFS = [
    ("adapt",  "Adapt",      "sheet: programs"),
    ("bsr",    "BSR",        "sheet: Database · row 6"),
    ("yt",     "YouTube MM", "sheet: MM"),
    ("matex",  "Matex",      "sheet: exposures"),
    ("sample", "Sample",     "sheet: programs"),
]

uploaded = {}
loaded_count = 0

# Native Streamlit columns. CSS will style these to look like cards.
upload_cols = st.columns(5)

for col, (key, label, hint) in zip(upload_cols, FILE_DEFS):
    with col:
        uf = st.session_state.get(f"upload_{key}")
        loaded_now = uf is not None
        
        # Header HTML rendered BEFORE the widget
        _html(f"""
        <div class="n-upload-card-head">
          <span class="n-upload-name">{label}</span>
          <span style="color:{'#1F9D6C' if loaded_now else '#E4E1EC'}">●</span>
        </div>
        <div class="n-upload-hint">{hint}</div>
        """)
        
        # The Native Widget
        uf = st.file_uploader(label, type=["xlsx"], key=f"upload_{key}", label_visibility="collapsed")
        uploaded[key] = uf
        
        if uf:
            loaded_count += 1
            status_txt = f"✓ {uf.size/1024:.0f} KB loaded"
        else:
            status_txt = "Awaiting upload"
            
        # Footer HTML rendered AFTER the widget
        _html(f'<div class="n-upload-status{" loaded" if uf else ""}">{status_txt}</div>')

pct = int(loaded_count / 5 * 100)
all_ready = loaded_count == 5
btn_label = "Run QC Pipeline →" if all_ready else f"Waiting for files"

# Native Action Row styling using Streamlit Columns
_html('<div class="n-action-container">')
ac1, ac2 = st.columns([3, 1])
with ac1:
    _html(f"""
    <div>
      <div class="n-progress-track"><div class="n-progress-fill" style="width:{pct}%"></div></div>
      <div class="n-progress-text">{loaded_count} of 5 files ready</div>
    </div>
    """)
with ac2:
    if st.button(btn_label, disabled=not all_ready, use_container_width=True):
        st.write("Pipeline function goes here...") # Replace with your run_pipeline call
_html('</div>')


# ── STEP 2 & 3: STAGES AND LOGS ───────────────────────────────────────────────
_html('<div class="n-step-head"><div class="n-step-num">2</div><div class="n-step-title">Pipeline Log</div></div>')

if not st.session_state.log_lines:
    _html('<div class="n-log-box" style="text-align:center; color:#888;">No pipeline run yet.</div>')
else:
    rows_html = ""
    for stamp, text, kind in st.session_state.log_lines:
        cls = {"ok":"", "good":"ll-good", "fail":"ll-fail", "dim":"ll-dim"}.get(kind, "")
        rows_html += f'<div><span style="color:#666; margin-right:12px;">{stamp}</span><span class="{cls}">{text}</span></div>'
    _html(f'<div class="n-log-box">{rows_html}</div>')

_html('</div>') # Close n-shell
