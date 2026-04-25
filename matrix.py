# ══════════════════════════════════════════════════════════════════════════════
# NRB Loan Transition Matrix Dashboard
# Optimized for Performance & Robust Parsing
# ══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import hashlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from typing import Tuple, Dict, List, Optional

# ── 1. CONFIGURATION & CONSTANTS ─────────────────────────────────────────────

# Page Config
st.set_page_config(
    page_title="NRB Loan Transition Matrix",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Constants
GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = len(GRADES)

# ── 2. CSS STYLING ───────────────────────────────────────────────────────────

def load_css():
    st.markdown("""
    <style>
    :root {
        --primary: #1565C0; --primary-dark: #0D47A1;
        --success: #2E7D32; --warning: #ED6C02; 
        --error: #C62828; --neutral: #7A7670;
        --bg-primary: #FAFAF9; --bg-secondary: #FFFFFF;
        --border: #E0E0E0; --text-primary: #1A1A18;
        --shadow: 0 2px 8px rgba(0,0,0,0.08);
        --radius: 12px;
    }
    
    .stApp { 
        background: var(--bg-primary); 
        color: var(--text-primary);
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .metric-card {
        background: var(--bg-secondary);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 16px 20px;
        box-shadow: var(--shadow);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 16px rgba(0,0,0,0.12);
    }
    .metric-title {
        color: var(--neutral);
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 24px;
        font-weight: 700;
        color: var(--text-primary);
        line-height: 1.2;
    }
    .metric-sub {
        font-size: 12px;
        color: var(--primary);
        margin-top: 4px;
        font-weight: 500;
    }
    
    .page-header {
        background: linear-gradient(135deg, var(--bg-secondary), #F5F5F4);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 20px 28px;
        margin-bottom: 24px;
        box-shadow: var(--shadow);
    }
    .page-header h1 {
        margin: 0 0 8px 0 !important;
        font-size: 26px !important;
        color: var(--text-primary) !important;
        font-weight: 700;
    }
    .page-header p {
        margin: 0;
        color: var(--neutral);
        font-size: 14px;
    }
    
    .badge {
        display: inline-flex;
        align-items: center;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 600;
        margin-right: 6px;
        background: #E3F2FD;
        color: var(--primary-dark);
    }
    .badge.success { background: #E8F5E9; color: var(--success); }
    
    .upload-zone {
        background: var(--bg-secondary);
        border: 2px dashed var(--border);
        border-radius: var(--radius);
        padding: 32px 24px;
        text-align: center;
        margin: 16px 0;
        transition: border-color 0.2s ease;
    }
    .upload-zone:hover {
        border-color: var(--primary);
    }
    
    .matrix-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        background: var(--bg-secondary);
        border-radius: var(--radius);
        overflow: hidden;
        box-shadow: var(--shadow);
        font-size: 12px;
    }
    .matrix-table th {
        background: #F5F5F4;
        color: var(--neutral);
        padding: 10px 12px;
        text-align: center;
        font-weight: 600;
        border-bottom: 2px solid var(--border);
    }
    .matrix-table td {
        padding: 8px 12px;
        text-align: center;
        border-bottom: 1px solid var(--border);
    }
    .matrix-table tr:last-child td { border-bottom: none; }
    .matrix-table .diag { color: var(--primary); font-weight: 700; }
    .matrix-table .upgrade { color: var(--success); }
    .matrix-table .downgrade { color: var(--error); }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        padding: 4px;
        background: #F5F5F4;
        border-radius: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
        border-radius: 6px;
        font-weight: 500;
        color: var(--neutral);
    }
    .stTabs [aria-selected="true"] {
        background: var(--bg-secondary);
        color: var(--primary) !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }
    
    .stButton > button {
        background: var(--primary);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        padding: 10px 24px;
        transition: background 0.2s ease;
    }
    .stButton > button:hover {
        background: var(--primary-dark);
    }
    
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { 
        background: #C1C1C1; 
        border-radius: 3px; 
    }
    
    @media (max-width: 768px) {
        .metric-value { font-size: 20px; }
        .page-header h1 { font-size: 22px !important; }
    }
    </style>
    """, unsafe_allow_html=True)

load_css()

# ── 3. SESSION STATE & UTILITIES ─────────────────────────────────────────────

def init_session_state():
    defaults = {
        "prev": None, "matrix": None, "period": "",
        "generated": False, "upload_error": None, 
        "filename": None, "stats_cache": None,
        "last_fig": None
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_session_state()

def compute_file_hash(file_bytes: bytes) -> str:
    return hashlib.md5(file_bytes).hexdigest()

def format_currency(value: float, suffix: str = " cr") -> str:
    return f"{value:,.1f}{suffix}"

def format_percentage(value: float, decimals: int = 1) -> str:
    return f"{value:.{decimals}f}%"

# ── 4. PARSING LOGIC (ROBUST FIX FOR YOUR DATA) ─────────────────────────────

def clean_numeric(val):
    """
    Cleans common Excel formatting issues:
    - Commas in numbers (3,394.79 -> 3394.79)
    - Text zeros ("- 0", "-- 0", "0 -" -> 0)
    - Empty strings
    """
    if pd.isna(val):
        return 0.0
    
    s = str(val).strip()
    
    # Remove non-numeric noise
    s = s.replace(',', '')       # Remove thousands separator
    s = s.replace('-', ' ')      # Handle stray hyphens
    s = s.replace('–', ' ')      # Handle en-dashes
    s = s.replace('—', ' ')      # Handle em-dashes
    s = ' '.join(s.split())      # Collapse whitespace
    
    # Check for textual zero representations
    if s.lower() in ['', '0', 'zero', 'nil', 'none']:
        return 0.0
    
    try:
        return float(s)
    except ValueError:
        return 0.0  # Fallback

@st.cache_data(ttl=3600, show_spinner=False)
def parse_template_cached(file_hash: str, file_bytes: bytes) -> Tuple[np.ndarray, np.ndarray]:
    """
    Cached Excel parser that handles:
    - Grade name matching
    - Data extraction
    - Numeric cleaning (commas, text zeros)
    - Ignoring extra columns (Grand Total)
    """
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Failed to read Excel file: {e}. Ensure valid .xlsx format.")

    if df.empty:
        raise ValueError("The uploaded file is empty.")

    def norm(s): 
        return str(s).strip().lower()
    
    grade_norms = [norm(g) for g in GRADES]

    # 1. Find Header Row
    # Look for a row containing at least 4 of the grade names
    header_row_idx = None
    for i in range(min(15, len(df))): # Check first 15 rows
        row_vals = [norm(str(x)) for x in df.iloc[i] if pd.notna(x)]
        matches = sum(1 for v in row_vals if v in grade_norms)
        if matches >= 4:
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError(
            "Could not find header row with grade names.\n"
            "Ensure the top row contains: Good, Watchlist, Substandard, Doubtful, Bad."
        )

    # 2. Map Columns to Grades
    header_row = df.iloc[header_row_idx]
    col_map = {}
    for ci in range(len(header_row)):
        val = norm(str(header_row.iloc[ci]))
        if val in grade_norms and val not in col_map:
            col_map[val] = ci
    
    # Validate we found enough columns
    missing_cols = [g for g in grade_norms if g not in col_map]
    if len(missing_cols) > 0:
        raise ValueError(f"Missing columns for grades: {missing_cols}")

    # 3. Extract Data Rows
    data_rows = {}
    # We look for rows where the *first* column (index 0) is a grade name
    # Start search after header row
    for i in range(header_row_idx + 1, len(df)):
        first_val = norm(str(df.iloc[i, 0]))
        if first_val in grade_norms and first_val not in data_rows:
            data_rows[first_val] = df.iloc[i]
            
    missing_rows = [g for g in grade_norms if g not in data_rows]
    if len(missing_rows) > 0:
        raise ValueError(f"Missing data rows for grades: {missing_rows}")

    # 4. Build Matrix
    trans = np.zeros((N, N), dtype=float)
    valid_count = 0
    
    for ri, g_from in enumerate(grade_norms):
        row_data = data_rows[g_from]
        for ci, g_to in enumerate(grade_norms):
            col_idx = col_map[g_to]
            # Read value from specific row/col intersection
            raw_val = row_data.iloc[col_idx]
            
            # CLEAN AND PARSE
            clean_val = clean_numeric(raw_val)
            trans[ri, ci] = clean_val
            if clean_val > 0:
                valid_count += 1

    # 5. Validation
    if trans.sum() == 0:
        raise ValueError(
            "All transition values are zero or unparseable.\n"
            f"Checked {valid_count} non-zero cells.\n"
            "Please ensure:\n"
            "- Numbers do not have commas (use 3394.79, not 3,394.79)\n"
            "- Zeros are numeric (0), not text ('- 0')"
        )

    if np.any(trans < 0):
        raise ValueError("Negative values detected. All amounts must be positive or zero.")

    return trans.sum(axis=1), trans

# ── 5. STATISTICS CACHING ───────────────────────────────────────────────────

@st.cache_data(ttl=1800, show_spinner=False)
def compute_statistics_cached(matrix_tuple: tuple, prev_tuple: tuple) -> Dict:
    trans = np.array(matrix_tuple)
    prev = np.array(prev_tuple)
    
    total_closing = float(trans.sum())
    col_closing = trans.sum(axis=0)
    retained = float(np.trace(trans))
    upgraded = float(sum(trans[r, c] for r in range(N) for c in range(r)))
    downgraded = float(sum(trans[r, c] for r in range(N) for c in range(r+1, N)))
    
    # Metrics
    retention_pct = (retained / total_closing * 100) if total_closing else 0
    upgrade_pct = (upgraded / total_closing * 100) if total_closing else 0
    downgrade_pct = (downgraded / total_closing * 100) if total_closing else 0
    migration_rate = upgrade_pct + downgrade_pct
    concentration = (col_closing.max() / total_closing * 100) if total_closing else 0
    
    retention_rates = [(trans[i,i] / prev[i] * 100) if prev[i] > 0 else 0 for i in range(N)]
    
    return {
        "total_opening": float(prev.sum()),
        "total_closing": total_closing,
        "retained": retained,
        "upgraded": upgraded,
        "downgraded": downgraded,
        "retention_pct": retention_pct,
        "upgrade_pct": upgrade_pct,
        "downgrade_pct": downgrade_pct,
        "migration_rate": migration_rate,
        "concentration_ratio": concentration,
        "col_closing": col_closing.tolist(),
        "retention_rates": retention_rates,
        "grade_details": [
            {
                "grade": GRADES[i], "icon": ICONS[i],
                "opening": float(prev[i]), "closing": float(col_closing[i]),
                "retained": float(trans[i, i]),
                "upgraded_out": float(sum(trans[i, :i])),
                "downgraded_out": float(sum(trans[i, i+1:])),
                "net_change": float(col_closing[i] - prev[i])
            } for i in range(N)
        ]
    }

# ── 6. RENDERERS & VISUALIZATIONS ───────────────────────────────────────────

def render_matrix_html(trans: np.ndarray, prev: np.ndarray) -> str:
    n = len(GRADES)
    col_closing = trans.sum(axis=0)
    
    headers = "".join(f'<th>{g}</th>' for g in GRADES)
    html_parts = [f'<table class="matrix-table"><thead><tr><th style="text-align:left;padding-left:14px;">From / To</th>{headers}<th>Opening</th></tr></thead><tbody>']
    
    for ri in range(n):
        cells = []
        for ci in range(n):
            v = trans[ri, ci]
            # Classification logic for coloring
            if ri == ci: cls = "diag"
            elif ci < ri: cls = "upgrade" 
            elif v == 0: cls = "zero" # Not explicitly colored in CSS yet but good practice
            else: 
                pct = (v/prev[ri]*100) if prev[ri]>0 else 0
                cls = "downgrade" if pct >= 5 else "downgrade" # Use downgrade style
            
            cells.append(f'<td class="{cls}">{v:,.1f}</td>')
        
        row_sum = prev[ri]
        html_parts.append(
            f'<tr><td style="text-align:left;padding-left:14px;font-weight:600;">'
            f'{ICONS[ri]} {GRADES[ri]}</td>'
            f'{"".join(cells)}'
            f'<td style="color:var(--neutral)">{row_sum:,.1f}</td></tr>'
        )
    
    total_cells = "".join(f'<td style="font-weight:700;color:var(--primary)">{col_closing[ci]:,.1f}</td>' for ci in range(n))
    html_parts.append(
        f'<tr style="background:#F5F9FF;"><td style="text-align:left;padding-left:14px;font-weight:700;">Closing</td>'
        f'{total_cells}'
        f'<td style="font-weight:700;color:var(--primary)">{prev.sum():,.1f}</td></tr>'
    )
    html_parts.append("</tbody></table>")
    return "".join(html_parts)

def build_figure_optimized(grades, trans, prev, period, figsize=(14, 8)):
    import matplotlib.pyplot as plt
    n = len(grades)
    col_closing = trans.sum(axis=0)
    
    RH, CW, HW = 0.95, 1.38, 2.45
    th = (n + 2) * RH
    tw = HW + n * CW
    
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    fig.patch.set_facecolor("#FAFAF9")
    ax.set_facecolor("#FAFAF9")
    ax.set_aspect("equal")
    ax.axis("off")
    
    # Colors
    COLORS = {
        "header": "#F5F5F4", "row_hdr": "#F1EFE8", "grid": "#E0E0E0",
        "diag": ("#B5D4F4", "#0D47A1"), "upgrade": ("#EAF3DE", "#1B5E20"),
        "zero": ("#F1EFE8", "#757575"), "mild": ("#FAEEDA", "#412402"),
        "mod": ("#F0997B", "#4A1B0C"), "severe": ("#A32D2D", "#FFFFFF"),
        "text": ("#1A1A18", "#757575")
    }
    
    def draw_cell(x, y, w, h, bg, lines, fgs, align="center"):
        ax.add_patch(mpatches.Rectangle((x, y), w, h, lw=0.5, edgecolor=COLORS["grid"], facecolor=bg, zorder=2))
        tx = x + (0.10 if align=="left" else 0.50) * w
        ax.text(tx, y+h*0.64, lines[0], ha=align, va="center", fontsize=9, color=fgs[0], fontweight="bold", zorder=3)
        if len(lines) > 1:
            ax.text(tx, y+h*0.29, lines[1], ha=align, va="center", fontsize=7.5, color=fgs[1], zorder=3)

    # Header Row
    ty = (n+1) * RH
    draw_cell(0, ty, HW, RH, COLORS["header"], [f"Period: {period}", "Opening → Closing"], COLORS["text"], align="left")
    for ci, g in enumerate(grades):
        draw_cell(HW+ci*CW, ty, CW, RH, COLORS["header"], [g, f"Close: {col_closing[ci]:.0f}"], COLORS["text"])
    
    # Matrix
    for ri in range(n):
        y = (n-ri) * RH
        draw_cell(0, y, HW, RH, COLORS["row_hdr"], [grades[ri], f"Open: {prev[ri]:.0f}"], COLORS["text"], align="left")
        for ci in range(n):
            v = trans[ri, ci]
            if ri == ci: bg, fg = COLORS["diag"]
            elif v == 0: bg, fg = COLORS["zero"]
            elif ci < ri: bg, fg = COLORS["upgrade"]
            else:
                pct = (v/prev[ri]*100) if prev[ri]>0 else 0
                if pct < 5: bg, fg = COLORS["mild"]
                elif pct < 30: bg, fg = COLORS["mod"]
                else: bg, fg = COLORS["severe"]
            
            p = (v/prev[ri]*100) if prev[ri]>0 else 0
            draw_cell(HW+ci*CW, y, CW, RH, bg, [f"{v:.1f}", f"({p:.0f}%)"], [fg, fg])

    # Legend
    ax.add_patch(mpatches.Rectangle((0, 0), tw, th, fill=False, lw=1, edgecolor="#BDBDBD", zorder=4))
    legend_items = [
        (COLORS["diag"][0], "Retained"), (COLORS["upgrade"][0], "Upgrade"),
        (COLORS["mild"][0], "Mild ↓"), (COLORS["mod"][0], "Mod ↓"),
        (COLORS["severe"][0], "Severe ↓"), (COLORS["zero"][0], "Zero")
    ]
    patches = [mpatches.Patch(facecolor=c, edgecolor=COLORS["grid"], lw=0.5, label=l) for c, l in legend_items]
    ax.legend(handles=patches, loc="upper center", bbox_to_anchor=(0.5, -0.08), ncol=3, fontsize=7.5, frameon=False)
    
    fig.suptitle("Loan Quality Transition Matrix", fontsize=14, fontweight="bold", y=0.98)
    plt.tight_layout(rect=(0, 0.03, 1, 0.99))
    return fig

def fig_to_bytes(fig, fmt="png", dpi=150):
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, dpi=dpi, bbox_inches="tight", pad_inches=0.1, facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf.read()

# ── 7. APP LAYOUT & LOGIC ───────────────────────────────────────────────────

# Sidebar
with st.sidebar:
    st.markdown("""
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;">
        <span style="font-size:32px;">🏦</span>
        <div>
            <div style="font-weight:700;font-size:18px;">NRB Matrix Tool</div>
            <div style="font-size:12px;color:var(--neutral);">v2.1 | Optimized</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    st.session_state.period = st.text_input("📅 Reporting Period", value=st.session_state.period, placeholder="e.g., Poush 2081")
    
    with st.expander("🖼️ Export Settings", expanded=False):
        export_dpi = st.select_slider("Chart DPI", options=[100, 150, 220, 300], value=150)
    
    st.divider()
    
    if st.session_state.generated:
        if st.button("🔄 Upload New File", type="secondary", use_container_width=True):
            st.session_state.update({"prev": None, "matrix": None, "generated": False, "stats_cache": None, "last_fig": None})
            st.rerun()
            
        st.download_button(
            "⬇️ Export JSON",
            json.dumps({"matrix": st.session_state.matrix.tolist(), "opening": st.session_state.prev.tolist()}, indent=2),
            "nrb_export.json", "application/json", use_container_width=True
        )

    st.divider()
    with st.expander("📖 Matrix Logic"):
        st.markdown("""
        **Row Sum** = Opening Balance<br>
        **Col Sum** = Closing Balance<br>
        **Diagonal** = Retained Loans
        """, unsafe_allow_html=True)

# Header
st.markdown("""
<div class="page-header">
    <h1>🏦 Loan Quality Transition Matrix</h1>
    <p>
        Upload Excel template → Analyze loan migration → 
        <span style="color:var(--primary);font-weight:600;">Export Insights</span>
        <span class="badge">NRB Compliant</span>
    </p>
</div>
""", unsafe_allow_html=True)

# Tabs
tab_upload, tab_dash, tab_analytics, tab_guide = st.tabs(["📂 Upload", "📊 Dashboard", "📈 Analytics", "ℹ️ Guide"])

# ── TAB 1: UPLOAD ────────────────────────────────────────────────────────────
with tab_upload:
    if not st.session_state.generated:
        uploaded = st.file_uploader("Upload Excel Template (.xlsx)", type=["xlsx"], label_visibility="collapsed")
        
        if uploaded:
            file_bytes = uploaded.read()
            with st.spinner("🔍 Parsing template (cleaning data)..."):
                try:
                    prev, trans = parse_template_cached(compute_file_hash(file_bytes), file_bytes)
                    st.session_state.update({"prev": prev, "matrix": trans, "filename": uploaded.name, "upload_error": None})
                    st.success(f"✅ Loaded: {uploaded.name} ({int(prev.sum()):,} cr opening)")
                except Exception as e:
                    st.error(f"❌ Parse Error: {e}")
                    st.session_state.update({"prev": None, "matrix": None})

    if st.session_state.prev is not None:
        prev_arr, trans_arr = st.session_state.prev, st.session_state.matrix
        
        st.divider()
        st.markdown("### 📊 Data Preview")
        
        # KPI Cards
        cols = st.columns(4)
        data = [
            (cols[0], "Opening", prev_arr.sum(), "Total", "var(--primary)"),
            (cols[1], "Closing", trans_arr.sum(), "Total", "var(--primary)"),
            (cols[2], "Retention", sum(trans_arr[i,i] for i in range(N))/trans_arr.sum()*100, "%", "var(--success)"),
            (cols[3], "Status", abs(trans_arr.sum()-prev_arr.sum()), "Delta", "var(--warning)" if abs(trans_arr.sum()-prev_arr.sum())>0.1 else "var(--success)")
        ]
        for col, title, val, sub, color in data:
            with col:
                val_str = format_currency(val) if 'Status' not in title else f"{val:.2f}"
                st.markdown(f"""<div class="metric-card"><div class="metric-title">{title}</div><div class="metric-value">{val_str}</div><div class="metric-sub" style="color:{color}">{sub}</div></div>""", unsafe_allow_html=True)

        st.markdown("#### Transition Matrix")
        st.markdown(render_matrix_html(trans_arr, prev_arr), unsafe_allow_html=True)
        
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 Generate Dashboard", type="primary", use_container_width=True):
                stats = compute_statistics_cached(tuple(map(tuple, trans_arr)), tuple(prev_arr))
                st.session_state.update({"generated": True, "stats_cache": stats})
                st.rerun()
    else:
        st.markdown('<div class="upload-zone"><div style="font-size:42px;">📁</div><div style="font-weight:600;">Upload Template</div><div style="color:var(--neutral);">Supports .xlsx with Grade Rows/Cols</div></div>', unsafe_allow_html=True)

# ── TAB 2: DASHBOARD ────────────────────────────────────────────────────────
with tab_dash:
    if not st.session_state.generated:
        st.info("👆 Upload and generate a matrix first")
    else:
        stats = st.session_state.stats_cache or compute_statistics_cached(tuple(map(tuple, st.session_state.matrix)), tuple(st.session_state.prev))
        
        st.markdown("#### 🎯 Key Performance Indicators")
        kpi_cols = st.columns(4)
        kpi_data = [
            ("Total Portfolio", format_currency(stats["total_closing"]), "Value at Risk", "var(--primary)"),
            ("Retention Rate", format_percentage(stats["retention_pct"]), f"{format_currency(stats['retained'])} Stable", "var(--success)"),
            ("Migration Rate", format_percentage(stats["migration_rate"]), f"{format_percentage(stats['upgrade_pct'])} Upgraded", "var(--warning)"),
            ("Concentration", format_percentage(stats["concentration_ratio"]), "Highest Grade", "var(--neutral)")
        ]
        for col, (title, value, sub, color) in zip(kpi_cols, kpi_data):
            col.markdown(f"""<div class="metric-card"><div class="metric-title">{title}</div><div class="metric-value">{value}</div><div class="metric-sub" style="color:{color}">{sub}</div></div>""", unsafe_allow_html=True)
        
        st.divider()
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown("##### Opening vs Closing")
            st.bar_chart(pd.DataFrame({
                "Grade": [f"{ICONS[i]} {g}" for i, g in enumerate(GRADES)],
                "Opening": [stats["grade_details"][i]["opening"] for i in range(N)],
                "Closing": [stats["grade_details"][i]["closing"] for i in range(N)]
            }).set_index("Grade"), color=["#64B5F6", "#4DB6AC"], height=300)
            
        with col_c2:
            st.markdown("##### Retention by Grade")
            st.bar_chart(pd.DataFrame({
                "Grade": [f"{ICONS[i]} {g}" for i, g in enumerate(GRADES)],
                "Rate %": stats["retention_rates"]
            }).set_index("Grade"), color="#81C784", height=300)

        st.divider()
        st.markdown("#### 🔥 Transition Heatmap")
        heatmap_df = pd.DataFrame(st.session_state.matrix, index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)], columns=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)])
        st.dataframe(heatmap_df.style.format("{:.1f}").background_gradient(cmap="RdYlGn_r", axis=None), use_container_width=True, height=350)
        
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        with col_exp1:
            if st.button("🖼️ Generate Chart", use_container_width=True):
                with st.spinner("Rendering..."):
                    fig = build_figure_optimized(GRADES, st.session_state.matrix, st.session_state.prev, st.session_state.period)
                    st.session_state["last_fig"] = fig
                    st.pyplot(fig, use_container_width=True)
        
        if st.session_state.get("last_fig"):
            col_exp2.download_button("⬇️ PNG", fig_to_bytes(st.session_state["last_fig"], "png", export_dpi), "matrix.png", "image/png", use_container_width=True)
            col_exp3.download_button("⬇️ SVG", fig_to_bytes(st.session_state["last_fig"], "svg", export_dpi), "matrix.svg", "image/svg+xml", use_container_width=True)

# ── TAB 3: ANALYTICS ────────────────────────────────────────────────────────
with tab_analytics:
    if not st.session_state.generated:
        st.info("👆 Generate a matrix first")
    else:
        stats = st.session_state.stats_cache
        st.markdown("#### 📋 Transition Matrix (NPR Crore)")
        main_df = pd.DataFrame(np.round(st.session_state.matrix, 2), index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)], columns=GRADES)
        main_df["Opening"] = np.round(st.session_state.prev, 2)
        closing_row = list(np.round(st.session_state.matrix.sum(axis=0), 2)) + [np.round(st.session_state.matrix.sum(), 2)]
        st.dataframe(pd.concat([main_df, pd.DataFrame([closing_row], index=["Closing"], columns=main_df.columns)]).style.format("{:.2f}"), use_container_width=True)
        
        with st.expander("📊 View Percentages", expanded=False):
            pct_df = pd.DataFrame([[(st.session_state.matrix[r,c]/st.session_state.prev[r]*100) if st.session_state.prev[r]>0 else 0 for c in range(N)] for r in range(N)], index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)], columns=GRADES).round(1)
            st.dataframe(pct_df.style.format("{:.1f}%").background_gradient(cmap="RdYlGn_r", axis=None), use_container_width=True)

        st.markdown("#### 🔍 Grade Details")
        for i in range(N):
            with st.expander(f"{ICONS[i]} {GRADES[i]} Analysis"):
                d = stats["grade_details"][i]
                m1, m2, m3, m4, m5, m6 = st.columns(6)
                m1.metric("Opening", format_currency(d["opening"]))
                m2.metric("Closing", format_currency(d["closing"]))
                m3.metric("Retained", format_currency(d["retained"]))
                m4.metric("Upgraded", format_currency(d["upgraded_out"]))
                m5.metric("Downgraded", format_currency(d["downgraded_out"]))
                m6.metric("Net Flow", format_currency(d["net_change"]), delta=None)
                st.progress(min(100, int(stats["retention_rates"][i])), text=f"Retention: {stats['retention_rates'][i]:.1f}%")

# ── TAB 4: GUIDE ────────────────────────────────────────────────────────────
with tab_guide:
    st.markdown("## ℹ️ User Guide")
    st.markdown("""
    1. **Prepare Template**: Ensure your Excel file has:
       - Row 1: Headers (`Good`, `Watchlist`, etc.)
       - Column A: Row Labels (`Good`, `Watchlist`, etc.)
       - 5x5 Data Grid: Numbers only (No commas, no text).
    2. **Upload**: Drag & drop your `.xlsx` file in the **Upload** tab.
    3. **Analyze**: Use Dashboard for visuals and Analytics for deep dives.
    """)
    
    st.dataframe(pd.DataFrame({
        "Grade": GRADES, "Status": ["Performing", "Special Mention", "Classified", "Impaired", "Loss"],
        "Overdue": ["Current", "1-3m", "3-6m", "6-12m", "12m+"]
    }), use_container_width=True, hide_index=True)

# Dependency Check
try:
    import openpyxl
except ImportError:
    st.error("❌ Missing dependency: `openpyxl`. Run `pip install openpyxl`.")
