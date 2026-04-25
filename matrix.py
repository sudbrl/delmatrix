# ══════════════════════════════════════════════════════════════════════════════
# NRB Loan Transition Matrix Dashboard
# Original Heatmap Restored + Robust Excel Parser
# ══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import hashlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from typing import Tuple, Dict

# ── 1. PAGE CONFIGURATION ────────────────────────────────────────────────────
st.set_page_config(
    page_title="NRB Loan Transition Matrix",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 2. CONSTANTS & ORIGINAL HEATMAP COLORS ──────────────────────────────────
GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = len(GRADES)

plt.rcParams.update({"font.family": "DejaVu Sans", "figure.dpi": 140})

# Original Heatmap Color Palette
CANVAS_BG, HEADER_BG, ROW_HDR_BG = "#FAFAF8", "#E8E6DF", "#F1EFE8"
GRID_EDGE,  OUTER_EDGE            = "#D0CCC3", "#BDB7AC"
DIAG_BG,  DIAG_FG  = "#B5D4F4", "#042C53"
UPG_BG,   UPG_FG   = "#EAF3DE", "#173404"
ZERO_BG,  ZERO_FG  = "#F1EFE8", "#888780"
MILD_BG,  MILD_FG  = "#FAEEDA", "#412402"
MOD_BG,   MOD_FG   = "#F0997B", "#4A1B0C"
SEV_BG,   SEV_FG   = "#A32D2D", "#FCEBEB"
TEXT_DARK, TEXT_MID = "#2C2C2A", "#5F5E5A"

# ── 3. CSS STYLING (MODERN UI) ───────────────────────────────────────────────
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
    .stApp { background: var(--bg-primary); color: var(--text-primary); font-family: 'Inter', sans-serif; }
    .metric-card {
        background: var(--bg-secondary); border: 1px solid var(--border);
        border-radius: var(--radius); padding: 16px 20px; box-shadow: var(--shadow);
        transition: transform 0.2s ease;
    }
    .metric-card:hover { transform: translateY(-2px); box-shadow: 0 4px 16px rgba(0,0,0,0.12); }
    .metric-title { color: var(--neutral); font-size: 11px; font-weight: 600; text-transform: uppercase; margin-bottom: 4px; }
    .metric-value { font-size: 24px; font-weight: 700; color: var(--text-primary); line-height: 1.2; }
    .metric-sub { font-size: 12px; color: var(--primary); margin-top: 4px; font-weight: 500; }
    .page-header {
        background: linear-gradient(135deg, var(--bg-secondary), #F5F5F4);
        border: 1px solid var(--border); border-radius: var(--radius);
        padding: 20px 28px; margin-bottom: 24px; box-shadow: var(--shadow);
    }
    .page-header h1 { margin: 0 0 8px 0 !important; font-size: 26px !important; color: var(--text-primary) !important; font-weight: 700; }
    .page-header p { margin: 0; color: var(--neutral); font-size: 14px; }
    .badge { display: inline-flex; align-items: center; padding: 4px 10px; border-radius: 20px;
             font-size: 11px; font-weight: 600; margin-right: 6px; background: #E3F2FD; color: var(--primary-dark); }
    .upload-zone {
        background: var(--bg-secondary); border: 2px dashed var(--border); border-radius: var(--radius);
        padding: 32px 24px; text-align: center; margin: 16px 0; transition: border-color 0.2s ease;
    }
    .upload-zone:hover { border-color: var(--primary); }
    .matrix-table {
        width: 100%; border-collapse: separate; border-spacing: 0; background: var(--bg-secondary);
        border-radius: var(--radius); overflow: hidden; box-shadow: var(--shadow); font-size: 12px;
    }
    .matrix-table th { background: #F5F5F4; color: var(--neutral); padding: 10px 12px; text-align: center; font-weight: 600; border-bottom: 2px solid var(--border); }
    .matrix-table td { padding: 8px 12px; text-align: center; border-bottom: 1px solid var(--border); }
    .matrix-table tr:last-child td { border-bottom: none; }
    .matrix-table .diag { color: var(--primary); font-weight: 700; }
    .matrix-table .upgrade { color: var(--success); }
    .matrix-table .downgrade { color: var(--error); }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; padding: 4px; background: #F5F5F4; border-radius: 8px; }
    .stTabs [data-baseweb="tab"] { padding: 8px 16px; border-radius: 6px; font-weight: 500; color: var(--neutral); }
    .stTabs [aria-selected="true"] { background: var(--bg-secondary); color: var(--primary) !important; box-shadow: 0 2px 4px rgba(0,0,0,0.08); }
    .stButton > button { background: var(--primary); color: white; border: none; border-radius: 8px; font-weight: 600; padding: 10px 24px; transition: background 0.2s ease; }
    .stButton > button:hover { background: var(--primary-dark); }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #C1C1C1; border-radius: 3px; }
    </style>
    """, unsafe_allow_html=True)

load_css()

# ── 4. SESSION STATE & UTILITIES ─────────────────────────────────────────────
def init_session_state():
    defaults = {
        "prev": None, "matrix": None, "period": "",
        "generated": False, "upload_error": None, 
        "filename": None, "stats_cache": None, "last_fig": None
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

# ── 5. ROBUST PARSER (Fixes commas, "- 0", extra columns) ────────────────────
def clean_numeric(val):
    if pd.isna(val): return 0.0
    s = str(val).strip().replace(',', '').replace('-', ' ').replace('–', ' ').replace('—', ' ')
    s = ' '.join(s.split())
    if s.lower() in ['', '0', 'zero', 'nil', 'none']: return 0.0
    try: return float(s)
    except ValueError: return 0.0

@st.cache_data(ttl=3600, show_spinner=False)
def parse_template_cached(file_hash: str, file_bytes: bytes) -> Tuple[np.ndarray, np.ndarray]:
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Failed to read Excel file: {e}. Ensure valid .xlsx format.")
    if df.empty: raise ValueError("The uploaded file is empty.")

    def norm(s): return str(s).strip().lower()
    grade_norms = [norm(g) for g in GRADES]

    header_row_idx = None
    for i in range(min(15, len(df))):
        row_vals = [norm(str(x)) for x in df.iloc[i] if pd.notna(x)]
        if sum(1 for v in row_vals if v in grade_norms) >= 4:
            header_row_idx = i; break
    if header_row_idx is None:
        raise ValueError("Could not find header row with grade names: Good, Watchlist, Substandard, Doubtful, Bad.")

    header_row = df.iloc[header_row_idx]
    col_map = {}
    for ci in range(len(header_row)):
        val = norm(str(header_row.iloc[ci]))
        if val in grade_norms and val not in col_map: col_map[val] = ci
    
    if len(col_map) < N: raise ValueError(f"Missing columns for grades: {[g for g in grade_norms if g not in col_map]}")

    data_rows = {}
    for i in range(header_row_idx + 1, len(df)):
        first_val = norm(str(df.iloc[i, 0]))
        if first_val in grade_norms and first_val not in data_rows: data_rows[first_val] = df.iloc[i]
    if len(data_rows) < N: raise ValueError(f"Missing data rows for grades: {[g for g in grade_norms if g not in data_rows]}")

    trans = np.zeros((N, N), dtype=float)
    for ri, g_from in enumerate(grade_norms):
        row_data = data_rows[g_from]
        for ci, g_to in enumerate(grade_norms):
            col_idx = col_map[g_to]
            trans[ri, ci] = clean_numeric(row_data.iloc[col_idx])

    if trans.sum() == 0:
        raise ValueError("All transition values parsed as zero. Ensure numbers have no commas and zeros are numeric (0, not '- 0').")
    if np.any(trans < 0): raise ValueError("Negative values detected. All amounts must be ≥ 0.")
    return trans.sum(axis=1), trans

@st.cache_data(ttl=1800, show_spinner=False)
def compute_statistics_cached(matrix_tuple: tuple, prev_tuple: tuple) -> Dict:
    trans = np.array(matrix_tuple); prev = np.array(prev_tuple)
    total_closing = float(trans.sum()); col_closing = trans.sum(axis=0)
    retained = float(np.trace(trans))
    upgraded = float(sum(trans[r, c] for r in range(N) for c in range(r)))
    downgraded = float(sum(trans[r, c] for r in range(N) for c in range(r+1, N)))
    return {
        "total_opening": float(prev.sum()), "total_closing": total_closing,
        "col_closing": col_closing.tolist(), "retained": retained,
        "upgraded": upgraded, "downgraded": downgraded,
        "retention_pct": (retained / total_closing * 100) if total_closing else 0,
        "upgrade_pct": (upgraded / total_closing * 100) if total_closing else 0,
        "downgrade_pct": (downgraded / total_closing * 100) if total_closing else 0,
        "migration_rate": (upgraded + downgraded) / total_closing * 100 if total_closing else 0,
        "concentration_ratio": col_closing.max() / total_closing * 100 if total_closing else 0,
        "grade_details": [{
            "grade": GRADES[i], "icon": ICONS[i], "opening": float(prev[i]),
            "closing": float(col_closing[i]), "retained": float(trans[i, i]),
            "upgraded_out": float(sum(trans[i, :i])), "downgraded_out": float(sum(trans[i, i+1:]))
        } for i in range(N)]
    }

# ── 6. ORIGINAL HEATMAP FUNCTIONS (RESTORED) ─────────────────────────────────
def cell_colors(val, ri, ci, prev):
    if ri == ci: return DIAG_BG, DIAG_FG
    if val == 0: return ZERO_BG, ZERO_FG
    if ci < ri: return UPG_BG, UPG_FG
    pct = val / prev[ri] * 100 if prev[ri] > 0 else 0
    if pct < 5: return MILD_BG, MILD_FG
    if pct < 30: return MOD_BG, MOD_FG
    return SEV_BG, SEV_FG

def draw_cell(ax, x, y, w, h, bg, lines, fgs, ec=GRID_EDGE, fs1=9.5, fs2=8.0, w1="bold", w2="normal", ha="center"):
    ax.add_patch(mpatches.Rectangle((x, y), w, h, lw=0.8, edgecolor=ec, facecolor=bg, zorder=2))
    tx = x + (0.10 if ha == "left" else 0.50) * w
    if len(lines) == 1:
        ax.text(tx, y + h / 2, lines[0], ha=ha, va="center", fontsize=fs1, color=fgs[0], fontweight=w1, zorder=3)
    else:
        ax.text(tx, y + h * .64, lines[0], ha=ha, va="center", fontsize=fs1, color=fgs[0], fontweight=w1, zorder=3)
        ax.text(tx, y + h * .29, lines[1], ha=ha, va="center", fontsize=fs2, color=fgs[1], fontweight=w2, zorder=3)

def build_figure(grades, trans, prev, period):
    n = len(grades)
    col_closing = trans.sum(axis=0)
    RH, CW, HW = 0.95, 1.38, 2.45
    th = (n + 2) * RH; tw = HW + n * CW
    fig, ax = plt.subplots(figsize=(tw + .7, th + 1.8))
    fig.patch.set_facecolor(CANVAS_BG); ax.set_facecolor(CANVAS_BG)
    ax.set_aspect("equal"); ax.axis("off")

    ty = (n + 1) * RH
    draw_cell(ax, 0, ty, HW, RH, HEADER_BG, [f"Period: {period}", "Opening  ->  Closing"], [TEXT_DARK, TEXT_MID], ha="left", fs1=9.6, fs2=8.0)
    for ci, g in enumerate(grades):
        draw_cell(ax, HW + ci * CW, ty, CW, RH, HEADER_BG, [g, f"Closing: {col_closing[ci]:,.1f}"], [TEXT_DARK, TEXT_MID], fs1=9.3, fs2=7.8)

    for ri in range(n):
        y = (n - ri) * RH
        draw_cell(ax, 0, y, HW, RH, ROW_HDR_BG, [grades[ri], f"Opening: {prev[ri]:,.1f}"], [TEXT_DARK, TEXT_MID], ha="left")
        for ci in range(n):
            v = trans[ri, ci]
            bg, fg = cell_colors(v, ri, ci, prev)
            p = v / prev[ri] * 100 if prev[ri] > 0 else 0
            if v == 0 and ri != ci:
                draw_cell(ax, HW + ci * CW, y, CW, RH, bg, ["—"], [ZERO_FG], fs1=10, w1="normal")
            else:
                draw_cell(ax, HW + ci * CW, y, CW, RH, bg, [f"{v:,.2f}", f"({p:.1f}%)"], [fg, fg], fs1=9.5, fs2=8.0)

    ax.add_patch(mpatches.Rectangle((0, 0), tw, th, fill=False, lw=1.15, edgecolor=OUTER_EDGE, zorder=4))
    ax.set_xlim(-.08, tw + .08); ax.set_ylim(-.08, th + .1)

    legend_items = [(DIAG_BG, "Retained (diagonal)"), (UPG_BG, "Upgrade"), (MILD_BG, "Mild downgrade <5%"), (MOD_BG, "Moderate 5-30%"), (SEV_BG, "Severe >30%"), (ZERO_BG, "No flow")]
    patches = [mpatches.Patch(facecolor=c, edgecolor=GRID_EDGE, lw=.7, label=l) for c, l in legend_items]
    leg = ax.legend(handles=patches, loc="upper center", bbox_to_anchor=(.5, -.07), ncol=3, fontsize=8.3, frameon=True, fancybox=False, edgecolor=GRID_EDGE, columnspacing=1.3, handlelength=1.5, borderpad=.6)
    leg.get_frame().set_facecolor(CANVAS_BG)
    fig.suptitle("Loan Quality Transition Matrix", fontsize=13, fontweight="bold", y=.98, color=TEXT_DARK)
    plt.tight_layout(rect=(0, .05, 1, .99))
    return fig

def fig_to_bytes(fig, fmt="png", dpi=150):
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, dpi=dpi, bbox_inches="tight", pad_inches=0.1, facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf.read()

def render_matrix_html(trans: np.ndarray, prev: np.ndarray) -> str:
    n = len(GRADES); col_closing = trans.sum(axis=0)
    headers = "".join(f'<th>{g}</th>' for g in GRADES)
    html = [f'<table class="matrix-table"><thead><tr><th style="text-align:left;padding-left:14px;">From / To</th>{headers}<th>Opening</th></tr></thead><tbody>']
    for ri in range(n):
        cells = []
        for ci in range(n):
            v = trans[ri, ci]
            cls = "diag" if ri == ci else ("upgrade" if ci < ri else "downgrade")
            cells.append(f'<td class="{cls}">{v:,.1f}</td>')
        html.append(f'<tr><td style="text-align:left;padding-left:14px;font-weight:600;">{ICONS[ri]} {GRADES[ri]}</td>{"".join(cells)}<td style="color:var(--neutral)">{prev[ri]:,.1f}</td></tr>')
    total_cells = "".join(f'<td style="font-weight:700;color:var(--primary)">{col_closing[ci]:,.1f}</td>' for ci in range(n))
    html.append(f'<tr style="background:#F5F9FF;"><td style="text-align:left;padding-left:14px;font-weight:700;">Closing</td>{total_cells}<td style="font-weight:700;color:var(--primary)">{prev.sum():,.1f}</td></tr></tbody></table>')
    return "".join(html)

# ── 7. SIDEBAR & HEADER ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;"><span style="font-size:32px;">🏦</span><div><div style="font-weight:700;font-size:18px;">NRB Matrix Tool</div><div style="font-size:12px;color:var(--neutral);">v2.2 | Original Heatmap</div></div></div>', unsafe_allow_html=True)
    st.divider()
    st.session_state.period = st.text_input("📅 Reporting Period", value=st.session_state.period, placeholder="e.g., Poush 2081")
    with st.expander("🖼️ Export Settings", expanded=False):
        export_dpi = st.select_slider("Chart DPI", options=[100, 150, 220, 300], value=150)
    st.divider()
    if st.session_state.generated:
        if st.button("🔄 Upload New File", type="secondary", use_container_width=True):
            st.session_state.update({"prev": None, "matrix": None, "generated": False, "stats_cache": None, "last_fig": None})
            st.rerun()
        st.download_button("⬇️ Export JSON", json.dumps({"matrix": st.session_state.matrix.tolist(), "opening": st.session_state.prev.tolist()}, indent=2), "nrb_export.json", "application/json", use_container_width=True)
    st.divider()
    st.markdown('<div style="font-size:12px;color:var(--neutral);line-height:1.6;"><strong>Row Sum</strong> = Opening Balance<br><strong>Col Sum</strong> = Closing Balance<br><strong>Diagonal</strong> = Retained Loans</div>', unsafe_allow_html=True)

st.markdown('<div class="page-header"><h1>🏦 Loan Quality Transition Matrix</h1><p>Upload Excel template → Analyze loan migration → <span style="color:var(--primary);font-weight:600;">Export Insights</span><span class="badge">NRB Compliant</span></p></div>', unsafe_allow_html=True)

tab_upload, tab_dash, tab_analytics, tab_guide = st.tabs(["📂 Upload", "📊 Dashboard", "📈 Analytics", "ℹ️ Guide"])

# ── TAB 1: UPLOAD ────────────────────────────────────────────────────────────
with tab_upload:
    if not st.session_state.generated:
        uploaded = st.file_uploader("Upload Excel Template (.xlsx)", type=["xlsx"], label_visibility="collapsed")
        if uploaded:
            file_bytes = uploaded.read()
            with st.spinner("🔍 Parsing & cleaning data..."):
                try:
                    prev, trans = parse_template_cached(compute_file_hash(file_bytes), file_bytes)
                    st.session_state.update({"prev": prev, "matrix": trans, "filename": uploaded.name, "upload_error": None})
                    st.success(f"✅ Loaded: {uploaded.name} ({int(prev.sum()):,} cr opening)")
                except Exception as e:
                    st.error(f"❌ Parse Error: {e}")
                    st.session_state.update({"prev": None, "matrix": None})

    if st.session_state.prev is not None:
        prev_arr, trans_arr = st.session_state.prev, st.session_state.matrix
        st.divider(); st.markdown("### 📊 Data Preview")
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
                st.markdown(f'<div class="metric-card"><div class="metric-title">{title}</div><div class="metric-value">{val_str}</div><div class="metric-sub" style="color:{color}">{sub}</div></div>', unsafe_allow_html=True)
        st.markdown("#### Transition Matrix Preview")
        st.markdown(render_matrix_html(trans_arr, prev_arr), unsafe_allow_html=True)
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 Generate Dashboard", type="primary", use_container_width=True):
                stats = compute_statistics_cached(tuple(map(tuple, trans_arr)), tuple(prev_arr))
                st.session_state.update({"generated": True, "stats_cache": stats})
                st.rerun()
    else:
        st.markdown('<div class="upload-zone"><div style="font-size:42px;">📁</div><div style="font-weight:600;">Upload Template</div><div style="color:var(--neutral);">Supports .xlsx with Grade Rows/Cols</div></div>', unsafe_allow_html=True)

# ── TAB 2: DASHBOARD (WITH RESTORED HEATMAP) ────────────────────────────────
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
            col.markdown(f'<div class="metric-card"><div class="metric-title">{title}</div><div class="metric-value">{value}</div><div class="metric-sub" style="color:{color}">{sub}</div></div>', unsafe_allow_html=True)
        
        st.divider()
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown("##### Opening vs Closing Distribution")
            st.bar_chart(pd.DataFrame({"Grade": [f"{ICONS[i]} {g}" for i, g in enumerate(GRADES)], "Opening": [stats["grade_details"][i]["opening"] for i in range(N)], "Closing": [stats["grade_details"][i]["closing"] for i in range(N)]}).set_index("Grade"), color=["#64B5F6", "#4DB6AC"], height=300)
        with col_c2:
            st.markdown("##### Retention by Grade")
            st.bar_chart(pd.DataFrame({"Grade": [f"{ICONS[i]} {g}" for i, g in enumerate(GRADES)], "Rate %": [(stats["grade_details"][i]["retained"]/stats["grade_details"][i]["opening"]*100) if stats["grade_details"][i]["opening"]>0 else 0 for i in range(N)]}).set_index("Grade"), color="#81C784", height=300)

        st.divider()
        st.markdown("#### 🔥 Transition Flow Matrix")
        with st.spinner("Rendering professional matrix..."):
            fig = build_figure(GRADES, st.session_state.matrix, st.session_state.prev, st.session_state.period)
            st.session_state["last_fig"] = fig
        st.pyplot(fig, use_container_width=True)
        
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        with col_exp2:
            st.download_button("⬇️ Download PNG", fig_to_bytes(fig, "png", export_dpi), "nrb_matrix.png", "image/png", use_container_width=True)
        with col_exp3:
            st.download_button("⬇️ Download SVG", fig_to_bytes(fig, "svg", export_dpi), "nrb_matrix.svg", "image/svg+xml", use_container_width=True)
        plt.close(fig)

# ── TAB 3: ANALYTICS ────────────────────────────────────────────────────────
with tab_analytics:
    if not st.session_state.generated:
        st.info("👆 Generate a matrix first")
    else:
        st.markdown("#### 📋 Transition Amounts (NPR Crore)")
        main_df = pd.DataFrame(np.round(st.session_state.matrix, 2), index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)], columns=GRADES)
        main_df["Opening"] = np.round(st.session_state.prev, 2)
        closing_row = list(np.round(st.session_state.matrix.sum(axis=0), 2)) + [np.round(st.session_state.matrix.sum(), 2)]
        st.dataframe(pd.concat([main_df, pd.DataFrame([closing_row], index=["Closing"], columns=main_df.columns)]).style.format("{:.2f}"), use_container_width=True)
        
        with st.expander("📊 View Percentages", expanded=False):
            pct_df = pd.DataFrame([[(st.session_state.matrix[r,c]/st.session_state.prev[r]*100) if st.session_state.prev[r]>0 else 0 for c in range(N)] for r in range(N)], index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)], columns=GRADES).round(1)
            st.dataframe(pct_df.style.format("{:.1f}%").background_gradient(cmap="RdYlGn_r", axis=None), use_container_width=True)

        st.markdown("#### 🔍 Grade Details")
        stats = st.session_state.stats_cache
        for i in range(N):
            with st.expander(f"{ICONS[i]} {GRADES[i]} Analysis"):
                d = stats["grade_details"][i]
                m1, m2, m3, m4, m5, m6 = st.columns(6)
                m1.metric("Opening", format_currency(d["opening"]))
                m2.metric("Closing", format_currency(d["closing"]))
                m3.metric("Retained", format_currency(d["retained"]))
                m4.metric("Upgraded", format_currency(d["upgraded_out"]))
                m5.metric("Downgraded", format_currency(d["downgraded_out"]))
                m6.metric("Net Flow", format_currency(d["closing"]-d["opening"]))

# ── TAB 4: GUIDE ────────────────────────────────────────────────────────────
with tab_guide:
    st.markdown("## ℹ️ User Guide")
    st.markdown("""1. **Prepare Template**: Ensure your Excel file has: Row 1 = Headers (`Good`, `Watchlist`, etc.), Column A = Row Labels, 5x5 Data Grid = Numbers only.\n2. **Upload**: Drag & drop your `.xlsx` file in the **Upload** tab.\n3. **Analyze**: Use Dashboard for visuals and Analytics for deep dives.""")
    st.dataframe(pd.DataFrame({"Grade": GRADES, "Status": ["Performing", "Special Mention", "Classified", "Impaired", "Loss"], "Overdue": ["Current", "1-3m", "3-6m", "6-12m", "12m+"]}), use_container_width=True, hide_index=True)

# ── DEPENDENCY CHECK ─────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    st.error("❌ Missing dependency: `openpyxl`. Run `pip install openpyxl`.")
