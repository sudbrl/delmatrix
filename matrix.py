# ── Optimized NRB Loan Transition Matrix Dashboard ───────────────────────────
import numpy as np
import pandas as pd
import streamlit as st
import io
import json
import hashlib
from typing import Tuple, Dict, List, Optional

# ── Page Configuration ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="NRB Loan Transition Matrix",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://nrb.org.np',
        'Report a bug': "mailto:support@nrb.org.np",
        'About': "## NRB Loan Transition Matrix Tool\nBuilt for Nepal Rastra Bank reporting standards."
    }
)

# ── Constants & Configuration ────────────────────────────────────────────────
GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = len(GRADES)

# ── Session State Initialization ─────────────────────────────────────────────
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

# ── Caching Layer for Performance ────────────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner=False)
def parse_template_cached(file_hash: str, file_bytes: bytes) -> Tuple[np.ndarray, np.ndarray]:
    """Cached Excel parser with hash-based invalidation."""
    try:
        import openpyxl
    except ImportError:
        raise ImportError("Missing dependency: openpyxl. Run: pip install openpyxl")
    
    df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")
    
    if df.empty:
        raise ValueError("The uploaded file appears to be empty.")

    def norm(s): 
        return str(s).strip().lower()
    
    grade_norms = [norm(g) for g in GRADES]

    # Vectorized header detection
    header_candidates = df.apply(
        lambda row: row.astype(str).str.strip().str.lower().isin(grade_norms).sum(), 
        axis=1
    )
    header_row_idx = None
    if (header_candidates >= 4).any():
        header_row_idx = header_candidates[header_candidates >= 4].index[0]

    if header_row_idx is None:
        raise ValueError("Could not find header row with grade names: Good, Watchlist, Substandard, Doubtful, Bad")

    # Build column mapping efficiently
    header_row = df.iloc[header_row_idx].astype(str).str.strip().str.lower()
    col_map = {}
    for i, gn in enumerate(grade_norms):
        matches = header_row[header_row == gn]
        if len(matches) > 0:
            col_map[gn] = matches.index[0]
    
    if len(col_map) < N:
        missing = [g for g in grade_norms if g not in col_map]
        raise ValueError(f"Missing grades in header: {missing}")

    # Extract data rows with vectorized operations
    first_col = df.iloc[:, 0].astype(str).str.strip().str.lower()
    data_mask = first_col.isin(grade_norms)
    data_rows_df = df[data_mask].copy()
    if not data_rows_df.empty:
        data_rows_df = data_rows_df.set_index(data_rows_df.columns[0])
    
    # Build transition matrix
    trans = np.zeros((N, N), dtype=float)
    for ri, g_from in enumerate(grade_norms):
        if g_from in data_rows_df.index:
            row_data = data_rows_df.loc[g_from]
            for ci, g_to in enumerate(grade_norms):
                if g_to in col_map:
                    col_idx = col_map[g_to]
                    if col_idx < len(row_data):
                        val = row_data.iloc[col_idx]
                        try:
                            trans[ri, ci] = float(val) if pd.notna(val) else 0.0
                        except (ValueError, TypeError):
                            trans[ri, ci] = 0.0

    # Validation
    if trans.sum() == 0:
        raise ValueError("All transition values are zero. Check your data cells.")
    if np.any(trans < 0):
        raise ValueError("Negative values detected. All amounts must be non-negative.")

    return trans.sum(axis=1), trans


@st.cache_data(ttl=1800, show_spinner=False)
def compute_statistics_cached(matrix_tuple: tuple, prev_tuple: tuple) -> Dict:
    """Cached statistics computation with tuple inputs for hashability."""
    trans = np.array(matrix_tuple)
    prev = np.array(prev_tuple)
    
    total_opening = float(prev.sum())
    total_closing = float(trans.sum())
    col_closing = trans.sum(axis=0)
    
    # Core metrics
    retained = float(np.trace(trans))
    upgraded = float(sum(trans[r, c] for r in range(N) for c in range(r)))
    downgraded = float(sum(trans[r, c] for r in range(N) for c in range(r+1, N)))
    
    # Advanced risk metrics
    migration_rate = (upgraded + downgraded) / total_closing * 100 if total_closing else 0
    concentration_ratio = col_closing.max() / total_closing * 100 if total_closing else 0
    
    # Grade-level retention rates
    retention_rates = np.array([
        trans[i, i] / prev[i] * 100 if prev[i] > 0 else 0 
        for i in range(N)
    ])
    
    # Net flow analysis
    net_flows = col_closing - prev
    
    return {
        "total_opening": total_opening,
        "total_closing": total_closing,
        "col_closing": col_closing.tolist(),
        "retained": retained,
        "upgraded": upgraded,
        "downgraded": downgraded,
        "retention_pct": retained / total_closing * 100 if total_closing else 0,
        "upgrade_pct": upgraded / total_closing * 100 if total_closing else 0,
        "downgrade_pct": downgraded / total_closing * 100 if total_closing else 0,
        "migration_rate": migration_rate,
        "concentration_ratio": concentration_ratio,
        "retention_rates": retention_rates.tolist(),
        "net_flows": net_flows.tolist(),
        "grade_details": [
            {
                "grade": GRADES[i],
                "icon": ICONS[i],
                "opening": float(prev[i]),
                "closing": float(col_closing[i]),
                "retained": float(trans[i, i]),
                "upgraded_out": float(sum(trans[i, :i])),
                "downgraded_out": float(sum(trans[i, i+1:])),
                "inflow": float(sum(trans[r, i] for r in range(N) if r != i)),
                "retention_rate": float(retention_rates[i]),
                "net_change": float(net_flows[i])
            }
            for i in range(N)
        ]
    }


# ── Helper Functions ─────────────────────────────────────────────────────────
def compute_file_hash(file_bytes: bytes) -> str:
    """Compute simple hash for cache invalidation."""
    return hashlib.md5(file_bytes).hexdigest()


def format_currency(value: float, suffix: str = " cr") -> str:
    """Format currency values consistently."""
    return f"{value:,.1f}{suffix}"


def format_percentage(value: float, decimals: int = 1) -> str:
    """Format percentage values."""
    return f"{value:.{decimals}f}%"


# ── CSS Styling ──────────────────────────────────────────────────────────────
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
    .badge.warning { background: #FFF3E0; color: var(--warning); }
    
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
    ::-webkit-scrollbar-thumb:hover { background: #A0A0A0; }
    
    @media (max-width: 768px) {
        .metric-value { font-size: 20px; }
        .page-header h1 { font-size: 22px !important; }
    }
    </style>
    """, unsafe_allow_html=True)

load_css()

# ── HTML Renderer Function (defined before use) ───────────────────────────────
def render_matrix_html(trans: np.ndarray, prev: np.ndarray) -> str:
    """Optimized HTML table renderer with CSS classes."""
    n = len(GRADES)
    col_closing = trans.sum(axis=0)
    
    # Build header
    headers = "".join(f'<th>{g}</th>' for g in GRADES)
    html_parts = [f'<table class="matrix-table"><thead><tr><th style="text-align:left;padding-left:14px;">From / To</th>{headers}<th>Opening</th></tr></thead><tbody>']
    
    # Build rows
    for ri in range(n):
        cells = []
        for ci in range(n):
            v = trans[ri, ci]
            if ri == ci: 
                cls = "diag"
            elif ci < ri: 
                cls = "upgrade" 
            elif v == 0: 
                cls = "zero"
            else:
                pct = v/prev[ri]*100 if prev[ri]>0 else 0
                cls = "downgrade" if pct >= 5 else ""
            cells.append(f'<td class="{cls}">{v:,.1f}</td>')
        
        row_sum = prev[ri]
        html_parts.append(
            f'<tr><td style="text-align:left;padding-left:14px;font-weight:600;">'
            f'{ICONS[ri]} {GRADES[ri]}</td>'
            f'{"".join(cells)}'
            f'<td style="color:var(--neutral)">{row_sum:,.1f}</td></tr>'
        )
    
    # Closing row
    total_cells = "".join(
        f'<td style="font-weight:700;color:var(--primary)">{col_closing[ci]:,.1f}</td>' 
        for ci in range(n)
    )
    html_parts.append(
        f'<tr style="background:#F5F9FF;"><td style="text-align:left;padding-left:14px;font-weight:700;">Closing</td>'
        f'{total_cells}'
        f'<td style="font-weight:700;color:var(--primary)">{prev.sum():,.1f}</td></tr>'
    )
    
    html_parts.append("</tbody></table>")
    return "".join(html_parts)


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;">
        <span style="font-size:32px;">🏦</span>
        <div>
            <div style="font-weight:700;font-size:18px;color:var(--text-primary);">
                NRB Matrix Tool
            </div>
            <div style="font-size:12px;color:var(--neutral);">
                Loan Quality Analytics
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    # Period Input
    st.session_state.period = st.text_input(
        "📅 Reporting Period",
        value=st.session_state.period,
        placeholder="e.g., Poush 2081",
        help="Label for the reporting period"
    )
    
    # Export Settings
    with st.expander("🖼️ Export Settings", expanded=False):
        export_dpi = st.select_slider(
            "Chart Resolution", 
            options=[100, 150, 220, 300], 
            value=220,
            help="Higher DPI = better quality but larger file size"
        )
    
    st.divider()
    
    # Reset Option
    if st.session_state.generated:
        if st.button("🔄 Upload New File", use_container_width=True, type="secondary"):
            for key in ["prev", "matrix", "generated", "upload_error", "filename", "stats_cache", "last_fig"]:
                st.session_state[key] = None if key != "period" else st.session_state.period
            st.rerun()
    
    st.divider()
    
    # Quick Export
    if st.session_state.generated and st.session_state.matrix is not None:
        st.markdown("### 📤 Quick Export")
        exp_data = {
            "metadata": {
                "period": st.session_state.period,
                "grades": GRADES,
                "generated_at": pd.Timestamp.now().isoformat()
            },
            "data": {
                "opening": st.session_state.prev.tolist(),
                "transition_matrix": st.session_state.matrix.tolist(),
                "closing": st.session_state.matrix.sum(axis=0).tolist()
            }
        }
        st.download_button(
            "⬇️ Export JSON",
            json.dumps(exp_data, indent=2),
            f"nrb_matrix_{st.session_state.period or 'export'}.json",
            "application/json",
            use_container_width=True
        )
    
    st.divider()
    
    # Legend
    with st.expander("📖 Matrix Legend"):
        st.markdown("""
        <div style="font-size:12px;line-height:1.6;color:var(--neutral);">
            <strong>Row Sum</strong> = Opening balance<br>
            <strong>Column Sum</strong> = Closing balance<br>
            <strong>Diagonal</strong> = Retained loans<br>
            <strong>↑ Green</strong> = Upgrade flows<br>
            <strong>↓ Red</strong> = Downgrade flows
        </div>
        """, unsafe_allow_html=True)

# ── Main Header ──────────────────────────────────────────────────────────────
st.markdown("""
<div class="page-header">
    <h1>🏦 Loan Quality Transition Matrix</h1>
    <p>
        Upload your Excel template → Analyze loan migration patterns → 
        <span style="color:var(--primary);font-weight:600;">Export insights</span>
        <span class="badge">NRB Compliant</span>
        <span class="badge success">v2.1</span>
    </p>
</div>
""", unsafe_allow_html=True)

# ── Tab Navigation ───────────────────────────────────────────────────────────
tabs = st.tabs(["📂 Upload", "📊 Dashboard", "📈 Analytics", "ℹ️ Guide"])

# ═════════════════════════════════════════════════════════════════════════════
# TAB 1: UPLOAD & PREVIEW
# ═════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    # Template Guide
    with st.expander("📋 Expected Template Format", expanded=True):
        st.markdown("""
        <div style="background:#F5F9FF;border-left:4px solid var(--primary);
                    padding:12px 16px;border-radius:0 8px 8px 0;font-size:13px;">
            <strong>Required Structure:</strong>
            <ul style="margin:8px 0 0 20px;padding:0;">
                <li>Header row with grades: <code>Good | Watchlist | Substandard | Doubtful | Bad</code></li>
                <li>5 data rows starting with grade names</li>
                <li>Cell values = amount flowing from row grade → column grade</li>
                <li><em>Optional:</em> Grand Total column (auto-ignored)</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        # Visual example
        example_df = pd.DataFrame(
            [[3394.8, 363.7, 12.6, 0, 0],
             [230.9, 425.4, 84.3, 0, 0],
             [24.6, 20.3, 23.8, 53.0, 0],
             [2.8, 2.3, 2.9, 44.1, 10.5],
             [20.6, 0.4, 0.1, 0.3, 247.8]],
            index=GRADES,
            columns=GRADES
        )
        st.dataframe(
            example_df.style.format("{:.1f}").background_gradient(
                cmap="RdYlGn_r", subset=example_df.columns, axis=None
            ),
            use_container_width=True,
            height=200
        )

    st.markdown("")
    
    # File Upload
    if not st.session_state.generated:
        uploaded = st.file_uploader(
            "Upload Excel Template (.xlsx)",
            type=["xlsx"],
            help="File must contain 5×5 transition matrix with NRB grade labels",
            label_visibility="collapsed"
        )

        if uploaded:
            file_bytes = uploaded.read()
            file_hash = compute_file_hash(file_bytes)
            
            with st.spinner("🔍 Parsing template..."):
                try:
                    prev, trans = parse_template_cached(file_hash, file_bytes)
                    st.session_state.update({
                        "prev": prev,
                        "matrix": trans,
                        "filename": uploaded.name,
                        "upload_error": None
                    })
                    st.success(f"✅ Loaded: {uploaded.name} ({int(prev.sum()):,} cr opening)")
                except Exception as e:
                    st.session_state.update({
                        "prev": None, "matrix": None, 
                        "upload_error": str(e)
                    })
                    st.error(f"❌ Parse error: {e}")

    # Data Preview (if loaded)
    if st.session_state.prev is not None:
        prev_arr = st.session_state.prev
        trans_arr = st.session_state.matrix
        
        st.divider()
        st.markdown("### 📊 Data Preview")
        
        # KPI Cards
        cols = st.columns(4)
        kpis = [
            (cols[0], "Opening Balance", prev_arr.sum(), "Total row sum", "var(--primary)"),
            (cols[1], "Closing Balance", trans_arr.sum(), "Total column sum", "var(--primary)"),
            (cols[2], "Retention Rate", 
             sum(trans_arr[i,i] for i in range(N))/trans_arr.sum()*100 if trans_arr.sum() else 0,
             "% on diagonal", "var(--success)"),
            (cols[3], "Balance Check", 
             abs(trans_arr.sum() - prev_arr.sum()),
             "Should be ~0", "var(--warning)" if abs(trans_arr.sum()-prev_arr.sum())>0.1 else "var(--success)")
        ]
        
        for col, title, value, subtitle, color in kpis:
            with col:
                display_val = format_currency(value) if 'Balance' in title or 'Check' in subtitle else format_percentage(value)
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-title">{title}</div>
                    <div class="metric-value">{display_val}</div>
                    <div class="metric-sub" style="color:{color}">{subtitle}</div>
                </div>
                """, unsafe_allow_html=True)
        
        # Matrix Preview
        st.markdown("#### Transition Matrix Preview")
        st.markdown(render_matrix_html(trans_arr, prev_arr), unsafe_allow_html=True)
        
        # Generate Button
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if not st.session_state.period:
                st.info("💡 Add a Period Label in sidebar for better exports")
            if st.button("🚀 Generate Dashboard", type="primary", use_container_width=True):
                # Pre-compute stats for performance
                stats = compute_statistics_cached(
                    tuple(map(tuple, trans_arr)), 
                    tuple(prev_arr)
                )
                st.session_state.update({
                    "generated": True,
                    "stats_cache": stats
                })
                st.success("✨ Dashboard ready! Switch to Analytics tab.")
                st.rerun()

    elif not st.session_state.generated:
        st.markdown("""
        <div class="upload-zone">
            <div style="font-size:42px;margin-bottom:12px;">📁</div>
            <div style="font-weight:600;font-size:16px;margin-bottom:6px;">
                Upload your Excel template
            </div>
            <div style="color:var(--neutral);font-size:13px;">
                Supports .xlsx files with NRB-compliant 5×5 grade matrix
            </div>
        </div>
        """, unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# TAB 2: DASHBOARD (Visualizations)
# ═════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    if not st.session_state.generated:
        st.info("👆 Upload and generate a matrix first")
    else:
        # Retrieve cached stats
        stats = st.session_state.stats_cache or compute_statistics_cached(
            tuple(map(tuple, st.session_state.matrix)),
            tuple(st.session_state.prev)
        )
        
        # Header Metrics
        st.markdown("#### 🎯 Key Performance Indicators")
        kpi_cols = st.columns(4)
        kpi_data = [
            ("Total Portfolio", format_currency(stats["total_closing"]), 
             f"↑{format_percentage(stats['upgrade_pct'])} upgraded", "var(--primary)"),
            ("Retention Rate", format_percentage(stats["retention_pct"]),
             f"{format_currency(stats['retained'])} retained", "var(--success)"),
            ("Migration Rate", format_percentage(stats["migration_rate"]),
             "Loans changing grade", "var(--warning)"),
            ("Concentration", format_percentage(stats["concentration_ratio"]),
             f"Largest grade share", "var(--neutral)")
        ]
        
        for col, (title, value, subtitle, color) in zip(kpi_cols, kpi_data):
            col.markdown(f"""
            <div class="metric-card">
                <div class="metric-title">{title}</div>
                <div class="metric-value">{value}</div>
                <div class="metric-sub" style="color:{color}">{subtitle}</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.divider()
        
        # Grade Distribution Charts
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.markdown("##### Opening vs Closing Distribution")
            # FIXED: Added proper brackets for list comprehensions
            dist_data = pd.DataFrame({
                "Grade": [f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
                "Opening": [stats["grade_details"][i]["opening"] for i in range(N)],
                "Closing": [stats["grade_details"][i]["closing"] for i in range(N)]
            })
            st.bar_chart(
                dist_data.set_index("Grade"),
                color=["#64B5F6", "#4DB6AC"],
                height=300
            )
        
        with col_chart2:
            st.markdown("##### Retention Rates by Grade")
            retention_data = pd.DataFrame({
                "Grade": [f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
                "Retention %": stats["retention_rates"]
            }).set_index("Grade")
            st.bar_chart(
                retention_data,
                color="#81C784",
                height=300
            )
        
        st.divider()
        
        # Interactive Matrix Heatmap
        st.markdown("#### 🔥 Transition Flow Heatmap")
        
        heatmap_df = pd.DataFrame(
            st.session_state.matrix,
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)]
        )
        
        st.dataframe(
            heatmap_df.style.format("{:.1f}").background_gradient(
                cmap="RdYlGn_r", axis=None, vmin=0, 
                vmax=st.session_state.matrix.max() if st.session_state.matrix.max() > 0 else 1
            ),
            use_container_width=True,
            height=350
        )
        
        # Export Options
        st.divider()
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        
        with col_exp1:
            if st.button("🖼️ Generate Chart", use_container_width=True):
                with st.spinner("Rendering..."):
                    fig = build_figure_optimized(
                        GRADES, st.session_state.matrix, 
                        st.session_state.prev, st.session_state.period
                    )
                    st.session_state["last_fig"] = fig
                    st.pyplot(fig, use_container_width=True)
        
        if st.session_state.get("last_fig") is not None:
            with col_exp2:
                st.download_button(
                    "⬇️ PNG",
                    fig_to_bytes(st.session_state["last_fig"], "png", export_dpi),
                    "matrix.png", "image/png",
                    use_container_width=True
                )
            with col_exp3:
                st.download_button(
                    "⬇️ SVG",
                    fig_to_bytes(st.session_state["last_fig"], "svg", export_dpi),
                    "matrix.svg", "image/svg+xml",
                    use_container_width=True
                )

# ═════════════════════════════════════════════════════════════════════════════
# TAB 3: ANALYTICS (Advanced Statistics)
# ═════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    if not st.session_state.generated:
        st.info("👆 Generate a matrix first")
    else:
        stats = st.session_state.stats_cache or compute_statistics_cached(
            tuple(map(tuple, st.session_state.matrix)),
            tuple(st.session_state.prev)
        )
        
        # Detailed Tables
        st.markdown("#### 📋 Transition Amounts (NPR Crore)")
        
        # Main transition table
        main_df = pd.DataFrame(
            np.round(st.session_state.matrix, 2),
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES
        )
        main_df["Row Sum"] = np.round(main_df.sum(axis=1), 2)
        
        # Add closing row
        closing_vals = np.round(st.session_state.matrix.sum(axis=0), 2).tolist()
        closing_vals.append(np.round(st.session_state.matrix.sum(), 2))
        closing_df = pd.DataFrame(
            [closing_vals], 
            index=["Col Sum"], 
            columns=list(GRADES) + ["Row Sum"]
        )
        
        display_df = pd.concat([main_df, closing_df])
        st.dataframe(
            display_df.style.format("{:.2f}").set_properties(
                **{"text-align": "center"}
            ).set_table_styles([
                {"selector": "th", "props": [("background", "#F5F5F4"), ("font-weight", "600")]},
                {"selector": ".diag", "props": [("color", "var(--primary)"), ("font-weight", "700")]},
                {"selector": ".upgrade", "props": [("color", "var(--success)")]},
                {"selector": ".downgrade", "props": [("color", "var(--error)")]},
            ]),
            use_container_width=True
        )
        
        # Percentage View
        with st.expander("📊 View as Percentages of Opening", expanded=False):
            pct_df = pd.DataFrame(
                [[st.session_state.matrix[r,c]/st.session_state.prev[r]*100 
                  if st.session_state.prev[r] > 0 else 0 
                  for c in range(N)] for r in range(N)],
                index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
                columns=GRADES
            ).round(1)
            st.dataframe(
                pct_df.style.format("{:.1f}%").background_gradient(
                    cmap="RdYlGn_r", axis=None, vmin=0, vmax=100
                ),
                use_container_width=True
            )
        
        # Grade-Level Analytics
        st.markdown("#### 🔍 Grade-Level Deep Dive")
        
        for i in range(N):
            with st.expander(f"{ICONS[i]} {GRADES[i]} Analysis", expanded=(i==0)):
                detail = stats["grade_details"][i]
                
                # Metrics grid
                m1, m2, m3, m4, m5, m6 = st.columns(6)
                metrics = [
                    (m1, "Opening", format_currency(detail["opening"])),
                    (m2, "Closing", format_currency(detail["closing"])),
                    (m3, "Retained", format_currency(detail["retained"])),
                    (m4, "↑ Upgraded", format_currency(detail["upgraded_out"])),
                    (m5, "↓ Downgraded", format_currency(detail["downgraded_out"])),
                    (m6, "↗ Inflow", format_currency(detail["inflow"]))
                ]
                for col, label, value in metrics:
                    col.metric(label, value)
                
                # Retention gauge
                st.progress(min(100, int(detail["retention_rate"])), 
                           text=f"Retention: {detail['retention_rate']:.1f}%")
                
                # Net change indicator
                net_color = "🟢" if detail["net_change"] >= 0 else "🔴"
                st.caption(f"{net_color} Net Change: {format_currency(detail['net_change'])}")
        
        # Export Full Dataset
        st.divider()
        csv_buffer = io.StringIO()
        display_df.to_csv(csv_buffer)
        st.download_button(
            "⬇️ Export Full Data (CSV)",
            csv_buffer.getvalue(),
            f"nrb_transition_{st.session_state.period or 'data'}.csv",
            "text/csv",
            use_container_width=True
        )

# ═════════════════════════════════════════════════════════════════════════════
# TAB 4: GUIDE & DOCUMENTATION
# ═════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown("## ℹ️ User Guide")
    
    col_guide1, col_guide2 = st.columns([2, 1])
    
    with col_guide1:
        st.markdown("### 🚀 Quick Start")
        st.markdown("""
        1. **Prepare Template**: Create Excel with 5×5 grade matrix
        2. **Upload**: Use the Upload tab to load your file
        3. **Configure**: Set period label in sidebar
        4. **Generate**: Click "Generate Dashboard"
        5. **Analyze**: Explore Dashboard and Analytics tabs
        6. **Export**: Download PNG, SVG, CSV, or JSON outputs
        """)
        
        st.markdown("### 📐 NRB Grade Definitions")
        st.dataframe(
            pd.DataFrame({
                "Grade": GRADES,
                "Status": ["Performing", "Special Mention", "Classified", "Impaired", "Loss"],
                "Overdue Range": ["Current", "1-3 months", "3-6 months", "6-12 months", "12+ months"],
                "Risk Level": ["Low", "Medium-Low", "Medium", "High", "Critical"]
            }),
            use_container_width=True,
            hide_index=True
        )
        
        st.markdown("### 🔢 Matrix Interpretation")
        st.code("""
        ROWS (From)    = Opening grade at period start
        COLUMNS (To)   = Closing grade at period end
        
        Example: Substandard row
        ├─→ Good ............ 24.60 cr  (upgraded)
        ├─→ Watchlist ....... 20.29 cr  (upgraded)  
        ├─→ Substandard ..... 23.81 cr  (retained) ★
        ├─→ Doubtful ........ 52.99 cr  (downgraded)
        └─→ Bad .............  0.00 cr
        ──────────────────────────────
        Row Sum (Opening): 121.68 cr
        """, language="text")
    
    with col_guide2:
        st.markdown("### 🎨 Color Legend")
        st.markdown("""
        <div style="font-size:13px;line-height:1.8;">
            <span style="background:#B5D4F4;padding:2px 8px;border-radius:4px;">🔵 Diagonal</span> Retained<br>
            <span style="background:#EAF3DE;padding:2px 8px;border-radius:4px;">🟢 Upgrade</span> Better grade<br>
            <span style="background:#FAEEDA;padding:2px 8px;border-radius:4px;">🟡 Mild ↓</span> &lt;5% of opening<br>
            <span style="background:#F0997B;padding:2px 8px;border-radius:4px;">🟠 Mod ↓</span> 5-30% of opening<br>
            <span style="background:#A32D2D;color:white;padding:2px 8px;border-radius:4px;">🔴 Severe ↓</span> &gt;30% of opening<br>
            <span style="background:#F1EFE8;padding:2px 8px;border-radius:4px;">⬜ Zero</span> No flow
        </div>
        """, unsafe_allow_html=True)
        
        st.divider()
        
        st.markdown("### ⚡ Performance Tips")
        st.markdown("""
        - ✅ Use .xlsx format (faster parsing)
        - ✅ Keep file size &lt;10MB for best performance
        - ✅ Cache persists for 1 hour (auto-refreshes)
        - ✅ Export DPI 220 balances quality/speed
        """)
    
    st.divider()
    st.markdown("""
    <div style="text-align:center;color:var(--neutral);font-size:12px;">
        <strong>NRB Loan Transition Matrix Tool</strong><br>
        Built for Nepal Rastra Bank reporting standards • v2.1<br>
        <em>Performance optimized • Cache-enabled • Mobile responsive</em>
    </div>
    """, unsafe_allow_html=True)


# ── Optimized Visualization Functions ────────────────────────────────────────
def build_figure_optimized(grades, trans, prev, period, figsize=(14, 8)):
    """Optimized matplotlib figure builder with reduced overhead."""
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    
    # Pre-compute values
    n = len(grades)
    col_closing = trans.sum(axis=0)
    
    # Layout constants
    RH, CW, HW = 0.95, 1.38, 2.45
    th = (n + 2) * RH
    tw = HW + n * CW
    
    # Create figure with optimized DPI
    fig, ax = plt.subplots(figsize=figsize, dpi=100)
    fig.patch.set_facecolor("#FAFAF9")
    ax.set_facecolor("#FAFAF9")
    ax.set_aspect("equal")
    ax.axis("off")
    
    # Color constants
    COLORS = {
        "canvas": "#FAFAF9", "header": "#F5F5F4", "row_hdr": "#F1EFE8",
        "grid": "#E0E0E0", "outer": "#BDBDBD",
        "diag": ("#B5D4F4", "#0D47A1"), "upgrade": ("#EAF3DE", "#1B5E20"),
        "zero": ("#F1EFE8", "#757575"), "mild": ("#FAEEDA", "#412402"),
        "mod": ("#F0997B", "#4A1B0C"), "severe": ("#A32D2D", "#FFEBEE"),
        "text": ("#1A1A18", "#757575")
    }
    
    def draw_cell_opt(x, y, w, h, bg, lines, fgs, align="center"):
        """Optimized cell drawer with minimal operations."""
        ax.add_patch(mpatches.Rectangle(
            (x, y), w, h, lw=0.5, edgecolor=COLORS["grid"], 
            facecolor=bg, zorder=2, antialiased=True
        ))
        tx = x + (0.10 if align=="left" else 0.50) * w
        if len(lines) == 1:
            ax.text(tx, y+h/2, lines[0], ha=align, va="center",
                   fontsize=9, color=fgs[0], fontweight="bold", zorder=3)
        else:
            ax.text(tx, y+h*0.64, lines[0], ha=align, va="center",
                   fontsize=9, color=fgs[0], fontweight="bold", zorder=3)
            ax.text(tx, y+h*0.29, lines[1], ha=align, va="center", 
                   fontsize=7.5, color=fgs[1], zorder=3)
    
    # Header row
    ty = (n+1) * RH
    draw_cell_opt(ax, 0, ty, HW, RH, COLORS["header"],
                 [f"Period: {period}", "Opening → Closing"],
                 COLORS["text"], align="left")
    
    # Column headers
    for ci, g in enumerate(grades):
        draw_cell_opt(ax, HW+ci*CW, ty, CW, RH, COLORS["header"],
                     [g, f"Close: {col_closing[ci]:.0f}"], COLORS["text"])
    
    # Matrix cells
    for ri in range(n):
        y = (n-ri) * RH
        draw_cell_opt(ax, 0, y, HW, RH, COLORS["row_hdr"],
                     [grades[ri], f"Open: {prev[ri]:.0f}"], COLORS["text"], align="left")
        for ci in range(n):
            v = trans[ri, ci]
            if ri == ci:
                bg, fg = COLORS["diag"]
            elif v == 0:
                bg, fg = COLORS["zero"]
            elif ci < ri:
                bg, fg = COLORS["upgrade"]
            else:
                pct = v/prev[ri]*100 if prev[ri]>0 else 0
                if pct < 5: bg, fg = COLORS["mild"]
                elif pct < 30: bg, fg = COLORS["mod"]
                else: bg, fg = COLORS["severe"]
            
            if v == 0 and ri != ci:
                draw_cell_opt(ax, HW+ci*CW, y, CW, RH, bg, ["—"], [fg])
            else:
                p = v/prev[ri]*100 if prev[ri]>0 else 0
                draw_cell_opt(ax, HW+ci*CW, y, CW, RH, bg,
                             [f"{v:.1f}", f"({p:.0f}%)"], [fg, fg])
    
    # Outer border
    ax.add_patch(mpatches.Rectangle(
        (0, 0), tw, th, fill=False, lw=1, edgecolor=COLORS["outer"], zorder=4))
    
    # Legend (simplified)
    legend_data = [
        (COLORS["diag"][0], "Retained"), (COLORS["upgrade"][0], "Upgrade"),
        (COLORS["mild"][0], "Mild ↓"), (COLORS["mod"][0], "Mod ↓"),
        (COLORS["severe"][0], "Severe ↓"), (COLORS["zero"][0], "Zero")
    ]
    patches = [mpatches.Patch(facecolor=c, edgecolor=COLORS["grid"], 
                             lw=0.5, label=l) for c, l in legend_data]
    ax.legend(handles=patches, loc="upper center", bbox_to_anchor=(0.5, -0.08),
             ncol=3, fontsize=7.5, frameon=True, columnspacing=1.2)
    
    # Title
    fig.suptitle("Loan Quality Transition Matrix", 
                fontsize=14, fontweight="bold", y=0.98, color=COLORS["text"][0])
    
    plt.tight_layout(rect=(0, 0.03, 1, 0.99))
    return fig


def fig_to_bytes(fig, fmt="png", dpi=150):
    """Optimized figure export with memory management."""
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, dpi=dpi, bbox_inches="tight",
                pad_inches=0.1, facecolor=fig.get_facecolor(),
                optimize=True)
    buf.seek(0)
    return buf.read()


# ── Dependency Check ─────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    st.error("""
    ### ⚠️ Missing Dependency
    Please install required package:
    ```bash
    pip install openpyxl
    ```
    Then restart the application.
    """)
    st.stop()
