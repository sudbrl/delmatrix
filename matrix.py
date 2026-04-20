import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import pandas as pd
import streamlit as st
import io
import json

# ── Dependency Check ──────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    st.error("""
    ❌ **Missing dependency: `openpyxl`**
    
    Run one of the following in your terminal:
    ```bash
    pip install openpyxl
    ```
    or
    ```bash
    conda install openpyxl
    ```
    Then restart the app.
    """)
    st.stop()

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="NRB Loan Transition Matrix",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

    .stApp { background-color: #F5F3EF; color: #2C2C2A;
             font-family: 'IBM Plex Sans', sans-serif; }
    .main .block-container { background-color: #F5F3EF; padding-top: 1.5rem; }

    [data-testid="stSidebar"] {
        background-color: #EDE9E3; border-right: 1px solid #D4CFCA;
    }
    [data-testid="stSidebar"] * { color: #2C2C2A !important; }

    h1,h2,h3,h4,h5,h6 {
        color: #1A1A18 !important;
        font-family: 'IBM Plex Sans', sans-serif !important;
    }

    .gh-card {
        background: #FFFFFF; border: 1px solid #D4CFCA;
        border-radius: 8px; padding: 14px 18px; margin: 6px 0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .gh-card-title {
        color: #7A7670; font-size: 11px; font-weight: 600;
        text-transform: uppercase; letter-spacing: .5px; margin-bottom: 2px;
        font-family: 'IBM Plex Mono', monospace;
    }
    .gh-card-value { color: #1A1A18; font-size: 24px; font-weight: 700; }
    .gh-green { color: #2E7D32; } .gh-red { color: #C62828; }
    .gh-blue  { color: #1565C0; } .gh-amber { color: #E65100; }

    .gh-header {
        background: linear-gradient(135deg, #FFFFFF, #F0EDE7);
        border: 1px solid #D4CFCA; border-radius: 8px;
        padding: 20px 28px; margin-bottom: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .gh-header h1 { font-size: 26px !important; margin: 0 0 6px !important; color: #1A1A18 !important; }
    .gh-header p  { color: #7A7670; font-size: 13px; margin: 0; }

    .badge {
        display: inline-block; padding: 2px 8px; border-radius: 12px;
        font-size: 11px; font-weight: 600; margin-right: 4px;
    }
    .badge-blue  { background: #BBDEFB; color: #0D47A1; }
    .badge-green { background: #C8E6C9; color: #1B5E20; }

    .upload-box {
        background: #FFFFFF; border: 2px dashed #BDBAB4;
        border-radius: 10px; padding: 32px 24px; text-align: center;
        margin: 16px 0; transition: border-color 0.2s;
    }
    .upload-box:hover { border-color: #1565C0; }
    .upload-box-icon { font-size: 42px; margin-bottom: 10px; }
    .upload-box-title { color: #1A1A18; font-size: 17px; font-weight: 700; margin-bottom: 6px; }
    .upload-box-sub { color: #7A7670; font-size: 13px; line-height: 1.6; }

    .template-card {
        background: #EEF4FF; border: 1px solid #BBDEFB;
        border-radius: 8px; padding: 14px 18px; margin: 12px 0;
    }
    .template-card-title { color: #0D47A1; font-size: 13px; font-weight: 700; margin-bottom: 6px; }
    .template-card pre {
        background: #F5F8FF; border: 1px solid #BBDEFB; border-radius: 6px;
        padding: 10px 14px; font-size: 11px; color: #1565C0;
        font-family: 'IBM Plex Mono', monospace; margin: 0;
        white-space: pre-wrap;
    }

    .parse-success {
        background: #E8F5E9; border: 1px solid #A5D6A7;
        border-radius: 8px; padding: 12px 16px; margin: 10px 0;
        color: #1B5E20; font-size: 13px; font-weight: 600;
    }
    .parse-error {
        background: #FFEBEE; border: 1px solid #FFCDD2;
        border-radius: 8px; padding: 12px 16px; margin: 10px 0;
        color: #B71C1C; font-size: 13px;
    }

    .plot-box {
        background: #FFFFFF; border: 1px solid #D4CFCA;
        border-radius: 8px; padding: 16px; margin-top: 12px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .gh-divider { border: none; border-top: 1px solid #D4CFCA; margin: 16px 0; }

    .stButton>button {
        background: #1565C0; color: #FFFFFF;
        border: none; border-radius: 6px; font-weight: 600; padding: 8px 20px;
        font-family: 'IBM Plex Sans', sans-serif;
        transition: background 0.15s ease;
    }
    .stButton>button:hover { background: #1976D2; }
    .stDownloadButton>button {
        background: #F0EDE7; color: #2C2C2A;
        border: 1px solid #D4CFCA; border-radius: 6px;
        font-family: 'IBM Plex Sans', sans-serif;
    }

    .stTabs [data-baseweb="tab-list"] {
        background: #F5F3EF; border-bottom: 1px solid #D4CFCA;
    }
    .stTabs [data-baseweb="tab"] {
        color: #7A7670; background: transparent;
        font-family: 'IBM Plex Sans', sans-serif;
    }
    .stTabs [aria-selected="true"] {
        color: #1A1A18 !important; border-bottom: 2px solid #1565C0 !important;
    }

    .matrix-summary {
        background: #FFFFFF; border: 1px solid #D4CFCA;
        border-radius: 8px; overflow: hidden; margin-top: 16px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }
    .matrix-summary table {
        width: 100%; border-collapse: collapse;
        font-family: 'IBM Plex Mono', monospace; font-size: 12px;
    }
    .matrix-summary th {
        background: #EDE9E3; color: #5F5E5A; padding: 8px 12px;
        text-align: center; font-weight: 600; letter-spacing: .3px;
        border-bottom: 1px solid #D4CFCA;
    }
    .matrix-summary td {
        padding: 7px 12px; text-align: center;
        border-bottom: 1px solid #EDE9E3; color: #2C2C2A;
    }
    .matrix-summary tr:last-child td { border-bottom: none; }
    .matrix-summary td.diag { color: #1565C0; font-weight: 700; }
    .matrix-summary td.upgrade { color: #2E7D32; }
    .matrix-summary td.downgrade { color: #C62828; }
    .matrix-summary td.zero { color: #BDBAB4; }
    .matrix-summary td.row-hdr {
        text-align: left; color: #1A1A18; font-weight: 600;
        background: #F8F6F2; padding-left: 14px;
    }
    .matrix-summary td.totals-row {
        background: #EEF4FF; color: #1565C0; font-weight: 700;
        border-top: 2px solid #D4CFCA;
    }
    .matrix-summary td.totals-hdr {
        background: #EEF4FF; color: #5F5E5A; font-weight: 700;
        text-align: left; padding-left: 14px;
        border-top: 2px solid #D4CFCA;
    }

    ::-webkit-scrollbar { width: 8px; }
    ::-webkit-scrollbar-track { background: #F5F3EF; }
    ::-webkit-scrollbar-thumb { background: #D4CFCA; border-radius: 4px; }

    [data-testid="stFileUploader"] {
        background: #FFFFFF !important;
        border: 2px dashed #BDBAB4 !important;
        border-radius: 10px !important;
        padding: 8px !important;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #1565C0 !important;
    }
    .stFileUploader label { color: #2C2C2A !important; }

    .stInfo { background: #EEF4FF !important; color: #0D47A1 !important; }
    .stWarning { background: #FFF8E1 !important; color: #E65100 !important; }
    .stSuccess { background: #E8F5E9 !important; color: #1B5E20 !important; }

    .stTextInput input {
        background: #FFFFFF !important; border: 1px solid #D4CFCA !important;
        color: #1A1A18 !important; border-radius: 6px !important;
    }
    .stTextInput label { color: #5F5E5A !important; }
    .stSelectSlider label { color: #5F5E5A !important; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = 5

# ── Color Scheme ──────────────────────────────────────────────────────────────
CANVAS_BG, HEADER_BG, ROW_HDR_BG = "#FAFAF8", "#E8E6DF", "#F1EFE8"
GRID_EDGE, OUTER_EDGE = "#D0CCC3", "#BDB7AC"
DIAG_BG, DIAG_FG = "#B5D4F4", "#042C53"
UPG_BG,  UPG_FG  = "#EAF3DE", "#173404"
ZERO_BG, ZERO_FG = "#F1EFE8", "#888780"
MILD_BG, MILD_FG = "#FAEEDA", "#412402"
MOD_BG,  MOD_FG  = "#F0997B", "#4A1B0C"
SEV_BG,  SEV_FG  = "#A32D2D", "#FCEBEB"
TEXT_DARK, TEXT_MID = "#2C2C2A", "#5F5E5A"

plt.rcParams.update({"font.family": "DejaVu Sans", "figure.dpi": 140})

# ── Session State ─────────────────────────────────────────────────────────────
defaults = {
    "prev": None,
    "matrix": None,
    "period": "",
    "generated": False,
    "upload_error": None,
    "filename": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Helper Functions ──────────────────────────────────────────────────────────

def parse_template(file_bytes: bytes):
    """
    Parse uploaded Excel template.
    
    Expected format:
        Row 0 (header): blank, Good, Watchlist, Substandard, Doubtful, Bad [, Grand Total]
        Rows 1-5: grade name, 5 values [, optional grand total]
    
    Returns:
        prev  : np.ndarray shape (5,)  — opening balances (row sums)
        trans : np.ndarray shape (5,5) — transition amounts
    
    Raises:
        ValueError with descriptive message on any parse failure.
    """
    # ── Try reading with openpyxl engine first, fallback to xlrd ────────────
    try:
        df = pd.read_excel(
            io.BytesIO(file_bytes),
            header=None,
            engine="openpyxl"
        )
    except Exception as e:
        # Try xlrd as fallback for older .xls files
        try:
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                header=None,
                engine="xlrd"
            )
        except Exception:
            raise ValueError(
                f"Could not read Excel file: {e}. "
                "Ensure the file is a valid .xlsx or .xls format."
            )

    if df.empty:
        raise ValueError("The uploaded file appears to be empty.")

    def norm(s: str) -> str:
        return str(s).strip().lower()

    grade_norms = [norm(g) for g in GRADES]

    # ── Find header row ───────────────────────────────────────────────────
    header_row_idx = None
    for i, row in df.iterrows():
        matches = sum(1 for cell in row if norm(str(cell)) in grade_norms)
        if matches >= 4:
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError(
            "Could not find header row with NRB grade names. "
            "Ensure your file contains: Good, Watchlist, Substandard, Doubtful, Bad "
            "in a single header row."
        )

    # ── Map column indices ────────────────────────────────────────────────
    header = df.iloc[header_row_idx]
    col_map = {}
    for ci, cell in enumerate(header):
        n = norm(str(cell))
        if n in grade_norms:
            col_map[n] = ci

    if len(col_map) < 5:
        raise ValueError(
            f"Header row found at row {header_row_idx} but only "
            f"{len(col_map)}/5 grades detected: {list(col_map.keys())}. "
            f"Missing: {[g for g in grade_norms if g not in col_map]}"
        )

    # ── Find data rows ────────────────────────────────────────────────────
    data_rows = {}
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]
        first_val = norm(str(row.iloc[0]))
        if first_val in grade_norms:
            data_rows[first_val] = row

    if len(data_rows) < 5:
        missing = [g for g in grade_norms if g not in data_rows]
        raise ValueError(
            f"Only {len(data_rows)}/5 grade rows found. "
            f"Missing rows: {missing}. "
            "Check that each data row starts with the exact grade name."
        )

    # ── Build transition matrix ───────────────────────────────────────────
    trans = np.zeros((5, 5), dtype=float)
    for ri, g_from in enumerate(grade_norms):
        row = data_rows[g_from]
        for ci, g_to in enumerate(grade_norms):
            val = row.iloc[col_map[g_to]]
            try:
                trans[ri, ci] = float(val) if pd.notna(val) else 0.0
            except (ValueError, TypeError):
                trans[ri, ci] = 0.0

    # Opening balance = row sums
    prev = trans.sum(axis=1)

    # ── Sanity checks ─────────────────────────────────────────────────────
    if trans.sum() == 0:
        raise ValueError(
            "All transition values are zero. "
            "Please check that your data cells contain numeric values."
        )

    if np.any(trans < 0):
        raise ValueError(
            "Negative values detected in the transition matrix. "
            "All loan amounts should be non-negative."
        )

    return prev, trans


def cell_colors(val: float, ri: int, ci: int, prev: np.ndarray):
    """Return (background_color, foreground_color) for a matrix cell."""
    if ri == ci:
        return DIAG_BG, DIAG_FG
    if val == 0:
        return ZERO_BG, ZERO_FG
    if ci < ri:
        return UPG_BG, UPG_FG
    pct = val / prev[ri] * 100 if prev[ri] > 0 else 0
    if pct < 5:
        return MILD_BG, MILD_FG
    if pct < 30:
        return MOD_BG, MOD_FG
    return SEV_BG, SEV_FG


def draw_cell(ax, x, y, w, h, bg, lines, fgs,
              ec=GRID_EDGE, fs1=9.5, fs2=8.0,
              w1="bold", w2="normal", ha="center"):
    """Draw a single matrix cell with optional two-line text."""
    ax.add_patch(mpatches.Rectangle(
        (x, y), w, h, lw=0.8, edgecolor=ec, facecolor=bg, zorder=2))
    tx = x + (0.10 if ha == "left" else 0.50) * w
    if len(lines) == 1:
        ax.text(tx, y + h / 2, lines[0], ha=ha, va="center",
                fontsize=fs1, color=fgs[0], fontweight=w1, zorder=3)
    else:
        ax.text(tx, y + h * .64, lines[0], ha=ha, va="center",
                fontsize=fs1, color=fgs[0], fontweight=w1, zorder=3)
        ax.text(tx, y + h * .29, lines[1], ha=ha, va="center",
                fontsize=fs2, color=fgs[1], fontweight=w2, zorder=3)


def build_figure(grades, trans, prev, period):
    """Build and return the full transition matrix matplotlib figure."""
    n = len(grades)
    col_closing = trans.sum(axis=0)

    RH, CW, HW = 0.95, 1.38, 2.45
    th = (n + 2) * RH
    tw = HW + n * CW

    fig, ax = plt.subplots(figsize=(tw + .7, th + 1.8))
    fig.patch.set_facecolor(CANVAS_BG)
    ax.set_facecolor(CANVAS_BG)
    ax.set_aspect("equal")
    ax.axis("off")

    # ── Header row ────────────────────────────────────────────────────────
    ty = (n + 1) * RH
    draw_cell(ax, 0, ty, HW, RH, HEADER_BG,
              [f"Period: {period}", "Opening  →  Closing"],
              [TEXT_DARK, TEXT_MID], ha="left", fs1=9.6, fs2=8.0)

    for ci, g in enumerate(grades):
        draw_cell(ax, HW + ci * CW, ty, CW, RH, HEADER_BG,
                  [g, f"Closing: {col_closing[ci]:,.1f}"],
                  [TEXT_DARK, TEXT_MID], fs1=9.3, fs2=7.8)

    # ── Data rows ─────────────────────────────────────────────────────────
    for ri in range(n):
        y = (n - ri) * RH
        draw_cell(ax, 0, y, HW, RH, ROW_HDR_BG,
                  [grades[ri], f"Opening: {prev[ri]:,.1f}"],
                  [TEXT_DARK, TEXT_MID], ha="left")
        for ci in range(n):
            v = trans[ri, ci]
            bg, fg = cell_colors(v, ri, ci, prev)
            p = v / prev[ri] * 100 if prev[ri] > 0 else 0
            if v == 0 and ri != ci:
                draw_cell(ax, HW + ci * CW, y, CW, RH, bg,
                          ["—"], [ZERO_FG], fs1=10, w1="normal")
            else:
                draw_cell(ax, HW + ci * CW, y, CW, RH, bg,
                          [f"{v:,.2f}", f"({p:.1f}%)"],
                          [fg, fg], fs1=9.5, fs2=8.0)

    # ── Outer border ──────────────────────────────────────────────────────
    ax.add_patch(mpatches.Rectangle(
        (0, 0), tw, th, fill=False, lw=1.15, edgecolor=OUTER_EDGE, zorder=4))

    ax.set_xlim(-.08, tw + .08)
    ax.set_ylim(-.08, th + .1)

    # ── Legend ────────────────────────────────────────────────────────────
    legend_items = [
        (DIAG_BG, "Retained (diagonal)"),
        (UPG_BG,  "Upgrade"),
        (MILD_BG, "Mild downgrade <5%"),
        (MOD_BG,  "Moderate 5–30%"),
        (SEV_BG,  "Severe >30%"),
        (ZERO_BG, "No flow"),
    ]
    patches = [
        mpatches.Patch(facecolor=c, edgecolor=GRID_EDGE, lw=.7, label=l)
        for c, l in legend_items
    ]
    leg = ax.legend(
        handles=patches, loc="upper center",
        bbox_to_anchor=(.5, -.07), ncol=3, fontsize=8.3,
        frameon=True, fancybox=False, edgecolor=GRID_EDGE,
        columnspacing=1.3, handlelength=1.5, borderpad=.6
    )
    leg.get_frame().set_facecolor(CANVAS_BG)

    fig.suptitle("Loan Quality Transition Matrix",
                 fontsize=13, fontweight="bold", y=.98, color=TEXT_DARK)
    plt.tight_layout(rect=(0, .05, 1, .99))
    return fig


def fig_to_bytes(fig, fmt: str = "png", dpi: int = 220) -> bytes:
    """Convert matplotlib figure to bytes for download."""
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, dpi=dpi, bbox_inches="tight",
                pad_inches=.15, facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf.read()


def compute_stats(trans: np.ndarray, prev: np.ndarray) -> dict:
    """Compute summary statistics from transition matrix."""
    n = len(prev)
    ret = sum(trans[i, i] for i in range(n))
    up  = sum(trans[r, c] for r in range(n) for c in range(r))
    dn  = sum(trans[r, c] for r in range(n) for c in range(r + 1, n))
    tot = trans.sum()
    return dict(
        total_opening=float(prev.sum()),
        total_closing=float(tot),
        col_closing=trans.sum(axis=0),
        retained=float(ret),
        upgraded=float(up),
        downgraded=float(dn),
        retention_pct=ret / tot * 100 if tot else 0,
        upgrade_pct=up / tot * 100 if tot else 0,
        downgrade_pct=dn / tot * 100 if tot else 0,
    )


def render_matrix_html(trans: np.ndarray, prev: np.ndarray) -> str:
    """Render the transition matrix as an HTML table string."""
    n = len(GRADES)
    col_closing = trans.sum(axis=0)

    rows_html = ""
    for ri in range(n):
        cells = ""
        for ci in range(n):
            v = trans[ri, ci]
            if ri == ci:
                cls = "diag"
            elif ci < ri:
                cls = "upgrade"
            elif v == 0:
                cls = "zero"
            else:
                pct = v / prev[ri] * 100 if prev[ri] > 0 else 0
                cls = "downgrade" if pct >= 5 else ""
            cells += f'<td class="{cls}">{v:,.2f}</td>'
        cells += f'<td style="color:#5F5E5A">{prev[ri]:,.2f}</td>'
        rows_html += (
            f'<tr><td class="row-hdr">'
            f'{ICONS[ri]} {GRADES[ri]}</td>{cells}</tr>'
        )

    total_cells = "".join(
        f'<td class="totals-row">{col_closing[ci]:,.2f}</td>'
        for ci in range(n)
    )
    total_cells += f'<td class="totals-row">{prev.sum():,.2f}</td>'
    headers = "".join(f"<th>{g}</th>" for g in GRADES)

    return f"""
    <div class="matrix-summary">
      <table>
        <thead>
          <tr>
            <th style="text-align:left;padding-left:14px;">From \\ To</th>
            {headers}
            <th>Opening (Row Sum)</th>
          </tr>
        </thead>
        <tbody>
          {rows_html}
          <tr>
            <td class="totals-hdr">↓ Closing (Col Sum)</td>
            {total_cells}
          </tr>
        </tbody>
      </table>
    </div>
    """


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
        <span style="font-size:28px;">🏦</span>
        <span style="color:#1A1A18;font-size:18px;font-weight:700;
                     font-family:'IBM Plex Sans',sans-serif;">
            NRB Matrix Tool</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    st.session_state.period = st.text_input(
        "📅 Period Label",
        value=st.session_state.period,
        placeholder="e.g. Poush 2081"
    )

    export_dpi = st.select_slider(
        "🖼️ Export DPI", [100, 150, 220, 300], value=220)

    st.markdown("---")

    if st.session_state.generated:
        if st.button("🔄 Upload New File", use_container_width=True):
            for key in ["prev", "matrix", "generated", "upload_error", "filename"]:
                st.session_state[key] = None if key != "generated" else False
            st.rerun()

    st.markdown("---")

    if st.session_state.generated and st.session_state.matrix is not None:
        st.markdown("### 📤 Export Data")
        exp_data = {
            "grades": GRADES,
            "period": st.session_state.period,
            "opening": st.session_state.prev.tolist(),
            "transition": st.session_state.matrix.tolist(),
        }
        st.download_button(
            "⬇️ Export JSON",
            json.dumps(exp_data, indent=2),
            "nrb_data.json",
            "application/json",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("""
    <div style="color:#9A9690;font-size:11px;line-height:1.6;
                font-family:'IBM Plex Mono',monospace;">
        <b style="color:#5F5E5A;">Row sum</b> = Opening balance<br>
        <b style="color:#5F5E5A;">Col sum</b> = Closing balance<br>
        <b style="color:#5F5E5A;">Diagonal</b> = Retained loans
    </div>
    """, unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="gh-header">
    <h1>🏦 NRB Loan Classification — Transition Matrix</h1>
    <p>
        Upload your Excel transition matrix template →
        <strong style="color:#1A1A18;">Row sum = Opening</strong> |
        <strong style="color:#1565C0;">Column sum = Closing</strong>
        <span class="badge badge-blue">NRB Standard</span>
        <span class="badge badge-green">Template Upload</span>
    </p>
</div>
""", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_upload, tab_heatmap, tab_stats, tab_about = st.tabs(
    ["📂 Upload Template", "📊 Heatmap", "📈 Statistics", "ℹ️ About"]
)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — UPLOAD TEMPLATE
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:

    # ── Template format guide ─────────────────────────────────────────────
    st.markdown("""
    <div class="template-card">
        <div class="template-card-title">📋 Expected Template Format</div>
        <pre>
| (blank)     | Good   | Watchlist | Substandard | Doubtful | Bad  | Grand Total |
|-------------|--------|-----------|-------------|----------|------|-------------|
| Good        | 3394.8 | 363.7     | 12.6        | 0        | 0    | 3771.1      |
| Watchlist   | 230.9  | 425.4     | 84.3        | 0        | 0    | 740.7       |
| Substandard | 24.6   | 20.3      | 23.8        | 53.0     | 0    | 121.7       |
| Doubtful    | 2.8    | 2.3       | 2.9         | 44.1     | 10.5 | 62.6        |
| Bad         | 20.6   | 0.4       | 0.1         | 0.3      | 247.8| 269.2       |
        </pre>
        <div style="color:#5F5E5A;font-size:12px;margin-top:8px;">
            ✅ Rows = Opening grade &nbsp;|&nbsp; Columns = Closing grade &nbsp;|&nbsp;
            Grand Total column is optional and will be ignored.
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")

    # ── File uploader ─────────────────────────────────────────────────────
    if not st.session_state.generated:
        uploaded = st.file_uploader(
            "Upload your Excel transition matrix template (.xlsx / .xls)",
            type=["xlsx", "xls"],
            help="Upload an Excel file with grade rows and columns as shown above.",
            label_visibility="visible",
        )

        if uploaded is not None:
            file_bytes = uploaded.read()
            try:
                prev, trans = parse_template(file_bytes)
                st.session_state.prev     = prev
                st.session_state.matrix   = trans
                st.session_state.filename = uploaded.name
                st.session_state.upload_error = None

                st.markdown(f"""
                <div class="parse-success">
                    ✅ Template parsed successfully: <b>{uploaded.name}</b> —
                    {N}×{N} matrix loaded, {int(prev.sum()):,} cr total opening balance.
                </div>
                """, unsafe_allow_html=True)

            except Exception as e:
                st.session_state.upload_error = str(e)
                st.session_state.prev   = None
                st.session_state.matrix = None
                st.markdown(f"""
                <div class="parse-error">
                    ❌ <b>Parse error:</b> {e}<br>
                    <span style="font-size:12px;">
                    Please check that your file matches the expected format above.
                    </span>
                </div>
                """, unsafe_allow_html=True)

    # ── Preview if parsed ─────────────────────────────────────────────────
    if (st.session_state.prev is not None
            and st.session_state.matrix is not None):

        prev_arr  = st.session_state.prev
        trans_arr = st.session_state.matrix

        st.markdown("---")
        st.markdown("### 📊 Parsed Data Preview")

        col_closing = trans_arr.sum(axis=0)
        sc1, sc2, sc3, sc4 = st.columns(4)

        with sc1:
            st.markdown(f"""
            <div class="gh-card">
                <div class="gh-card-title">Total Opening (Row Sum)</div>
                <div class="gh-card-value gh-blue">{prev_arr.sum():,.1f}
                    <span style="font-size:12px;color:#7A7670;"> cr</span></div>
            </div>""", unsafe_allow_html=True)

        with sc2:
            st.markdown(f"""
            <div class="gh-card">
                <div class="gh-card-title">Total Closing (Col Sum)</div>
                <div class="gh-card-value gh-blue">{trans_arr.sum():,.1f}
                    <span style="font-size:12px;color:#7A7670;"> cr</span></div>
            </div>""", unsafe_allow_html=True)

        with sc3:
            retained  = sum(trans_arr[i, i] for i in range(N))
            ret_pct   = retained / trans_arr.sum() * 100 if trans_arr.sum() else 0
            st.markdown(f"""
            <div class="gh-card">
                <div class="gh-card-title">Retained on Diagonal</div>
                <div class="gh-card-value gh-blue">{retained:,.1f}
                    <span style="font-size:12px;color:#7A7670;"> cr</span></div>
                <div style="font-size:12px;color:#1565C0;">{ret_pct:.1f}% of total</div>
            </div>""", unsafe_allow_html=True)

        with sc4:
            diff     = abs(trans_arr.sum() - prev_arr.sum())
            balanced = diff < 0.01
            st.markdown(f"""
            <div class="gh-card">
                <div class="gh-card-title">Balance Check</div>
                <div class="gh-card-value {'gh-green' if balanced else 'gh-amber'}"
                     style="font-size:18px;">
                    {'✅ Balanced' if balanced else f'⚠️ Δ {diff:,.2f}'}
                </div>
            </div>""", unsafe_allow_html=True)

        st.markdown("### 🔢 Matrix Preview")
        st.markdown("""
        <div style="color:#7A7670;font-size:12px;margin-bottom:8px;
                    font-family:'IBM Plex Mono',monospace;">
            Row sum = Opening balance &nbsp;|&nbsp; Column sum (↓ last row) = Closing balance
        </div>
        """, unsafe_allow_html=True)
        st.markdown(render_matrix_html(trans_arr, prev_arr), unsafe_allow_html=True)

        # ── Generate Button ───────────────────────────────────────────────
        st.markdown("")
        bc1, bc2, bc3 = st.columns([1, 2, 1])
        with bc2:
            if not st.session_state.period:
                st.warning("⚠️ Add a Period Label in the sidebar for better chart labeling.")
            if st.button("🚀 Generate Transition Matrix",
                         use_container_width=True, type="primary"):
                st.session_state.generated = True
                st.success("✅ Matrix generated! Switch to **📊 Heatmap** tab.")

    elif not st.session_state.generated:
        st.markdown("""
        <div class="upload-box">
            <div class="upload-box-icon">📂</div>
            <div class="upload-box-title">Upload your Excel template above</div>
            <div class="upload-box-sub">
                Supports <b>.xlsx</b> files with 5 NRB loan grade rows and columns.<br>
                The parser auto-detects the header row and grade labels.
            </div>
        </div>
        """, unsafe_allow_html=True)

    if st.session_state.generated:
        st.markdown("""
        <div class="parse-success">
            ✅ Matrix already generated. Head to the <b>📊 Heatmap</b> or
            <b>📈 Statistics</b> tabs. Use <b>🔄 Upload New File</b> in the sidebar to reset.
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — HEATMAP
# ══════════════════════════════════════════════════════════════════════════════
with tab_heatmap:
    if not st.session_state.generated:
        st.info("👈 Upload a template in **📂 Upload Template** tab and click Generate.")
    else:
        prev_arr  = st.session_state.prev
        trans_arr = st.session_state.matrix
        stats     = compute_stats(trans_arr, prev_arr)

        # ── KPI Cards ─────────────────────────────────────────────────────
        k1, k2, k3, k4 = st.columns(4)
        kpi_data = [
            (k1, "Total Opening",  stats["total_opening"],
             "NPR Crore (row sum)",         "gh-blue"),
            (k2, "Retained",       stats["retained"],
             f'↔ {stats["retention_pct"]:.1f}%',  "gh-blue"),
            (k3, "Upgraded",       stats["upgraded"],
             f'↑ {stats["upgrade_pct"]:.1f}%',    "gh-green"),
            (k4, "Downgraded",     stats["downgraded"],
             f'↓ {stats["downgrade_pct"]:.1f}%',  "gh-red"),
        ]
        for col, title, val, sub, cls in kpi_data:
            with col:
                st.markdown(f"""
                <div class="gh-card">
                    <div class="gh-card-title">{title}</div>
                    <div class="gh-card-value {cls}">{val:,.1f}</div>
                    <div style="font-size:12px;margin-top:4px;" class="{cls}">{sub}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)

        # ── Closing Balances ──────────────────────────────────────────────
        col_closing = stats["col_closing"]
        st.markdown("#### Closing Balances by Grade (Column Sums)")
        cc = st.columns(N)
        for ci in range(N):
            with cc[ci]:
                delta     = col_closing[ci] - prev_arr[ci]
                delta_str = f"+{delta:,.1f}" if delta >= 0 else f"{delta:,.1f}"
                d_color   = "#2E7D32" if delta >= 0 else "#C62828"
                st.markdown(f"""
                <div class="gh-card">
                    <div class="gh-card-title">{ICONS[ci]} {GRADES[ci]}</div>
                    <div class="gh-card-value" style="font-size:18px;">
                        {col_closing[ci]:,.1f}
                        <span style="font-size:11px;color:#7A7670;"> cr</span>
                    </div>
                    <div style="font-size:11px;color:{d_color};">
                        {delta_str} vs opening</div>
                </div>""", unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)

        # ── Chart ─────────────────────────────────────────────────────────
        with st.spinner("🎨 Rendering matrix chart…"):
            fig = build_figure(
                GRADES, trans_arr, prev_arr, st.session_state.period)

        st.markdown('<div class="plot-box">', unsafe_allow_html=True)
        st.pyplot(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)

        # ── Downloads ─────────────────────────────────────────────────────
        d1, d2, _ = st.columns([1, 1, 2])
        with d1:
            st.download_button(
                "⬇️ Download PNG",
                fig_to_bytes(fig, "png", export_dpi),
                "nrb_matrix.png", "image/png",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "⬇️ Download SVG",
                fig_to_bytes(fig, "svg", export_dpi),
                "nrb_matrix.svg", "image/svg+xml",
                use_container_width=True,
            )
        plt.close(fig)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 3 — STATISTICS
# ══════════════════════════════════════════════════════════════════════════════
with tab_stats:
    if not st.session_state.generated:
        st.info("👈 Generate a matrix first.")
    else:
        prev_arr    = st.session_state.prev
        trans_arr   = st.session_state.matrix
        col_closing = trans_arr.sum(axis=0)

        # ── Amounts Table ─────────────────────────────────────────────────
        st.markdown("### Transition Amounts (NPR Crore)")
        st.caption("Row sum = Opening balance | Column sum = Closing balance")

        df_a = pd.DataFrame(
            trans_arr.round(2),
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES,
        )
        df_a["Row Sum (Opening)"] = trans_arr.sum(axis=1).round(2)

        closing_row = list(col_closing.round(2)) + [round(trans_arr.sum(), 2)]
        df_closing  = pd.DataFrame(
            [closing_row],
            index=["↓ Col Sum (Closing)"],
            columns=df_a.columns,
        )
        df_combined = pd.concat([df_a, df_closing])
        st.dataframe(df_combined, use_container_width=True)

        # ── Percentage Table ──────────────────────────────────────────────
        st.markdown("### Percentage of Opening (%)")
        pct = [
            [trans_arr[r, c] / prev_arr[r] * 100 if prev_arr[r] > 0 else 0
             for c in range(N)]
            for r in range(N)
        ]
        df_p = pd.DataFrame(
            pct,
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES,
        ).round(1)
        st.dataframe(
            df_p.style.background_gradient(cmap="RdYlGn_r", axis=None),
            use_container_width=True,
        )

        # ── Summary Table ─────────────────────────────────────────────────
        st.markdown("### Opening vs Closing by Grade")
        summary = pd.DataFrame({
            "Grade":              [f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            "Opening (Row Sum)":  prev_arr.round(2),
            "Retained":           [round(trans_arr[i, i], 2) for i in range(N)],
            "Closing (Col Sum)":  col_closing.round(2),
            "Retention %":        [
                round(trans_arr[i, i] / prev_arr[i] * 100, 1)
                if prev_arr[i] > 0 else 0
                for i in range(N)
            ],
            "Net Change":         (col_closing - prev_arr).round(2),
        })
        st.dataframe(summary, use_container_width=True)

        # ── Grade Detail Expanders ────────────────────────────────────────
        st.markdown("### Grade Details")
        for i in range(N):
            with st.expander(f"{ICONS[i]} {GRADES[i]}"):
                ret  = trans_arr[i, i]
                up   = sum(trans_arr[i, c] for c in range(i))
                dn   = sum(trans_arr[i, c] for c in range(i + 1, N))
                inf  = sum(trans_arr[r, i] for r in range(N) if r != i)
                m1, m2, m3, m4, m5, m6 = st.columns(6)
                with m1: st.metric("Opening",           f"{prev_arr[i]:,.2f} cr")
                with m2: st.metric("Closing (col sum)", f"{col_closing[i]:,.2f} cr")
                with m3: st.metric("Retained",          f"{ret:,.2f} cr")
                with m4: st.metric("Upgraded out",      f"{up:,.2f} cr")
                with m5: st.metric("Downgraded out",    f"{dn:,.2f} cr")
                with m6: st.metric("Inflow from others",f"{inf:,.2f} cr")

        # ── CSV Download ──────────────────────────────────────────────────
        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)
        csv_buf = io.StringIO()
        df_combined.to_csv(csv_buf)
        st.download_button(
            "⬇️ Download CSV", csv_buf.getvalue(),
            "nrb_data.csv", "text/csv",
            use_container_width=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 4 — ABOUT
# ══════════════════════════════════════════════════════════════════════════════
with tab_about:
    st.markdown("""
## ℹ️ About

### How to Use

1. **Upload** your Excel template in the **📂 Upload Template** tab.
2. The app auto-detects the header row and grade labels — no manual mapping needed.
3. Set a **Period Label** in the sidebar (e.g. *Poush 2081*).
4. Click **🚀 Generate Transition Matrix**.
5. Explore the **📊 Heatmap** and **📈 Statistics** tabs.
6. Download PNG, SVG, or CSV exports as needed.

### Template Format

Your Excel file should have:
- A **header row** containing the 5 grade names: `Good`, `Watchlist`, `Substandard`, `Doubtful`, `Bad`
- **5 data rows** with each row starting with its grade name
- Cell values = amount flowing from the row grade (opening) to the column grade (closing)
- An optional **Grand Total** column (ignored by the parser)

### Matrix Logic
