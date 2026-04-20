import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import pandas as pd
import streamlit as st
import io
import json

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

    .stApp { background-color: #0d1117; color: #c9d1d9;
             font-family: 'IBM Plex Sans', sans-serif; }
    .main .block-container { background-color: #0d1117; padding-top: 1.5rem; }

    [data-testid="stSidebar"] {
        background-color: #161b22; border-right: 1px solid #30363d;
    }
    [data-testid="stSidebar"] * { color: #c9d1d9 !important; }

    h1,h2,h3,h4,h5,h6 {
        color: #f0f6fc !important;
        font-family: 'IBM Plex Sans', sans-serif !important;
    }

    /* ── Cards ── */
    .gh-card {
        background: #161b22; border: 1px solid #30363d;
        border-radius: 6px; padding: 14px 18px; margin: 6px 0;
    }
    .gh-card-title {
        color: #8b949e; font-size: 11px; font-weight: 600;
        text-transform: uppercase; letter-spacing: .5px; margin-bottom: 2px;
        font-family: 'IBM Plex Mono', monospace;
    }
    .gh-card-value { color: #f0f6fc; font-size: 24px; font-weight: 700; }
    .gh-green { color: #3fb950; } .gh-red { color: #f85149; }
    .gh-blue  { color: #58a6ff; } .gh-amber { color: #d29922; }

    /* ── Header ── */
    .gh-header {
        background: linear-gradient(135deg,#161b22,#0d1117);
        border: 1px solid #30363d; border-radius: 6px;
        padding: 20px 28px; margin-bottom: 20px;
    }
    .gh-header h1 { font-size: 26px !important; margin: 0 0 6px !important; }
    .gh-header p  { color: #8b949e; font-size: 13px; margin: 0; }

    /* ── Badges ── */
    .badge {
        display: inline-block; padding: 2px 8px; border-radius: 12px;
        font-size: 11px; font-weight: 600; margin-right: 4px;
    }
    .badge-blue  { background: #1f6feb; color: #f0f6fc; }
    .badge-green { background: #238636; color: #f0f6fc; }

    /* ── Plot Box ── */
    .plot-box {
        background: #161b22; border: 1px solid #30363d;
        border-radius: 6px; padding: 16px; margin-top: 12px;
    }
    .gh-divider { border: none; border-top: 1px solid #21262d; margin: 16px 0; }

    /* ── Buttons ── */
    .stButton>button {
        background: #238636; color: #f0f6fc;
        border: 1px solid rgba(240,246,252,.1);
        border-radius: 6px; font-weight: 600; padding: 8px 20px;
        font-family: 'IBM Plex Sans', sans-serif;
        transition: background 0.15s ease;
    }
    .stButton>button:hover { background: #2ea043; }
    .stDownloadButton>button {
        background: #21262d; color: #c9d1d9;
        border: 1px solid #30363d; border-radius: 6px;
        font-family: 'IBM Plex Sans', sans-serif;
    }

    /* ── Number Inputs — Smart & Clickable ── */
    div[data-baseweb="input"] input {
        background: #0d1117 !important;
        border: 1.5px solid #30363d !important;
        color: #f0f6fc !important;
        border-radius: 6px !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 15px !important;
        font-weight: 600 !important;
        text-align: center !important;
        cursor: text !important;
        padding: 8px 4px !important;
        transition: border-color 0.15s ease, box-shadow 0.15s ease,
                    background 0.15s ease !important;
        -moz-appearance: textfield !important;
    }
    div[data-baseweb="input"] input:hover {
        border-color: #58a6ff !important;
        background: #0d1117 !important;
    }
    div[data-baseweb="input"] input:focus {
        border-color: #58a6ff !important;
        box-shadow: 0 0 0 3px rgba(88,166,255,0.18) !important;
        background: #161b22 !important;
        outline: none !important;
    }
    /* Remove spinner arrows */
    div[data-baseweb="input"] input::-webkit-outer-spin-button,
    div[data-baseweb="input"] input::-webkit-inner-spin-button {
        -webkit-appearance: none !important;
        margin: 0 !important;
    }
    /* Step buttons styling */
    div[data-baseweb="input"] button {
        background: #21262d !important;
        border-color: #30363d !important;
        color: #8b949e !important;
        transition: background 0.1s ease, color 0.1s ease !important;
    }
    div[data-baseweb="input"] button:hover {
        background: #30363d !important;
        color: #f0f6fc !important;
    }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] {
        background: #0d1117; border-bottom: 1px solid #21262d;
    }
    .stTabs [data-baseweb="tab"] {
        color: #8b949e; background: transparent;
        font-family: 'IBM Plex Sans', sans-serif;
    }
    .stTabs [aria-selected="true"] {
        color: #f0f6fc !important; border-bottom: 2px solid #f78166 !important;
    }

    /* ── Grade Blocks ── */
    .grade-block {
        background: #161b22; border: 1px solid #30363d;
        border-radius: 8px; padding: 16px 20px; margin-bottom: 12px;
        display: flex; align-items: center; gap: 10px;
    }
    .grade-block-title {
        color: #f0f6fc; font-size: 17px; font-weight: 700;
        font-family: 'IBM Plex Sans', sans-serif;
    }
    .grade-block-sub {
        color: #8b949e; font-size: 12px; margin-top: 2px;
    }

    /* ── Remaining Tracker ── */
    .remaining-ok   { color: #3fb950; font-weight: 700; }
    .remaining-warn { color: #f85149; font-weight: 700; }
    .remaining-left { color: #d29922; font-weight: 700; }

    /* ── Matrix Summary Table ── */
    .matrix-summary {
        background: #161b22; border: 1px solid #30363d;
        border-radius: 8px; overflow: hidden; margin-top: 16px;
    }
    .matrix-summary table {
        width: 100%; border-collapse: collapse;
        font-family: 'IBM Plex Mono', monospace; font-size: 12px;
    }
    .matrix-summary th {
        background: #21262d; color: #8b949e; padding: 8px 12px;
        text-align: center; font-weight: 600; letter-spacing: .3px;
        border-bottom: 1px solid #30363d;
    }
    .matrix-summary td {
        padding: 7px 12px; text-align: center;
        border-bottom: 1px solid #21262d; color: #c9d1d9;
    }
    .matrix-summary tr:last-child td { border-bottom: none; }
    .matrix-summary td.diag { color: #58a6ff; font-weight: 700; }
    .matrix-summary td.upgrade { color: #3fb950; }
    .matrix-summary td.downgrade { color: #f85149; }
    .matrix-summary td.zero { color: #484f58; }
    .matrix-summary td.row-hdr {
        text-align: left; color: #f0f6fc; font-weight: 600;
        background: #161b22; padding-left: 14px;
    }
    .matrix-summary td.totals-row {
        background: #0d1117; color: #58a6ff; font-weight: 700;
        border-top: 2px solid #30363d;
    }
    .matrix-summary td.totals-hdr {
        background: #0d1117; color: #8b949e; font-weight: 700;
        text-align: left; padding-left: 14px;
        border-top: 2px solid #30363d;
    }

    /* ── Scrollbar ── */
    ::-webkit-scrollbar { width: 8px; }
    ::-webkit-scrollbar-track { background: #0d1117; }
    ::-webkit-scrollbar-thumb { background: #30363d; border-radius: 4px; }

    /* ── Label styling ── */
    .stNumberInput label {
        font-family: 'IBM Plex Sans', sans-serif !important;
        font-size: 11px !important;
        color: #8b949e !important;
        font-weight: 600 !important;
    }
    .stTextInput label {
        font-family: 'IBM Plex Sans', sans-serif !important;
        color: #8b949e !important;
    }
</style>

<script>
// Auto-select all text when clicking a number input for smart entry
document.addEventListener('focusin', function(e) {
    if (e.target && e.target.type === 'number') {
        setTimeout(() => e.target.select(), 10);
    }
});
</script>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────

GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = 5

DEFAULT_PREV = [100, 320, 430, 340, 210]
DEFAULT_MATRIX = [
    [100,   0,   0,   0,   0],
    [ 10, 300,  10,   0,   0],
    [  0, 120, 300,  10,   0],
    [  0,   0,   0, 300,  40],
    [  0,   0,  10,   0, 200],
]

# ── Color Scheme for Chart ────────────────────────────────────────────────────

CANVAS_BG, HEADER_BG, ROW_HDR_BG = "#FAFAF8", "#E8E6DF", "#F1EFE8"
GRID_EDGE, OUTER_EDGE = "#D0CCC3", "#BDB7AC"
DIAG_BG, DIAG_FG = "#B5D4F4", "#042C53"
UPG_BG,  UPG_FG  = "#EAF3DE", "#173404"
ZERO_BG, ZERO_FG = "#F1EFE8", "#888780"
MILD_BG, MILD_FG = "#FAEEDA", "#412402"
MOD_BG,  MOD_FG  = "#F0997B", "#4A1B0C"
SEV_BG,  SEV_FG  = "#A32D2D", "#FCEBEB"
CLOSE_BG, CLOSE_FG = "#E8E6DF", "#2C2C2A"
TEXT_DARK, TEXT_MID = "#2C2C2A", "#5F5E5A"

plt.rcParams.update({"font.family": "DejaVu Sans", "figure.dpi": 140})

# ── Session State Init ────────────────────────────────────────────────────────

if "prev" not in st.session_state:
    st.session_state.prev = DEFAULT_PREV.copy()
if "matrix" not in st.session_state:
    st.session_state.matrix = [row.copy() for row in DEFAULT_MATRIX]
if "period" not in st.session_state:
    st.session_state.period = "Poush"
if "generated" not in st.session_state:
    st.session_state.generated = False

# ── Helper Functions ──────────────────────────────────────────────────────────

def cell_colors(val, ri, ci, prev):
    if ri == ci:   return DIAG_BG, DIAG_FG
    if val == 0:   return ZERO_BG, ZERO_FG
    if ci < ri:    return UPG_BG,  UPG_FG
    pct = val / prev[ri] * 100 if prev[ri] > 0 else 0
    if pct < 5:    return MILD_BG, MILD_FG
    if pct < 30:   return MOD_BG,  MOD_FG
    return SEV_BG, SEV_FG


def draw_cell(ax, x, y, w, h, bg, lines, fgs,
              ec=GRID_EDGE, fs1=9.5, fs2=8.0,
              w1="bold", w2="normal", ha="center"):
    ax.add_patch(mpatches.Rectangle(
        (x, y), w, h, lw=0.8, edgecolor=ec, facecolor=bg, zorder=2))
    tx = x + (0.10 if ha == "left" else 0.50) * w
    if len(lines) == 1:
        ax.text(tx, y+h/2, lines[0], ha=ha, va="center",
                fontsize=fs1, color=fgs[0], fontweight=w1, zorder=3)
    else:
        ax.text(tx, y+h*.64, lines[0], ha=ha, va="center",
                fontsize=fs1, color=fgs[0], fontweight=w1, zorder=3)
        ax.text(tx, y+h*.29, lines[1], ha=ha, va="center",
                fontsize=fs2, color=fgs[1], fontweight=w2, zorder=3)


def build_figure(grades, trans, prev, period):
    """
    Rows = opening grade (row header shows opening balance)
    Columns = closing grade (column header shows CLOSING = sum of column)
    Extra bottom row = closing totals per column
    """
    n = len(grades)
    # Column closing totals = sum of each column
    col_closing = trans.sum(axis=0)
    # Row opening totals = prev (opening balance per grade)
    RH, CW, HW = 0.95, 1.38, 2.45
    # +1 row for closing totals row at bottom
    th = (n + 2) * RH
    tw = HW + n * CW

    fig, ax = plt.subplots(figsize=(tw + .7, th + 1.8))
    fig.patch.set_facecolor(CANVAS_BG)
    ax.set_facecolor(CANVAS_BG)
    ax.set_aspect("equal")
    ax.axis("off")

    # ── Top header row (grade column headers) ─────────────────────────────
    ty = (n + 1) * RH
    draw_cell(ax, 0, ty, HW, RH, HEADER_BG,
              [f"Period: {period}", "Opening  →  Closing"],
              [TEXT_DARK, TEXT_MID], ha="left", fs1=9.6, fs2=8.0)
    for ci, g in enumerate(grades):
        draw_cell(ax, HW + ci * CW, ty, CW, RH, HEADER_BG,
                  [g, f"Closing: {int(col_closing[ci]):,} cr"],
                  [TEXT_DARK, TEXT_MID], fs1=9.3, fs2=7.8)

    # ── Data rows ─────────────────────────────────────────────────────────
    for ri in range(n):
        y = (n - ri) * RH   # row 0 = top data row
        draw_cell(ax, 0, y, HW, RH, ROW_HDR_BG,
                  [grades[ri], f"Opening: {int(prev[ri]):,} cr"],
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
                          [f"{int(v):,}", f"({p:.1f}%)"],
                          [fg, fg], fs1=10, fs2=8.2)

    # ── Bottom closing-totals row ──────────────────────────────────────────
    y_bottom = 0
    draw_cell(ax, 0, y_bottom, HW, RH, "#D8D4CB",
              ["Closing Total", "(Sum of column)"],
              [TEXT_DARK, TEXT_MID], ha="left", fs1=9.0, fs2=7.8)
    for ci in range(n):
        draw_cell(ax, HW + ci * CW, y_bottom, CW, RH, DIAG_BG,
                  [f"{int(col_closing[ci]):,} cr", "↑ closing"],
                  [DIAG_FG, DIAG_FG], fs1=10, fs2=8.0)

    # ── Outer border ──────────────────────────────────────────────────────
    ax.add_patch(mpatches.Rectangle(
        (0, 0), tw, th, fill=False, lw=1.15, edgecolor=OUTER_EDGE, zorder=4))
    ax.set_xlim(-.08, tw + .08)
    ax.set_ylim(-.08, th + .1)

    # ── Legend ────────────────────────────────────────────────────────────
    legend = [
        (DIAG_BG, "Retained (diagonal)"),
        (UPG_BG,  "Upgrade"),
        (MILD_BG, "Mild downgrade <5%"),
        (MOD_BG,  "Moderate 5–30%"),
        (SEV_BG,  "Severe >30%"),
        (ZERO_BG, "No flow"),
    ]
    patches = [mpatches.Patch(facecolor=c, edgecolor=GRID_EDGE,
               lw=.7, label=l) for c, l in legend]
    leg = ax.legend(handles=patches, loc="upper center",
                    bbox_to_anchor=(.5, -.07), ncol=3, fontsize=8.3,
                    frameon=True, fancybox=False, edgecolor=GRID_EDGE,
                    columnspacing=1.3, handlelength=1.5, borderpad=.6)
    leg.get_frame().set_facecolor(CANVAS_BG)

    fig.suptitle("Loan Quality Transition Matrix",
                 fontsize=13, fontweight="bold", y=.98, color=TEXT_DARK)
    plt.tight_layout(rect=(0, .05, 1, .99))
    return fig


def fig_to_bytes(fig, fmt="png", dpi=220):
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, dpi=dpi, bbox_inches="tight",
                pad_inches=.15, facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf.read()


def compute_stats(trans, prev):
    n = len(prev)
    ret = sum(trans[i, i] for i in range(n))
    up  = sum(trans[r, c] for r in range(n) for c in range(r))
    dn  = sum(trans[r, c] for r in range(n) for c in range(r+1, n))
    tot = trans.sum()
    col_closing = trans.sum(axis=0)
    return dict(
        total_opening=int(prev.sum()),
        total_closing=int(tot),
        col_closing=col_closing,
        retained=int(ret), upgraded=int(up), downgraded=int(dn),
        retention_pct=ret/tot*100 if tot else 0,
        upgrade_pct=up/tot*100 if tot else 0,
        downgrade_pct=dn/tot*100 if tot else 0)


def render_matrix_html(trans, prev):
    """Render the live matrix preview as HTML with row=opening, col=closing."""
    n = len(GRADES)
    col_closing = trans.sum(axis=0)

    rows_html = ""
    for ri in range(n):
        cells = ""
        for ci in range(n):
            v = int(trans[ri, ci])
            if ri == ci:
                cls = "diag"
            elif ci < ri:
                cls = "upgrade"
            elif v == 0:
                cls = "zero"
            else:
                pct = v / prev[ri] * 100 if prev[ri] > 0 else 0
                cls = "upgrade" if ci < ri else (
                    "zero" if v == 0 else
                    ("" if pct < 5 else "downgrade"))
            cells += f'<td class="{cls}">{v:,}</td>'
        row_sum = int(sum(trans[ri]))
        remaining = int(prev[ri]) - row_sum
        rem_style = ("color:#3fb950" if remaining == 0 else
                     "color:#f85149" if remaining < 0 else "color:#d29922")
        cells += (f'<td style="color:#8b949e">{int(prev[ri]):,}</td>'
                  f'<td style="{rem_style};font-weight:700">'
                  f'{remaining:+,}</td>')
        rows_html += (f'<tr><td class="row-hdr">'
                      f'{ICONS[ri]} {GRADES[ri]}</td>{cells}</tr>')

    # Closing totals row (column sums)
    total_cells = ""
    for ci in range(n):
        total_cells += (f'<td class="totals-row">'
                        f'{int(col_closing[ci]):,}</td>')
    total_cells += (f'<td class="totals-row">{int(prev.sum()):,}</td>'
                    f'<td class="totals-row">—</td>')

    headers = "".join(f"<th>{g}</th>" for g in GRADES)

    html = f"""
    <div class="matrix-summary">
      <table>
        <thead>
          <tr>
            <th style="text-align:left;padding-left:14px;">From \\ To</th>
            {headers}
            <th>Opening</th>
            <th>Remaining</th>
          </tr>
        </thead>
        <tbody>
          {rows_html}
          <tr>
            <td class="totals-hdr">↓ Closing Total</td>
            {total_cells}
          </tr>
        </tbody>
      </table>
    </div>
    """
    return html


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
        <svg height="28" viewBox="0 0 16 16" width="28" fill="#f0f6fc">
            <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53
            5.47 7.59.4.07.55-.17.55-.38
            0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52
            -.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87
            2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95
            0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12
            0 0 .67-.21 2.2.82.64-.18 1.32-.27
            2-.27.68 0 1.36.09 2 .27 1.53-1.04
            2.2-.82 2.2-.82.44 1.1.16 1.92.08
            2.12.51.56.82 1.27.82 2.15 0 3.07-1.87
            3.75-3.65 3.95.29.25.54.73.54 1.48
            0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013
            8.013 0 0016 8c0-4.42-3.58-8-8-8z"/>
        </svg>
        <span style="color:#f0f6fc;font-size:18px;font-weight:700;
                     font-family:'IBM Plex Sans',sans-serif;">
            NRB Matrix Tool</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    st.session_state.period = st.text_input(
        "📅 Period Label", value=st.session_state.period)

    export_dpi = st.select_slider(
        "🖼️ Export DPI", [100, 150, 220, 300], value=220)

    st.markdown("---")
    if st.button("🔄 Reset to Defaults", use_container_width=True):
        st.session_state.prev = DEFAULT_PREV.copy()
        st.session_state.matrix = [r.copy() for r in DEFAULT_MATRIX]
        st.session_state.generated = False
        st.rerun()

    st.markdown("---")
    st.markdown("### 📤 Export Data")
    exp_data = {
        "grades": GRADES,
        "period": st.session_state.period,
        "opening": st.session_state.prev,
        "transition": st.session_state.matrix
    }
    st.download_button(
        "⬇️ Export JSON", json.dumps(exp_data, indent=2),
        "nrb_data.json", "application/json",
        use_container_width=True)

    st.markdown("---")
    st.markdown("""
    <div style="color:#484f58;font-size:11px;line-height:1.6;
                font-family:'IBM Plex Mono',monospace;">
        <b style="color:#8b949e;">Row sum</b> = Opening balance<br>
        <b style="color:#8b949e;">Col sum</b> = Closing balance<br>
        <b style="color:#8b949e;">Diagonal</b> = Retained loans
    </div>
    """, unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="gh-header">
    <h1>🏦 NRB Loan Classification — Transition Matrix</h1>
    <p>
        Enter opening balance per grade → distribute across closing grades →
        <strong style="color:#f0f6fc;">Row sum = Opening</strong> |
        <strong style="color:#58a6ff;">Column sum = Closing</strong>
        <span class="badge badge-blue">NRB Standard</span>
        <span class="badge badge-green">Auto Totals</span>
    </p>
</div>
""", unsafe_allow_html=True)


# ── Tabs ──────────────────────────────────────────────────────────────────────

tab_entry, tab_heatmap, tab_stats, tab_about = st.tabs(
    ["✏️ Enter Data", "📊 Heatmap", "📈 Statistics", "ℹ️ About"])


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — DATA ENTRY
# ══════════════════════════════════════════════════════════════════════════════

with tab_entry:
    st.markdown("""
    <div style="background:#161b22;border:1px solid #30363d;border-radius:6px;
                padding:14px 18px;margin-bottom:20px;">
        <div style="color:#f0f6fc;font-size:14px;font-weight:600;
                    margin-bottom:6px;">📐 How the matrix works</div>
        <div style="color:#8b949e;font-size:12px;line-height:1.7;
                    font-family:'IBM Plex Mono',monospace;">
            <b style="color:#c9d1d9;">Rows</b> = Opening grade &nbsp;|&nbsp;
            <b style="color:#c9d1d9;">Columns</b> = Closing grade &nbsp;|&nbsp;
            <b style="color:#3fb950;">Row sum = Opening balance</b> &nbsp;|&nbsp;
            <b style="color:#58a6ff;">Column sum = Closing balance</b>
        </div>
    </div>
    """, unsafe_allow_html=True)

    all_valid = True

    for ri in range(N):
        # Grade header
        grade_color = ["#3fb950", "#d29922", "#f0883e", "#f85149", "#8b949e"][ri]
        st.markdown(f"""
        <div class="grade-block">
            <span style="font-size:22px;">{ICONS[ri]}</span>
            <div>
                <div class="grade-block-title"
                     style="color:{grade_color};">{GRADES[ri]}</div>
                <div class="grade-block-sub">
                    Distribute opening balance → across all closing grades
                    (row sum must equal opening)
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Opening Balance
        col_open, col_info = st.columns([1, 3])
        with col_open:
            opening = st.number_input(
                f"Opening Balance (cr)",
                min_value=0, max_value=9_999_999,
                value=int(st.session_state.prev[ri]),
                step=10, key=f"op_{ri}",
                help=f"Total {GRADES[ri]} loan portfolio at period start (NPR Crore)")
            st.session_state.prev[ri] = opening

        with col_info:
            st.markdown(f"""
            <div style="padding-top:32px;color:#8b949e;font-size:12px;
                        font-family:'IBM Plex Mono',monospace;">
                Distribute <b style="color:#f0f6fc;">{opening:,} cr</b>
                across the 5 closing grades below. Click any field to edit.
            </div>
            """, unsafe_allow_html=True)

        # Distribution inputs (5 closing grade columns + remaining)
        dist_cols = st.columns(N + 1)

        row_vals = []
        for ci in range(N):
            with dist_cols[ci]:
                default_val = int(st.session_state.matrix[ri][ci])

                # Label with direction indicator
                if ci == ri:
                    label = f"✦ {GRADES[ci]}"
                elif ci < ri:
                    label = f"↑ {GRADES[ci]}"
                else:
                    label = f"↓ {GRADES[ci]}"

                v = st.number_input(
                    label,
                    min_value=0,
                    max_value=max(opening, 9_999_999),
                    value=min(default_val, max(opening, 0)),
                    step=5,
                    key=f"d_{ri}_{ci}",
                    help=f"Amount flowing from {GRADES[ri]} (opening) "
                         f"to {GRADES[ci]} (closing)")
                row_vals.append(v)

        st.session_state.matrix[ri] = row_vals

        # Remaining tracker
        allocated = sum(row_vals)
        remaining = opening - allocated

        with dist_cols[N]:
            if remaining == 0:
                cls = "remaining-ok"
                icon = "✅"
                msg = "Balanced"
                bar_color = "#3fb950"
            elif remaining > 0:
                cls = "remaining-left"
                icon = "🔸"
                msg = f"{remaining:,} unallocated"
                all_valid = False
                bar_color = "#d29922"
            else:
                cls = "remaining-warn"
                icon = "⚠️"
                msg = f"{abs(remaining):,} over-allocated"
                all_valid = False
                bar_color = "#f85149"

            pct_used = min(allocated / opening * 100, 100) if opening > 0 else 0

            st.markdown(f"""
            <div style="padding-top:22px;text-align:center;">
                <div style="color:#8b949e;font-size:10px;
                            text-transform:uppercase;letter-spacing:.6px;
                            font-family:'IBM Plex Mono',monospace;">
                    Remaining</div>
                <div class="{cls}" style="font-size:19px;margin:4px 0;
                             font-family:'IBM Plex Mono',monospace;">
                    {icon} {remaining:,}
                </div>
                <div style="color:#8b949e;font-size:10px;">{msg}</div>
                <div style="color:#484f58;font-size:10px;margin-top:3px;
                            font-family:'IBM Plex Mono',monospace;">
                    {allocated:,} / {opening:,}</div>
                <div style="background:#21262d;border-radius:4px;height:5px;
                            margin-top:7px;overflow:hidden;">
                    <div style="background:{bar_color};height:100%;
                                width:{pct_used:.1f}%;border-radius:4px;
                                transition:width .3s;"></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)

    # ── Grand Summary ─────────────────────────────────────────────────────

    total_opening   = sum(st.session_state.prev)
    trans_arr       = np.array(st.session_state.matrix, dtype=float)
    total_allocated = int(trans_arr.sum())
    total_remaining = int(total_opening - total_allocated)
    prev_arr        = np.array(st.session_state.prev, dtype=float)
    col_closing     = trans_arr.sum(axis=0)

    sc1, sc2, sc3, sc4 = st.columns(4)
    with sc1:
        st.markdown(f"""
        <div class="gh-card">
            <div class="gh-card-title">Total Opening (Row Sum)</div>
            <div class="gh-card-value gh-blue">{int(total_opening):,}
                <span style="font-size:12px;color:#8b949e;"> cr</span></div>
        </div>""", unsafe_allow_html=True)
    with sc2:
        st.markdown(f"""
        <div class="gh-card">
            <div class="gh-card-title">Total Closing (Col Sum)</div>
            <div class="gh-card-value gh-blue">{total_allocated:,}
                <span style="font-size:12px;color:#8b949e;"> cr</span></div>
        </div>""", unsafe_allow_html=True)
    with sc3:
        rem_color = "gh-green" if total_remaining == 0 else (
            "gh-amber" if total_remaining > 0 else "gh-red")
        st.markdown(f"""
        <div class="gh-card">
            <div class="gh-card-title">Unallocated</div>
            <div class="gh-card-value {rem_color}">{total_remaining:,}
                <span style="font-size:12px;color:#8b949e;"> cr</span></div>
        </div>""", unsafe_allow_html=True)
    with sc4:
        balance_icon = "✅" if total_remaining == 0 else "⚠️"
        balance_color = "gh-green" if total_remaining == 0 else "gh-red"
        st.markdown(f"""
        <div class="gh-card">
            <div class="gh-card-title">Balance Status</div>
            <div class="gh-card-value {balance_color}" style="font-size:18px;">
                {balance_icon} {'Balanced' if total_remaining == 0
                                else 'Unbalanced'}
            </div>
        </div>""", unsafe_allow_html=True)

    # ── Live Matrix Preview ────────────────────────────────────────────────
    st.markdown("### 🔢 Live Matrix Preview")
    st.markdown("""
    <div style="color:#8b949e;font-size:12px;margin-bottom:8px;
                font-family:'IBM Plex Mono',monospace;">
        Row sum = Opening balance &nbsp;|&nbsp;
        Column sum (↓ last row) = Closing balance
    </div>
    """, unsafe_allow_html=True)
    st.markdown(render_matrix_html(trans_arr, prev_arr), unsafe_allow_html=True)

    # ── Generate Button ───────────────────────────────────────────────────
    st.markdown("")
    bc1, bc2, bc3 = st.columns([1, 2, 1])
    with bc2:
        if not all_valid:
            st.warning("⚠️ Some grades are not fully allocated. "
                       "You can still generate.")
        if st.button("🚀 Generate Transition Matrix",
                     use_container_width=True, type="primary"):
            st.session_state.generated = True
            st.success("✅ Matrix generated! Switch to **📊 Heatmap** tab.")


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — HEATMAP
# ══════════════════════════════════════════════════════════════════════════════

with tab_heatmap:
    if not st.session_state.generated:
        st.info("👈 Enter data in **✏️ Enter Data** tab and click Generate.")
    else:
        prev_arr  = np.array(st.session_state.prev, dtype=float)
        trans_arr = np.array(st.session_state.matrix, dtype=float)
        stats = compute_stats(trans_arr, prev_arr)

        k1, k2, k3, k4 = st.columns(4)
        for col, title, val, sub, cls in [
            (k1, "Total Opening",  stats["total_opening"],
                 "NPR Crore (row sum)",        "gh-blue"),
            (k2, "Retained",       stats["retained"],
                 f'↔ {stats["retention_pct"]:.1f}%', "gh-blue"),
            (k3, "Upgraded",       stats["upgraded"],
                 f'↑ {stats["upgrade_pct"]:.1f}%',   "gh-green"),
            (k4, "Downgraded",     stats["downgraded"],
                 f'↓ {stats["downgrade_pct"]:.1f}%',  "gh-red"),
        ]:
            with col:
                st.markdown(f"""
                <div class="gh-card">
                    <div class="gh-card-title">{title}</div>
                    <div class="gh-card-value {cls}">{val:,}</div>
                    <div style="font-size:12px;margin-top:4px;"
                         class="{cls}">{sub}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)

        # Closing totals per grade
        col_closing = stats["col_closing"]
        st.markdown("#### Closing Balances by Grade (Column Sums)")
        cc = st.columns(N)
        for ci in range(N):
            with cc[ci]:
                delta = int(col_closing[ci]) - int(prev_arr[ci])
                delta_str = f"+{delta:,}" if delta >= 0 else f"{delta:,}"
                delta_color = "#3fb950" if delta >= 0 else "#f85149"
                st.markdown(f"""
                <div class="gh-card">
                    <div class="gh-card-title">{ICONS[ci]} {GRADES[ci]}</div>
                    <div class="gh-card-value" style="font-size:18px;">
                        {int(col_closing[ci]):,}
                        <span style="font-size:11px;color:#8b949e;"> cr</span>
                    </div>
                    <div style="font-size:11px;color:{delta_color};">
                        {delta_str} vs opening</div>
                </div>""", unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)

        with st.spinner("🎨 Rendering matrix chart…"):
            fig = build_figure(GRADES, trans_arr, prev_arr,
                               st.session_state.period)

        st.markdown('<div class="plot-box">', unsafe_allow_html=True)
        st.pyplot(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)
        d1, d2, _ = st.columns([1, 1, 2])
        with d1:
            st.download_button("⬇️ Download PNG",
                               fig_to_bytes(fig, "png", export_dpi),
                               "nrb_matrix.png", "image/png",
                               use_container_width=True)
        with d2:
            st.download_button("⬇️ Download SVG",
                               fig_to_bytes(fig, "svg", export_dpi),
                               "nrb_matrix.svg", "image/svg+xml",
                               use_container_width=True)
        plt.close(fig)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 3 — STATISTICS
# ══════════════════════════════════════════════════════════════════════════════

with tab_stats:
    if not st.session_state.generated:
        st.info("👈 Generate a matrix first.")
    else:
        prev_arr  = np.array(st.session_state.prev, dtype=float)
        trans_arr = np.array(st.session_state.matrix, dtype=float)
        col_closing = trans_arr.sum(axis=0)

        st.markdown("### Transition Amounts (NPR Crore)")
        st.caption("Row sum = Opening balance | Column sum = Closing balance")
        df_a = pd.DataFrame(
            trans_arr.astype(int),
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES)
        df_a["Row Sum (Opening)"] = trans_arr.sum(axis=1).astype(int)
        # Append closing row
        closing_row = list(col_closing.astype(int)) + [int(trans_arr.sum())]
        df_closing = pd.DataFrame(
            [closing_row],
            index=["↓ Col Sum (Closing)"],
            columns=df_a.columns)
        df_combined = pd.concat([df_a, df_closing])
        st.dataframe(df_combined, use_container_width=True)

        st.markdown("### Percentage of Opening (%)")
        pct = [[trans_arr[r, c] / prev_arr[r] * 100
                if prev_arr[r] > 0 else 0
                for c in range(N)] for r in range(N)]
        df_p = pd.DataFrame(
            pct,
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES).round(1)
        st.dataframe(df_p.style.background_gradient(
            cmap="RdYlGn_r", axis=None), use_container_width=True)

        st.markdown("### Opening vs Closing by Grade")
        summary = pd.DataFrame({
            "Grade": [f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            "Opening (Row Sum)":  prev_arr.astype(int),
            "Retained":           [int(trans_arr[i, i]) for i in range(N)],
            "Closing (Col Sum)":  col_closing.astype(int),
            "Retention %":        [round(trans_arr[i, i] / prev_arr[i] * 100, 1)
                                   if prev_arr[i] > 0 else 0 for i in range(N)],
            "Net Change":         (col_closing - prev_arr).astype(int),
        })
        st.dataframe(summary, use_container_width=True)

        st.markdown("### Grade Details")
        for i in range(N):
            with st.expander(f"{ICONS[i]} {GRADES[i]}"):
                ret = trans_arr[i, i]
                up  = sum(trans_arr[i, c] for c in range(i))
                dn  = sum(trans_arr[i, c] for c in range(i+1, N))
                inf = sum(trans_arr[r, i] for r in range(N) if r != i)
                opening_i  = int(prev_arr[i])
                closing_i  = int(col_closing[i])
                m1, m2, m3, m4, m5, m6 = st.columns(6)
                with m1: st.metric("Opening", f"{opening_i:,} cr")
                with m2: st.metric("Closing (col sum)", f"{closing_i:,} cr")
                with m3: st.metric("Retained", f"{int(ret):,} cr")
                with m4: st.metric("Upgraded out", f"{int(up):,} cr")
                with m5: st.metric("Downgraded out", f"{int(dn):,} cr")
                with m6: st.metric("Inflow from others", f"{int(inf):,} cr")

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)
        csv = io.StringIO()
        df_combined.to_csv(csv)
        st.download_button("⬇️ Download CSV", csv.getvalue(),
                           "nrb_data.csv", "text/csv",
                           use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 4 — ABOUT
# ══════════════════════════════════════════════════════════════════════════════

with tab_about:
    st.markdown("""
    ## ℹ️ About

    ### Matrix Logic

    ```
    ROWS    = Opening grade   →   Row sum  = Opening balance
    COLUMNS = Closing grade   →   Col sum  = Closing balance

    Example (Substandard row):
    ┌─────────────────────────────────────────────┐
    │  Opening: 430 cr                            │
    │                                             │
    │  → Good ............  0   (upgrade)         │
    │  → Watchlist ........ 120  (upgrade)        │
    │  → Substandard ...... 300  (retained) ✦     │
    │  → Doubtful .........  10  (downgrade)      │
    │  → Bad ..............   0                   │
    │                        ──                   │
    │  Row sum (Opening): 430 ✅                   │
    └─────────────────────────────────────────────┘

    Column sum = total flowing INTO that closing grade
               = Closing balance for that grade
    ```

    ### Key Identities

    | Formula | Meaning |
    |---------|---------|
    | `Σ row[i]` | Opening balance of grade i |
    | `Σ col[j]` | Closing balance of grade j |
    | `M[i][i]` | Retained in same grade |
    | `M[i][j], j < i` | Upgraded from i to j |
    | `M[i][j], j > i` | Downgraded from i to j |

    ### Color Legend

    | Color | Meaning |
    |-------|---------|
    | 🔵 Blue diagonal | Retained |
    | 🟢 Green | Upgrade (better grade) |
    | 🟡 Amber | Mild downgrade <5% of opening |
    | 🟠 Coral | Moderate downgrade 5–30% |
    | 🔴 Red | Severe downgrade >30% |
    | ⬜ Gray | No flow (zero) |

    ### NRB Grade Definitions

    | Grade | Status | Typical Overdue |
    |-------|--------|-----------------|
    | Good | Performing | Current |
    | Watchlist | Special mention | 1–3 months |
    | Substandard | Classified | 3–6 months |
    | Doubtful | Impaired | 6–12 months |
    | Bad | Loss / write-off | 12+ months |
    """)
