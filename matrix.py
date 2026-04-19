"""
NRB Loan Classification — Transition Matrix Heatmap (Streamlit App)
====================================================================
Interactive Streamlit app with GitHub-themed background.
Displays amount (NPR crore) + % of opening balance in each cell.
"""

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import streamlit as st
import io

# ── Page Configuration ────────────────────────────────────────────────────────

st.set_page_config(
    page_title="NRB Loan Transition Matrix",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── GitHub-Themed CSS ─────────────────────────────────────────────────────────

st.markdown("""
<style>
    /* ── GitHub dark background ── */
    .stApp {
        background-color: #0d1117;
        color: #c9d1d9;
    }

    /* ── Main content area ── */
    .main .block-container {
        background-color: #0d1117;
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background-color: #161b22;
        border-right: 1px solid #30363d;
    }

    [data-testid="stSidebar"] .stMarkdown h1,
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stMarkdown p {
        color: #c9d1d9 !important;
    }

    /* ── Headers ── */
    h1, h2, h3, h4, h5, h6 {
        color: #f0f6fc !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial !important;
    }

    /* ── GitHub-style metric cards ── */
    .github-card {
        background-color: #161b22;
        border: 1px solid #30363d;
        border-radius: 6px;
        padding: 16px 20px;
        margin: 8px 0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial;
    }

    .github-card-title {
        color: #8b949e;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }

    .github-card-value {
        color: #f0f6fc;
        font-size: 24px;
        font-weight: 700;
    }

    .github-card-sub {
        color: #3fb950;
        font-size: 12px;
        margin-top: 2px;
    }

    /* ── GitHub-style header banner ── */
    .github-header {
        background: linear-gradient(135deg, #161b22 0%, #0d1117 100%);
        border: 1px solid #30363d;
        border-radius: 6px;
        padding: 24px 32px;
        margin-bottom: 24px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial;
    }

    .github-header h1 {
        color: #f0f6fc !important;
        font-size: 28px !important;
        font-weight: 700 !important;
        margin: 0 0 8px 0 !important;
    }

    .github-header p {
        color: #8b949e;
        font-size: 14px;
        margin: 0;
    }

    /* ── Badge style ── */
    .badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        margin-right: 6px;
    }
    .badge-blue  { background-color: #1f6feb; color: #f0f6fc; }
    .badge-green { background-color: #238636; color: #f0f6fc; }
    .badge-amber { background-color: #9e6a03; color: #f0f6fc; }

    /* ── GitHub-style plot container ── */
    .plot-container {
        background-color: #161b22;
        border: 1px solid #30363d;
        border-radius: 6px;
        padding: 20px;
        margin-top: 16px;
    }

    /* ── Divider ── */
    .github-divider {
        border: none;
        border-top: 1px solid #21262d;
        margin: 20px 0;
    }

    /* ── Streamlit widget overrides ── */
    .stSlider label, .stNumberInput label,
    .stSelectbox label, .stCheckbox label {
        color: #c9d1d9 !important;
        font-size: 14px !important;
    }

    .stSlider [data-baseweb="slider"] {
        color: #1f6feb;
    }

    div[data-baseweb="input"] input {
        background-color: #0d1117 !important;
        border: 1px solid #30363d !important;
        color: #c9d1d9 !important;
        border-radius: 6px !important;
    }

    div[data-baseweb="select"] > div {
        background-color: #0d1117 !important;
        border: 1px solid #30363d !important;
        color: #c9d1d9 !important;
    }

    /* ── Buttons ── */
    .stButton > button {
        background-color: #238636;
        color: #f0f6fc;
        border: 1px solid rgba(240,246,252,0.1);
        border-radius: 6px;
        font-weight: 600;
        font-size: 14px;
        padding: 6px 16px;
        transition: background-color 0.15s ease;
    }
    .stButton > button:hover {
        background-color: #2ea043;
        border-color: rgba(240,246,252,0.2);
    }

    /* ── Download button ── */
    .stDownloadButton > button {
        background-color: #21262d;
        color: #c9d1d9;
        border: 1px solid #30363d;
        border-radius: 6px;
        font-size: 14px;
    }
    .stDownloadButton > button:hover {
        background-color: #30363d;
        color: #f0f6fc;
    }

    /* ── Info/warning boxes ── */
    .stAlert {
        background-color: #161b22 !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
        color: #c9d1d9 !important;
    }

    /* ── DataFrame / table ── */
    .stDataFrame {
        background-color: #161b22;
        border: 1px solid #30363d;
        border-radius: 6px;
    }

    /* ── Tab styling ── */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #0d1117;
        border-bottom: 1px solid #21262d;
    }
    .stTabs [data-baseweb="tab"] {
        color: #8b949e;
        background-color: transparent;
    }
    .stTabs [aria-selected="true"] {
        color: #f0f6fc !important;
        border-bottom: 2px solid #f78166 !important;
    }

    /* ── Scrollbar ── */
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: #0d1117; }
    ::-webkit-scrollbar-thumb { background: #30363d; border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #484f58; }
</style>
""", unsafe_allow_html=True)

# ── display settings ──────────────────────────────────────────────────────────

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "figure.dpi": 140,
    "savefig.dpi": 220,
})

# ── color scheme ──────────────────────────────────────────────────────────────

CANVAS_BG  = "#FAFAF8"
HEADER_BG  = "#E8E6DF"
ROW_HDR_BG = "#F1EFE8"
GRID_EDGE  = "#D0CCC3"
OUTER_EDGE = "#BDB7AC"

DIAG_BG, DIAG_FG = "#B5D4F4", "#042C53"
UPG_BG,  UPG_FG  = "#EAF3DE", "#173404"
ZERO_BG, ZERO_FG = "#F1EFE8", "#888780"
MILD_BG, MILD_FG = "#FAEEDA", "#412402"
MOD_BG,  MOD_FG  = "#F0997B", "#4A1B0C"
SEV_BG,  SEV_FG  = "#A32D2D", "#FCEBEB"

TEXT_DARK = "#2C2C2A"
TEXT_MID  = "#5F5E5A"

GRADES_DEFAULT = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]

TRANSITION_DEFAULT = np.array([
    [100,   0,   0,   0,   0],
    [ 10, 300,  10,   0,   0],
    [  0, 120, 300,  10,   0],
    [  0,   0,   0, 300,  40],
    [  0,   0,  10,   0, 200],
], dtype=float)

PREV_DEFAULT = np.array([100, 320, 430, 340, 210], dtype=float)

# ── helper functions ──────────────────────────────────────────────────────────

def cell_colors(val, ri, ci, prev):
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


def draw_cell(ax, x, y, w, h, bg, text_lines, fgs,
              edgecolor=GRID_EDGE, lw=0.8,
              fontsize1=9.5, fontsize2=8.0,
              weight1="bold", weight2="normal",
              ha="center"):
    rect = mpatches.Rectangle(
        (x, y), w, h,
        linewidth=lw, edgecolor=edgecolor,
        facecolor=bg, zorder=2,
    )
    ax.add_patch(rect)

    tx = (x + 0.10 * w) if ha == "left" else (x + 0.50 * w)

    if len(text_lines) == 1:
        ax.text(tx, y + h / 2, text_lines[0],
                ha=ha, va="center", fontsize=fontsize1,
                color=fgs[0], fontweight=weight1, zorder=3)
    else:
        ax.text(tx, y + h * 0.64, text_lines[0],
                ha=ha, va="center", fontsize=fontsize1,
                color=fgs[0], fontweight=weight1, zorder=3)
        ax.text(tx, y + h * 0.29, text_lines[1],
                ha=ha, va="center", fontsize=fontsize2,
                color=fgs[1], fontweight=weight2, zorder=3)


def build_figure(grades, transition, prev, period_label="Poush"):
    """Build and return the matplotlib figure."""
    n = len(grades)
    col_totals = transition.sum(axis=0)

    ROW_H = 0.95
    COL_W = 1.28
    HDR_W = 2.35

    total_h = (n + 1) * ROW_H
    total_w = HDR_W + n * COL_W
    fig_w = total_w + 0.7
    fig_h = total_h + 1.8

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor(CANVAS_BG)
    ax.set_facecolor(CANVAS_BG)
    ax.set_aspect("equal")
    ax.axis("off")

    # top header row
    top_y = n * ROW_H
    draw_cell(ax, 0, top_y, HDR_W, ROW_H, HEADER_BG,
              [f"From {period_label}", "Opening total"],
              [TEXT_DARK, TEXT_MID],
              ha="left", fontsize1=9.6, fontsize2=8.0)

    for ci, grade in enumerate(grades):
        x = HDR_W + ci * COL_W
        draw_cell(ax, x, top_y, COL_W, ROW_H, HEADER_BG,
                  [grade, f"Closing: {int(col_totals[ci]):,} cr"],
                  [TEXT_DARK, TEXT_MID],
                  fontsize1=9.3, fontsize2=7.8)

    # data rows
    for ri in range(n):
        y = (n - 1 - ri) * ROW_H
        draw_cell(ax, 0, y, HDR_W, ROW_H, ROW_HDR_BG,
                  [grades[ri], f"Opening: {int(prev[ri]):,} cr"],
                  [TEXT_DARK, TEXT_MID],
                  ha="left", fontsize1=9.5, fontsize2=8.0)

        for ci in range(n):
            val = transition[ri, ci]
            bg, fg = cell_colors(val, ri, ci, prev)
            x = HDR_W + ci * COL_W
            pct = val / prev[ri] * 100 if prev[ri] > 0 else 0

            if val == 0 and ri != ci:
                draw_cell(ax, x, y, COL_W, ROW_H, bg,
                          ["—"], [ZERO_FG],
                          fontsize1=10, weight1="normal")
            else:
                draw_cell(ax, x, y, COL_W, ROW_H, bg,
                          [f"{int(val):,}", f"({pct:.1f}%)"],
                          [fg, fg],
                          fontsize1=10.0, fontsize2=8.2)

    # outer border
    ax.add_patch(mpatches.Rectangle(
        (0, 0), total_w, total_h,
        fill=False, linewidth=1.15,
        edgecolor=OUTER_EDGE, zorder=4))

    ax.set_xlim(-0.08, total_w + 0.08)
    ax.set_ylim(-0.08, total_h + 0.10)

    # legend
    legend_items = [
        (DIAG_BG, "Retained (diagonal)"),
        (UPG_BG,  "Upgrade"),
        (MILD_BG, "Mild downgrade (<5%)"),
        (MOD_BG,  "Moderate downgrade (5–30%)"),
        (SEV_BG,  "Severe downgrade (>30%)"),
        (ZERO_BG, "No flow"),
    ]
    patches = [mpatches.Patch(facecolor=c, edgecolor=GRID_EDGE,
                               linewidth=0.7, label=l)
               for c, l in legend_items]

    leg = ax.legend(handles=patches, loc="upper center",
                    bbox_to_anchor=(0.5, -0.07), ncol=3,
                    fontsize=8.3, frameon=True, fancybox=False,
                    edgecolor=GRID_EDGE, columnspacing=1.3,
                    handlelength=1.5, borderpad=0.6)
    leg.get_frame().set_facecolor(CANVAS_BG)
    leg.get_frame().set_linewidth(0.8)

    fig.suptitle(
        "NRB Loan Classification — Transition Matrix",
        fontsize=13, fontweight="bold", y=0.98, color=TEXT_DARK)

    ax.set_title(
        "Rows = opening grade  |  Columns = closing grade  |  "
        "Cells = amount (NPR crore) and % of opening balance",
        fontsize=8.8, color=TEXT_MID, pad=8)

    plt.tight_layout(rect=(0, 0.05, 1, 0.94))
    return fig


def fig_to_bytes(fig, fmt="png", dpi=220):
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, dpi=dpi,
                bbox_inches="tight", pad_inches=0.15,
                facecolor=fig.get_facecolor())
    buf.seek(0)
    return buf.read()


def compute_stats(transition, prev):
    n = len(prev)
    retained = sum(transition[i, i] for i in range(n))
    upgraded = sum(transition[ri, ci]
                   for ri in range(n) for ci in range(ri)
                   if ci < ri)
    downgraded = sum(transition[ri, ci]
                     for ri in range(n) for ci in range(ri + 1, n))
    total = transition.sum()
    return {
        "total_opening":   int(prev.sum()),
        "total_closing":   int(total),
        "retained":        int(retained),
        "upgraded":        int(upgraded),
        "downgraded":      int(downgraded),
        "retention_rate":  retained / total * 100 if total else 0,
        "upgrade_rate":    upgraded  / total * 100 if total else 0,
        "downgrade_rate":  downgraded / total * 100 if total else 0,
    }


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div style="display:flex; align-items:center; gap:10px; margin-bottom:16px;">
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
        <span style="color:#f0f6fc; font-size:18px; font-weight:700;">
            NRB Matrix
        </span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### ⚙️ Configuration")

    period_label = st.text_input("📅 Period Label", value="Poush",
                                 help="Label shown in the top-left header cell")

    st.markdown("---")
    st.markdown("### 📊 Opening Balances (NPR Crore)")

    prev_inputs = []
    for i, g in enumerate(GRADES_DEFAULT):
        val = st.number_input(g, min_value=0, max_value=99999,
                              value=int(PREV_DEFAULT[i]),
                              step=10, key=f"prev_{i}")
        prev_inputs.append(float(val))
    prev = np.array(prev_inputs)

    st.markdown("---")
    st.markdown("### 🔢 Transition Matrix Values")
    st.caption("Row = From grade · Column = To grade")

    matrix_vals = []
    for ri in range(5):
        row = []
        with st.expander(f"↳ From {GRADES_DEFAULT[ri]}", expanded=(ri == 0)):
            cols = st.columns(5)
            for ci in range(5):
                v = cols[ci].number_input(
                    GRADES_DEFAULT[ci][:4],
                    min_value=0, max_value=99999,
                    value=int(TRANSITION_DEFAULT[ri, ci]),
                    step=10, key=f"t_{ri}_{ci}",
                    label_visibility="visible"
                )
                row.append(float(v))
        matrix_vals.append(row)
    transition = np.array(matrix_vals)

    st.markdown("---")
    export_dpi = st.select_slider("🖼️ Export DPI",
                                  options=[100, 150, 220, 300],
                                  value=220)

    regenerate = st.button("🔄 Regenerate Matrix", use_container_width=True)


# ── Main Area ─────────────────────────────────────────────────────────────────

# Header banner
st.markdown("""
<div class="github-header">
    <h1>🏦 NRB Loan Classification</h1>
    <p>
        Nepal Rastra Bank · Loan Transition Matrix Heatmap ·
        <span class="badge badge-blue">Interactive</span>
        <span class="badge badge-green">Exportable</span>
        <span class="badge badge-amber">NPR Crore</span>
    </p>
</div>
""", unsafe_allow_html=True)

# Tabs
tab1, tab2, tab3 = st.tabs(["📊 Matrix Heatmap", "📈 Summary Stats", "ℹ️ About"])

# ── Tab 1: Heatmap ────────────────────────────────────────────────────────────
with tab1:
    stats = compute_stats(transition, prev)

    # KPI row
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="github-card">
            <div class="github-card-title">Total Opening</div>
            <div class="github-card-value">{stats['total_opening']:,}</div>
            <div class="github-card-sub">NPR Crore</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="github-card">
            <div class="github-card-title">Retained</div>
            <div class="github-card-value">{stats['retained']:,}</div>
            <div class="github-card-sub">↔ {stats['retention_rate']:.1f}% of closing</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="github-card">
            <div class="github-card-title">Upgraded</div>
            <div class="github-card-value" style="color:#3fb950;">
                {stats['upgraded']:,}
            </div>
            <div class="github-card-sub">↑ {stats['upgrade_rate']:.1f}% of closing</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="github-card">
            <div class="github-card-title">Downgraded</div>
            <div class="github-card-value" style="color:#f85149;">
                {stats['downgraded']:,}
            </div>
            <div class="github-card-sub">↓ {stats['downgrade_rate']:.1f}% of closing</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<hr class="github-divider">', unsafe_allow_html=True)

    # Build and display figure
    with st.spinner("Rendering matrix…"):
        fig = build_figure(GRADES_DEFAULT, transition, prev, period_label)

    st.markdown('<div class="plot-container">', unsafe_allow_html=True)
    st.pyplot(fig, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Download buttons
    st.markdown('<hr class="github-divider">', unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns([1, 1, 2])
    with dl1:
        png_bytes = fig_to_bytes(fig, "png", export_dpi)
        st.download_button(
            label="⬇️ Download PNG",
            data=png_bytes,
            file_name="nrb_transition_matrix.png",
            mime="image/png",
            use_container_width=True,
        )
    with dl2:
        svg_bytes = fig_to_bytes(fig, "svg", export_dpi)
        st.download_button(
            label="⬇️ Download SVG",
            data=svg_bytes,
            file_name="nrb_transition_matrix.svg",
            mime="image/svg+xml",
            use_container_width=True,
        )

    plt.close(fig)

# ── Tab 2: Summary Stats ──────────────────────────────────────────────────────
with tab2:
    st.markdown("### Portfolio Flow Summary")

    import pandas as pd

    col_totals = transition.sum(axis=0)
    row_totals = transition.sum(axis=1)

    df_matrix = pd.DataFrame(
        transition.astype(int),
        index=[f"From {g}" for g in GRADES_DEFAULT],
        columns=[f"To {g}" for g in GRADES_DEFAULT],
    )
    df_matrix["Row Total"] = row_totals.astype(int)

    st.markdown("#### Transition Matrix (NPR Crore)")
    st.dataframe(df_matrix, use_container_width=True)

    st.markdown("#### Percentage of Opening Balance (%)")
    pct_matrix = pd.DataFrame(
        [[transition[ri, ci] / prev[ri] * 100
          if prev[ri] > 0 else 0
          for ci in range(5)]
         for ri in range(5)],
        index=[f"From {g}" for g in GRADES_DEFAULT],
        columns=[f"To {g}" for g in GRADES_DEFAULT],
    ).round(1)
    st.dataframe(pct_matrix.style.background_gradient(
        cmap="RdYlGn_r", axis=None), use_container_width=True)

    st.markdown("#### Grade-level Summary")
    summary_data = {
        "Grade":         GRADES_DEFAULT,
        "Opening (cr)":  prev.astype(int).tolist(),
        "Retained (cr)": [int(transition[i, i]) for i in range(5)],
        "Closing (cr)":  col_totals.astype(int).tolist(),
        "Retention %":   [round(transition[i, i] / prev[i] * 100, 1)
                          if prev[i] > 0 else 0 for i in range(5)],
    }
    st.dataframe(pd.DataFrame(summary_data), use_container_width=True)

# ── Tab 3: About ──────────────────────────────────────────────────────────────
with tab3:
    st.markdown("""
    ### About This App

    This interactive tool visualises **Nepal Rastra Bank (NRB)** loan classification
    transition matrices — showing how loan portfolios migrate between risk grades
    from one period to the next.

    ---

    #### 📐 How to Read the Matrix

    | Element | Meaning |
    |---|---|
    | **Row** | Opening grade (where the loan was classified) |
    | **Column** | Closing grade (where it ended up) |
    | **Diagonal cell** | Loan **retained** in the same grade |
    | **Cell left of diagonal** | Loan **upgraded** (improved) |
    | **Cell right of diagonal** | Loan **downgraded** (deteriorated) |
    | **Top number** | Amount in NPR Crore |
    | **Bottom number** | % of opening balance for that row |

    ---

    #### 🎨 Color Guide

    | Color | Meaning |
    |---|---|
    | 🔵 Blue | Retained in same grade |
    | 🟢 Green | Upgraded (improvement) |
    | 🟡 Amber | Mild downgrade < 5% |
    | 🟠 Coral | Moderate downgrade 5–30% |
    | 🔴 Red | Severe downgrade > 30% |
    | ⬜ Gray | No flow |

    ---

    #### 🏦 NRB Loan Grades

    | Grade | Description |
    |---|---|
    | **Good** | Pass / Standard loans |
    | **Watchlist** | Special mention / Watch |
    | **Substandard** | Impaired, 3–6 months overdue |
    | **Doubtful** | Significantly impaired |
    | **Bad** | Loss / Write-off candidates |

    ---

    > Built with **Streamlit** · Data in **NPR Crore**
    """)

    st.markdown("""
    <div style="margin-top:24px; padding:16px;
                background:#161b22; border:1px solid #30363d;
                border-radius:6px; font-size:13px; color:#8b949e;">
        <strong style="color:#f0f6fc;">📦 Stack</strong><br>
        Python · Streamlit · Matplotlib · NumPy · Pandas
    </div>
    """, unsafe_allow_html=True)