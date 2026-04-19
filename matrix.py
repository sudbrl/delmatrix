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

# ── GitHub Dark CSS ───────────────────────────────────────────────────────────

st.markdown("""
<style>
    .stApp { background-color: #0d1117; color: #c9d1d9; }
    .main .block-container { background-color: #0d1117; padding-top: 1.5rem; }
    [data-testid="stSidebar"] {
        background-color: #161b22; border-right: 1px solid #30363d;
    }
    [data-testid="stSidebar"] * { color: #c9d1d9 !important; }
    h1,h2,h3,h4,h5,h6 {
        color: #f0f6fc !important;
        font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial !important;
    }

    .gh-card {
        background: #161b22; border: 1px solid #30363d;
        border-radius: 6px; padding: 14px 18px; margin: 6px 0;
    }
    .gh-card-title {
        color: #8b949e; font-size: 11px; font-weight: 600;
        text-transform: uppercase; letter-spacing: .5px; margin-bottom: 2px;
    }
    .gh-card-value { color: #f0f6fc; font-size: 24px; font-weight: 700; }
    .gh-green { color: #3fb950; } .gh-red { color: #f85149; }
    .gh-blue  { color: #58a6ff; } .gh-amber { color: #d29922; }

    .gh-header {
        background: linear-gradient(135deg,#161b22,#0d1117);
        border: 1px solid #30363d; border-radius: 6px;
        padding: 20px 28px; margin-bottom: 20px;
    }
    .gh-header h1 { font-size: 26px !important; margin: 0 0 6px !important; }
    .gh-header p  { color: #8b949e; font-size: 13px; margin: 0; }

    .badge {
        display: inline-block; padding: 2px 8px; border-radius: 12px;
        font-size: 11px; font-weight: 600; margin-right: 4px;
    }
    .badge-blue  { background: #1f6feb; color: #f0f6fc; }
    .badge-green { background: #238636; color: #f0f6fc; }

    .plot-box {
        background: #161b22; border: 1px solid #30363d;
        border-radius: 6px; padding: 16px; margin-top: 12px;
    }
    .gh-divider { border: none; border-top: 1px solid #21262d; margin: 16px 0; }

    .stButton>button {
        background: #238636; color: #f0f6fc;
        border: 1px solid rgba(240,246,252,.1);
        border-radius: 6px; font-weight: 600; padding: 8px 20px;
    }
    .stButton>button:hover { background: #2ea043; }
    .stDownloadButton>button {
        background: #21262d; color: #c9d1d9;
        border: 1px solid #30363d; border-radius: 6px;
    }
    div[data-baseweb="input"] input {
        background: #0d1117 !important; border: 1px solid #30363d !important;
        color: #c9d1d9 !important; border-radius: 6px !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        background: #0d1117; border-bottom: 1px solid #21262d;
    }
    .stTabs [data-baseweb="tab"] { color: #8b949e; background: transparent; }
    .stTabs [aria-selected="true"] {
        color: #f0f6fc !important; border-bottom: 2px solid #f78166 !important;
    }

    .grade-block {
        background: #161b22; border: 1px solid #30363d;
        border-radius: 8px; padding: 20px; margin-bottom: 16px;
    }
    .remaining-ok   { color: #3fb950; font-weight: 700; }
    .remaining-warn { color: #f85149; font-weight: 700; }
    .remaining-left { color: #d29922; font-weight: 700; }

    .flow-arrow {
        color: #484f58; font-size: 18px; text-align: center;
        padding-top: 30px;
    }

    ::-webkit-scrollbar { width: 8px; }
    ::-webkit-scrollbar-track { background: #0d1117; }
    ::-webkit-scrollbar-thumb { background: #30363d; border-radius: 4px; }
</style>
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
    n = len(grades)
    ct = trans.sum(axis=0)
    RH, CW, HW = 0.95, 1.28, 2.35
    th, tw = (n+1)*RH, HW+n*CW

    fig, ax = plt.subplots(figsize=(tw+.7, th+1.8))
    fig.patch.set_facecolor(CANVAS_BG)
    ax.set_facecolor(CANVAS_BG); ax.set_aspect("equal"); ax.axis("off")

    ty = n*RH
    draw_cell(ax, 0, ty, HW, RH, HEADER_BG,
              [f"From {period}", "Opening total"],
              [TEXT_DARK, TEXT_MID], ha="left", fs1=9.6)
    for ci, g in enumerate(grades):
        draw_cell(ax, HW+ci*CW, ty, CW, RH, HEADER_BG,
                  [g, f"Closing: {int(ct[ci]):,} cr"],
                  [TEXT_DARK, TEXT_MID], fs1=9.3, fs2=7.8)

    for ri in range(n):
        y = (n-1-ri)*RH
        draw_cell(ax, 0, y, HW, RH, ROW_HDR_BG,
                  [grades[ri], f"Opening: {int(prev[ri]):,} cr"],
                  [TEXT_DARK, TEXT_MID], ha="left")
        for ci in range(n):
            v = trans[ri, ci]
            bg, fg = cell_colors(v, ri, ci, prev)
            p = v/prev[ri]*100 if prev[ri] > 0 else 0
            if v == 0 and ri != ci:
                draw_cell(ax, HW+ci*CW, y, CW, RH, bg,
                          ["—"], [ZERO_FG], fs1=10, w1="normal")
            else:
                draw_cell(ax, HW+ci*CW, y, CW, RH, bg,
                          [f"{int(v):,}", f"({p:.1f}%)"],
                          [fg, fg], fs1=10, fs2=8.2)

    ax.add_patch(mpatches.Rectangle(
        (0, 0), tw, th, fill=False, lw=1.15, edgecolor=OUTER_EDGE, zorder=4))
    ax.set_xlim(-.08, tw+.08); ax.set_ylim(-.08, th+.1)

    legend = [(DIAG_BG,"Retained"),(UPG_BG,"Upgrade"),
              (MILD_BG,"Mild ↓ <5%"),(MOD_BG,"Moderate ↓ 5–30%"),
              (SEV_BG,"Severe ↓ >30%"),(ZERO_BG,"No flow")]
    patches = [mpatches.Patch(facecolor=c, edgecolor=GRID_EDGE,
               lw=.7, label=l) for c, l in legend]
    leg = ax.legend(handles=patches, loc="upper center",
                    bbox_to_anchor=(.5, -.07), ncol=3, fontsize=8.3,
                    frameon=True, fancybox=False, edgecolor=GRID_EDGE,
                    columnspacing=1.3, handlelength=1.5, borderpad=.6)
    leg.get_frame().set_facecolor(CANVAS_BG)

    fig.suptitle("NRB Loan Classification — Transition Matrix",
                 fontsize=13, fontweight="bold", y=.98, color=TEXT_DARK)
    ax.set_title("Rows = opening grade | Columns = closing grade | "
                 "Cells = NPR crore (% of opening)",
                 fontsize=8.8, color=TEXT_MID, pad=8)
    plt.tight_layout(rect=(0, .05, 1, .94))
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
    return dict(
        total_opening=int(prev.sum()), total_closing=int(tot),
        retained=int(ret), upgraded=int(up), downgraded=int(dn),
        retention_pct=ret/tot*100 if tot else 0,
        upgrade_pct=up/tot*100 if tot else 0,
        downgrade_pct=dn/tot*100 if tot else 0)


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
        <span style="color:#f0f6fc;font-size:18px;font-weight:700;">
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
    st.markdown("### 📤 Import / Export")
    exp = {"grades": GRADES, "period": st.session_state.period,
           "opening": st.session_state.prev,
           "transition": st.session_state.matrix}
    st.download_button("⬇️ Export JSON", json.dumps(exp, indent=2),
                       "nrb_data.json", "application/json",
                       use_container_width=True)

    up = st.file_uploader("⬆️ Import JSON", type=["json"])
    if up:
        try:
            d = json.load(up)
            st.session_state.prev = d["opening"]
            st.session_state.matrix = d["transition"]
            if "period" in d:
                st.session_state.period = d["period"]
            st.success("✅ Imported!"); st.rerun()
        except Exception as e:
            st.error(f"❌ {e}")


# ── Header ────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="gh-header">
    <h1>🏦 NRB Loan Classification — Transition Matrix</h1>
    <p>Enter opening balance per grade → distribute to closing grades →
       matrix builds automatically
       <span class="badge badge-blue">Simplified</span>
       <span class="badge badge-green">One-Pass Entry</span>
    </p>
</div>
""", unsafe_allow_html=True)


# ── Tabs ──────────────────────────────────────────────────────────────────────

tab_entry, tab_heatmap, tab_stats, tab_about = st.tabs(
    ["✏️ Enter Data", "📊 Heatmap", "📈 Statistics", "ℹ️ About"])


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — SIMPLIFIED DATA ENTRY
# ══════════════════════════════════════════════════════════════════════════════

with tab_entry:
    st.markdown("""
    ## 📝 Simple One-Pass Entry

    **How it works:**
    1. Enter **opening balance** for each grade
    2. For each grade, **distribute** the opening into closing categories
    3. A live **remaining tracker** ensures you allocate exactly the opening
    4. Hit **Generate** — done!
    """)

    all_valid = True

    for ri in range(N):
        st.markdown(f"""
        <div class="grade-block">
            <span style="font-size:20px;">{ICONS[ri]}</span>
            <span style="color:#f0f6fc;font-size:18px;font-weight:700;">
                {GRADES[ri]}
            </span>
        </div>
        """, unsafe_allow_html=True)

        # ── Opening Balance ───────────────────────────────────────────────

        oc1, oc2 = st.columns([1, 3])
        with oc1:
            opening = st.number_input(
                f"Opening Balance (cr)",
                min_value=0, max_value=999999,
                value=int(st.session_state.prev[ri]),
                step=10, key=f"op_{ri}",
                help=f"Total {GRADES[ri]} loans at period start")
            st.session_state.prev[ri] = opening

        # ── Distribution to Closing Grades ────────────────────────────────

        with oc2:
            st.caption(f"Distribute {opening:,} cr across closing grades:")

        dist_cols = st.columns(N + 1)

        row_vals = []
        for ci in range(N):
            with dist_cols[ci]:
                # Determine smart default: diagonal gets most
                default_val = int(st.session_state.matrix[ri][ci])

                label = GRADES[ci]
                if ci == ri:
                    label = f"✦ {GRADES[ci]} (same)"
                elif ci < ri:
                    label = f"↑ {GRADES[ci]}"
                else:
                    label = f"↓ {GRADES[ci]}"

                v = st.number_input(
                    label,
                    min_value=0,
                    max_value=max(opening, 999999),
                    value=min(default_val, opening),
                    step=5,
                    key=f"d_{ri}_{ci}")
                row_vals.append(v)

        st.session_state.matrix[ri] = row_vals

        # ── Remaining Tracker ─────────────────────────────────────────────

        allocated = sum(row_vals)
        remaining = opening - allocated

        with dist_cols[N]:
            if remaining == 0:
                cls = "remaining-ok"
                icon = "✅"
                msg = "Fully allocated"
            elif remaining > 0:
                cls = "remaining-left"
                icon = "🔸"
                msg = f"{remaining:,} cr left"
                all_valid = False
            else:
                cls = "remaining-warn"
                icon = "⚠️"
                msg = f"{abs(remaining):,} cr over"
                all_valid = False

            st.markdown(f"""
            <div style="padding-top:24px;text-align:center;">
                <div style="color:#8b949e;font-size:10px;
                            text-transform:uppercase;letter-spacing:.5px;">
                    Remaining</div>
                <div class="{cls}" style="font-size:20px;margin:4px 0;">
                    {icon} {remaining:,}
                </div>
                <div style="color:#8b949e;font-size:11px;">{msg}</div>
                <div style="color:#484f58;font-size:10px;margin-top:4px;">
                    {allocated:,} / {opening:,} cr
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Progress bar visual
            pct_used = min(allocated / opening * 100, 100) if opening > 0 else 0
            bar_color = "#3fb950" if remaining == 0 else (
                "#d29922" if remaining > 0 else "#f85149")
            st.markdown(f"""
            <div style="background:#21262d;border-radius:4px;height:6px;
                        margin-top:8px;overflow:hidden;">
                <div style="background:{bar_color};height:100%;
                            width:{pct_used}%;border-radius:4px;
                            transition:width .3s;"></div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)

    # ── Grand Summary Before Generate ─────────────────────────────────────

    total_opening  = sum(st.session_state.prev)
    trans_arr      = np.array(st.session_state.matrix, dtype=float)
    total_allocated = int(trans_arr.sum())
    total_remaining = int(total_opening - total_allocated)

    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        st.markdown(f"""
        <div class="gh-card">
            <div class="gh-card-title">Total Opening</div>
            <div class="gh-card-value">{int(total_opening):,}
                <span style="font-size:13px;color:#8b949e;"> cr</span></div>
        </div>""", unsafe_allow_html=True)
    with sc2:
        st.markdown(f"""
        <div class="gh-card">
            <div class="gh-card-title">Total Allocated</div>
            <div class="gh-card-value">{total_allocated:,}
                <span style="font-size:13px;color:#8b949e;"> cr</span></div>
        </div>""", unsafe_allow_html=True)
    with sc3:
        rem_color = "gh-green" if total_remaining == 0 else (
            "gh-amber" if total_remaining > 0 else "gh-red")
        st.markdown(f"""
        <div class="gh-card">
            <div class="gh-card-title">Unallocated</div>
            <div class="gh-card-value {rem_color}">{total_remaining:,}
                <span style="font-size:13px;color:#8b949e;"> cr</span></div>
        </div>""", unsafe_allow_html=True)

    # ── Preview Table ─────────────────────────────────────────────────────

    with st.expander("👁️ Preview Data Table", expanded=False):
        prev_arr = np.array(st.session_state.prev, dtype=float)
        df = pd.DataFrame(
            trans_arr.astype(int),
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES)
        df.insert(0, "Opening", prev_arr.astype(int))
        df["Allocated"] = trans_arr.sum(axis=1).astype(int)
        df["Remaining"] = (prev_arr - trans_arr.sum(axis=1)).astype(int)
        st.dataframe(df, use_container_width=True)

    # ── Generate Button ───────────────────────────────────────────────────

    st.markdown("")
    bc1, bc2, bc3 = st.columns([1, 2, 1])
    with bc2:
        if not all_valid:
            st.warning("⚠️ Some grades are not fully allocated. "
                       "You can still generate, but results may be partial.")

        if st.button("🚀 Generate Transition Matrix",
                      use_container_width=True, type="primary"):
            st.session_state.generated = True
            st.success("✅ Done! Switch to **📊 Heatmap** tab.")


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — HEATMAP RESULT
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
                 "NPR Crore",                      "gh-blue"),
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

        with st.spinner("🎨 Rendering…"):
            fig = build_figure(GRADES, trans_arr, prev_arr,
                               st.session_state.period)

        st.markdown('<div class="plot-box">', unsafe_allow_html=True)
        st.pyplot(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)
        d1, d2, _ = st.columns([1, 1, 2])
        with d1:
            st.download_button("⬇️ PNG", fig_to_bytes(fig, "png", export_dpi),
                               "nrb_matrix.png", "image/png",
                               use_container_width=True)
        with d2:
            st.download_button("⬇️ SVG", fig_to_bytes(fig, "svg", export_dpi),
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
        ct = trans_arr.sum(axis=0)

        st.markdown("### Transition Amounts (NPR Crore)")
        df_a = pd.DataFrame(trans_arr.astype(int),
                            index=[f"From {g}" for g in GRADES],
                            columns=[f"To {g}" for g in GRADES])
        df_a["Total"] = trans_arr.sum(axis=1).astype(int)
        st.dataframe(df_a, use_container_width=True)

        st.markdown("### Percentage of Opening (%)")
        pct = [[trans_arr[r, c]/prev_arr[r]*100 if prev_arr[r] > 0 else 0
                for c in range(N)] for r in range(N)]
        df_p = pd.DataFrame(pct,
                            index=[f"From {g}" for g in GRADES],
                            columns=[f"To {g}" for g in GRADES]).round(1)
        st.dataframe(df_p.style.background_gradient(
            cmap="RdYlGn_r", axis=None), use_container_width=True)

        st.markdown("### Grade Summary")
        summary = pd.DataFrame({
            "Grade": GRADES,
            "Opening": prev_arr.astype(int),
            "Retained": [int(trans_arr[i, i]) for i in range(N)],
            "Closing": ct.astype(int),
            "Retention %": [round(trans_arr[i, i]/prev_arr[i]*100, 1)
                            if prev_arr[i] > 0 else 0 for i in range(N)],
            "Net Change": (ct - prev_arr).astype(int),
        })
        st.dataframe(summary, use_container_width=True)

        st.markdown("### Grade Details")
        for i in range(N):
            with st.expander(f"{ICONS[i]} {GRADES[i]}"):
                ret = trans_arr[i, i]
                up  = sum(trans_arr[i, c] for c in range(i))
                dn  = sum(trans_arr[i, c] for c in range(i+1, N))
                inf = sum(trans_arr[r, i] for r in range(N) if r != i)
                m1, m2, m3, m4 = st.columns(4)
                with m1: st.metric("Retained", f"{int(ret):,} cr")
                with m2: st.metric("Upgraded out", f"{int(up):,} cr")
                with m3: st.metric("Downgraded out", f"{int(dn):,} cr")
                with m4: st.metric("Inflow", f"{int(inf):,} cr")

        st.markdown('<hr class="gh-divider">', unsafe_allow_html=True)
        csv = io.StringIO()
        df_a.to_csv(csv)
        st.download_button("⬇️ Download CSV", csv.getvalue(),
                           "nrb_data.csv", "text/csv",
                           use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 4 — ABOUT
# ══════════════════════════════════════════════════════════════════════════════

with tab_about:
    st.markdown("""
    ## ℹ️ About

    ### How Entry Works (Simplified)

    ```
    For EACH grade:
    ┌─────────────────────────────────────────┐
    │  1. Enter Opening Balance  (e.g. 430)   │
    │                                         │
    │  2. Distribute into closing grades:     │
    │     ↑ Good ........... 0                │
    │     ↑ Watchlist ...... 120              │
    │     ✦ Substandard .... 300  (retained)  │
    │     ↓ Doubtful ....... 10               │
    │     ↓ Bad ............ 0                │
    │                       ───               │
    │     Allocated:        430  ✅            │
    │     Remaining:          0               │
    └─────────────────────────────────────────┘

    Opening  = How much was in this grade
    Closing  = Sum of all rows flowing INTO this grade
               (computed automatically!)
    ```

    ### Color Legend

    | Color | Meaning |
    |-------|---------|
    | 🔵 Blue | Retained (diagonal) |
    | 🟢 Green | Upgrade |
    | 🟡 Amber | Mild downgrade <5% |
    | 🟠 Coral | Moderate 5–30% |
    | 🔴 Red | Severe >30% |
    | ⬜ Gray | No flow |

    ### NRB Grades

    | Grade | Status |
    |-------|--------|
    | Good | Performing |
    | Watchlist | Special mention |
    | Substandard | 3–6 months overdue |
    | Doubtful | Significantly impaired |
    | Bad | Loss / write-off |
    """)
