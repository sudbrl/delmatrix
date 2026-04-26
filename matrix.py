# ══════════════════════════════════════════════════════════════════════════════
# Loan Transition Matrix Dashboard  —  v4.0  (PDF Redesign)
# ══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import numpy as np
import io
import hashlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.ticker as mticker
from datetime import datetime
from typing import Tuple, Dict

# ── 1. PAGE CONFIGURATION ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Loan Transition Matrix",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 2. CONSTANTS & ORIGINAL HEATMAP COLORS ──────────────────────────────────
GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = len(GRADES)

plt.rcParams.update({"font.family": "DejaVu Sans", "figure.dpi": 140})

CANVAS_BG, HEADER_BG, ROW_HDR_BG = "#FAFAF8", "#E8E6DF", "#F1EFE8"
GRID_EDGE,  OUTER_EDGE            = "#D0CCC3", "#BDB7AC"
DIAG_BG,  DIAG_FG  = "#B5D4F4", "#042C53"
UPG_BG,   UPG_FG   = "#EAF3DE", "#173404"
ZERO_BG,  ZERO_FG  = "#F1EFE8", "#888780"
MILD_BG,  MILD_FG  = "#FAEEDA", "#412402"
MOD_BG,   MOD_FG   = "#F0997B", "#4A1B0C"
SEV_BG,   SEV_FG   = "#A32D2D", "#FCEBEB"
TEXT_DARK, TEXT_MID = "#2C2C2A", "#5F5E5A"

# ── 3. CSS STYLING ───────────────────────────────────────────────────────────
def load_css():
    st.markdown("""
    <style>
    :root {
        --primary: #1565C0; --primary-dark: #0D47A1;
        --success: #2E7D32; --warning: #ED6C02;
        --error: #C62828; --neutral: #7A7670;
        --bg-primary: #FAFAF9; --bg-secondary: #FFFFFF;
        --border: #E0E0E0; --text-primary: #1A1A18;
        --shadow: 0 2px 8px rgba(0,0,0,0.08); --radius: 12px;
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
        padding: 32px 24px; text-align: center; margin: 16px 0;
    }
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

# ── 5. ROBUST PARSER ─────────────────────────────────────────────────────────
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
        raise ValueError(f"Failed to read Excel file: {e}.")
    if df.empty: raise ValueError("The uploaded file is empty.")

    def norm(s): return str(s).strip().lower()
    grade_norms = [norm(g) for g in GRADES]

    header_row_idx = None
    for i in range(min(15, len(df))):
        row_vals = [norm(str(x)) for x in df.iloc[i] if pd.notna(x)]
        if sum(1 for v in row_vals if v in grade_norms) >= 4:
            header_row_idx = i; break
    if header_row_idx is None:
        raise ValueError("Could not find header row with grade names.")

    header_row = df.iloc[header_row_idx]
    col_map = {}
    for ci in range(len(header_row)):
        val = norm(str(header_row.iloc[ci]))
        if val in grade_norms and val not in col_map: col_map[val] = ci

    if len(col_map) < N: raise ValueError("Missing grade columns.")

    data_rows = {}
    for i in range(header_row_idx + 1, len(df)):
        first_val = norm(str(df.iloc[i, 0]))
        if first_val in grade_norms and first_val not in data_rows:
            data_rows[first_val] = df.iloc[i]
    if len(data_rows) < N: raise ValueError("Missing grade rows.")

    trans = np.zeros((N, N), dtype=float)
    for ri, g_from in enumerate(grade_norms):
        row_data = data_rows[g_from]
        for ci, g_to in enumerate(grade_norms):
            trans[ri, ci] = clean_numeric(row_data.iloc[col_map[g_to]])

    if trans.sum() == 0: raise ValueError("All transition values parsed as zero.")
    if np.any(trans < 0): raise ValueError("Negative values detected.")
    return trans.sum(axis=1), trans

@st.cache_data(ttl=1800, show_spinner=False)
def compute_statistics_cached(matrix_tuple: tuple, prev_tuple: tuple) -> Dict:
    trans = np.array(matrix_tuple); prev = np.array(prev_tuple)
    total_closing = float(trans.sum()); col_closing = trans.sum(axis=0)
    retained  = float(np.trace(trans))
    upgraded  = float(sum(trans[r, c] for r in range(N) for c in range(r)))
    downgraded= float(sum(trans[r, c] for r in range(N) for c in range(r+1, N)))
    return {
        "total_opening": float(prev.sum()), "total_closing": total_closing,
        "col_closing": col_closing.tolist(), "retained": retained,
        "upgraded": upgraded, "downgraded": downgraded,
        "retention_pct":   (retained   / total_closing * 100) if total_closing else 0,
        "upgrade_pct":     (upgraded   / total_closing * 100) if total_closing else 0,
        "downgrade_pct":   (downgraded / total_closing * 100) if total_closing else 0,
        "migration_rate":  (upgraded + downgraded) / total_closing * 100 if total_closing else 0,
        "concentration_ratio": col_closing.max() / total_closing * 100 if total_closing else 0,
        "grade_details": [{
            "grade": GRADES[i], "opening": float(prev[i]),
            "closing": float(col_closing[i]), "retained": float(trans[i, i]),
            "upgraded_out": float(sum(trans[i, :i])),
            "downgraded_out": float(sum(trans[i, i+1:]))
        } for i in range(N)]
    }

# ── 6. HEATMAP FUNCTIONS (Streamlit display) ─────────────────────────────────
def cell_colors(val, ri, ci, prev):
    if ri == ci: return DIAG_BG, DIAG_FG
    if val == 0: return ZERO_BG, ZERO_FG
    if ci < ri:  return UPG_BG,  UPG_FG
    pct = val / prev[ri] * 100 if prev[ri] > 0 else 0
    if pct < 5:  return MILD_BG, MILD_FG
    if pct < 30: return MOD_BG,  MOD_FG
    return SEV_BG, SEV_FG

def draw_cell(ax, x, y, w, h, bg, lines, fgs, ec=GRID_EDGE,
              fs1=9.5, fs2=8.0, w1="bold", w2="normal", ha="center"):
    ax.add_patch(mpatches.Rectangle((x, y), w, h, lw=0.8,
                                    edgecolor=ec, facecolor=bg, zorder=2))
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
    n = len(grades); col_closing = trans.sum(axis=0)
    RH, CW, HW = 0.95, 1.38, 2.45
    th = (n + 2) * RH; tw = HW + n * CW
    fig, ax = plt.subplots(figsize=(tw + .7, th + 1.8))
    fig.patch.set_facecolor(CANVAS_BG); ax.set_facecolor(CANVAS_BG)
    ax.set_aspect("equal"); ax.axis("off")

    ty = (n + 1) * RH
    draw_cell(ax, 0, ty, HW, RH, HEADER_BG,
              [f"Period: {period}", "Opening  ->  Closing"],
              [TEXT_DARK, TEXT_MID], ha="left", fs1=9.6, fs2=8.0)
    for ci, g in enumerate(grades):
        draw_cell(ax, HW + ci * CW, ty, CW, RH, HEADER_BG,
                  [g, f"Closing: {col_closing[ci]:,.1f}"],
                  [TEXT_DARK, TEXT_MID], fs1=9.3, fs2=7.8)

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
                          [f"{v:,.2f}", f"({p:.1f}%)"], [fg, fg],
                          fs1=9.5, fs2=8.0)

    ax.add_patch(mpatches.Rectangle((0, 0), tw, th, fill=False,
                                    lw=1.15, edgecolor=OUTER_EDGE, zorder=4))
    ax.set_xlim(-.08, tw + .08); ax.set_ylim(-.08, th + .1)

    legend_items = [
        (DIAG_BG, "Retained (diagonal)"), (UPG_BG, "Upgrade"),
        (MILD_BG, "Mild downgrade <5%"), (MOD_BG, "Moderate 5-30%"),
        (SEV_BG, "Severe >30%"),          (ZERO_BG, "No flow")
    ]
    patches = [mpatches.Patch(facecolor=c, edgecolor=GRID_EDGE, lw=.7, label=l)
               for c, l in legend_items]
    leg = ax.legend(handles=patches, loc="upper center",
                    bbox_to_anchor=(.5, -.07), ncol=3, fontsize=8.3,
                    frameon=True, fancybox=False, edgecolor=GRID_EDGE,
                    columnspacing=1.3, handlelength=1.5, borderpad=.6)
    leg.get_frame().set_facecolor(CANVAS_BG)
    fig.suptitle("Loan Quality Transition Matrix", fontsize=13,
                 fontweight="bold", y=.98, color=TEXT_DARK)
    plt.tight_layout(rect=(0, .05, 1, .99))
    return fig

def fig_to_bytes(fig, fmt="png", dpi=150):
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, dpi=dpi, bbox_inches="tight",
                pad_inches=0.1, facecolor=fig.get_facecolor())
    buf.seek(0); return buf.read()

def render_matrix_html(trans: np.ndarray, prev: np.ndarray) -> str:
    n = len(GRADES); col_closing = trans.sum(axis=0)
    headers = "".join(f'<th>{g}</th>' for g in GRADES)
    html = [f'<table class="matrix-table"><thead><tr>'
            f'<th style="text-align:left;padding-left:14px;">From / To</th>'
            f'{headers}<th>Opening</th></tr></thead><tbody>']
    for ri in range(n):
        cells = []
        for ci in range(n):
            v = trans[ri, ci]
            cls = "diag" if ri == ci else ("upgrade" if ci < ri else "downgrade")
            cells.append(f'<td class="{cls}">{v:,.1f}</td>')
        html.append(f'<tr><td style="text-align:left;padding-left:14px;font-weight:600;">'
                    f'{ICONS[ri]} {GRADES[ri]}</td>{"".join(cells)}'
                    f'<td style="color:var(--neutral)">{prev[ri]:,.1f}</td></tr>')
    total_cells = "".join(f'<td style="font-weight:700;color:var(--primary)">'
                          f'{col_closing[ci]:,.1f}</td>' for ci in range(n))
    html.append(f'<tr style="background:#F5F9FF;"><td style="text-align:left;'
                f'padding-left:14px;font-weight:700;">Closing</td>{total_cells}'
                f'<td style="font-weight:700;color:var(--primary)">{prev.sum():,.1f}</td>'
                f'</tr></tbody></table>')
    return "".join(html)



# ══ PDF REPORT MODULE ══
# ══════════════════════════════════════════════════════════════════════════════
# PDF REPORT REDESIGN — Drop-in replacement for build_pdf_report and helpers
# All pages: landscape 14×10 inches for comfortable content area
# Zero overlap: tight bbox calculation drives all vertical positioning
# ══════════════════════════════════════════════════════════════════════════════

import io
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
import matplotlib.ticker as mticker
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime

# ── Colour palette ─────────────────────────────────────────────────────────
PDF_PRIMARY    = "#1565C0"
PDF_SECONDARY  = "#0D47A1"
PDF_ACCENT     = "#1976D2"
PDF_LIGHT_BLUE = "#E3F2FD"
PDF_SUCCESS    = "#2E7D32"
PDF_WARNING    = "#E65100"
PDF_ERROR      = "#B71C1C"
PDF_NEUTRAL    = "#546E7A"
PDF_BG         = "#F8F9FA"
PDF_WHITE      = "#FFFFFF"
PDF_DARK       = "#1A1A1A"
PDF_BORDER     = "#CFD8DC"
PDF_LIGHT_GREEN= "#E8F5E9"
PDF_LIGHT_AMBER= "#FFF8E1"
PDF_LIGHT_RED  = "#FFEBEE"

# Heatmap colours (reused from main app)
CANVAS_BG, HEADER_BG, ROW_HDR_BG = "#FAFAF8", "#E8E6DF", "#F1EFE8"
GRID_EDGE,  OUTER_EDGE            = "#D0CCC3", "#BDB7AC"
DIAG_BG,  DIAG_FG  = "#B5D4F4", "#042C53"
UPG_BG,   UPG_FG   = "#EAF3DE", "#173404"
ZERO_BG,  ZERO_FG  = "#F1EFE8", "#888780"
MILD_BG,  MILD_FG  = "#FAEEDA", "#412402"
MOD_BG,   MOD_FG   = "#F0997B", "#4A1B0C"
SEV_BG,   SEV_FG   = "#A32D2D", "#FCEBEB"
TEXT_DARK, TEXT_MID = "#2C2C2A", "#5F5E5A"

GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = 5

GRADE_DESCRIPTIONS = {
    "Good":        "Performing loans; borrowers current on all payments. Highest credit quality.",
    "Watchlist":   "Early-warning signals: minor delays (1–3 m overdue) or adverse conditions. Enhanced monitoring required.",
    "Substandard": "Well-defined weaknesses; 3–6 m overdue. Full repayment in doubt; provisioning mandatory.",
    "Doubtful":    "Highly impaired; 6–12 m overdue. Collection in full improbable. Significant provisioning required.",
    "Bad":         "Loss loans; 12 m+ overdue. Negligible recovery prospects. Full provisioning; write-off procedures initiated.",
}

PAGE_W, PAGE_H = 14, 10   # landscape inches


# ── Utility helpers ─────────────────────────────────────────────────────────
def _rgb(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))

def _tint(hex_color: str, factor: float = 0.88):
    """Return a very light tint of a hex colour."""
    c = _rgb(hex_color)
    return tuple(min(1.0, v + factor * (1 - v)) for v in c)

def _wrap(text: str, cpl: int = 115):
    words = text.split(); lines = []; cur = ""
    for w in words:
        if len(cur) + len(w) + 1 > cpl:
            lines.append(cur.strip()); cur = w
        else:
            cur += (" " if cur else "") + w
    if cur: lines.append(cur.strip())
    return lines


# ── Page chrome (header bar + footer) ──────────────────────────────────────
def _chrome(ax, title: str, page_num: int, total_pages: int,
            period: str, generated_at: str):
    """Draw header bar, page title rule, and footer onto a full-page axes."""
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # Header strip
    ax.add_patch(mpatches.Rectangle(
        (0, 0.957), 1, 0.043,
        facecolor=_rgb(PDF_PRIMARY), edgecolor="none",
        transform=ax.transAxes, clip_on=False, zorder=5
    ))
    ax.text(0.014, 0.979, "🏦  Loan Quality Transition Matrix — Management Report",
            transform=ax.transAxes, fontsize=9, color="white",
            fontweight="bold", va="center", zorder=6)
    ax.text(0.986, 0.979, f"Period: {period}",
            transform=ax.transAxes, fontsize=8, color="#BBDEFB",
            va="center", ha="right", zorder=6)

    # Section title
    ax.text(0.014, 0.937, title,
            transform=ax.transAxes, fontsize=12, color=_rgb(PDF_DARK),
            fontweight="bold", va="top", zorder=6)
    ax.axhline(0.924, color=_rgb(PDF_BORDER), lw=0.9, xmin=0.012, xmax=0.988)

    # Footer
    ax.axhline(0.030, color=_rgb(PDF_BORDER), lw=0.8, xmin=0.012, xmax=0.988)
    ax.text(0.014, 0.016, f"Generated: {generated_at}  |  NRB Compliant  |  Confidential",
            transform=ax.transAxes, fontsize=7, color=_rgb(PDF_NEUTRAL), va="center")
    ax.text(0.986, 0.016, f"Page {page_num} of {total_pages}",
            transform=ax.transAxes, fontsize=7, color=_rgb(PDF_NEUTRAL),
            va="center", ha="right")


# ── KPI card helper ─────────────────────────────────────────────────────────
def _kpi(ax, x, y, w, h, title, value, subtitle, color_hex):
    bg = _tint(color_hex)
    c  = _rgb(color_hex)
    ax.add_patch(mpatches.FancyBboxPatch(
        (x, y), w, h, boxstyle="round,pad=0.006",
        facecolor=bg, edgecolor=c, linewidth=1.3, zorder=3
    ))
    # colour top accent bar
    ax.add_patch(mpatches.Rectangle(
        (x, y + h - 0.007), w, 0.007,
        facecolor=c, edgecolor="none", zorder=4
    ))
    mid = x + w / 2
    ax.text(mid, y + h * 0.67, value,
            ha="center", va="center", fontsize=14, fontweight="bold",
            color=_rgb(PDF_DARK), zorder=5)
    ax.text(mid, y + h * 0.36, title,
            ha="center", va="center", fontsize=7.5, color=_rgb(PDF_NEUTRAL),
            fontweight="600", zorder=5)
    ax.text(mid, y + h * 0.14, subtitle,
            ha="center", va="center", fontsize=6.5, color=c, zorder=5)


# ── Heatmap draw helpers (self-contained, no dependency on Streamlit app) ───
def _cell_colors(val, ri, ci, prev_row):
    if ri == ci: return DIAG_BG, DIAG_FG
    if val == 0: return ZERO_BG, ZERO_FG
    if ci < ri:  return UPG_BG,  UPG_FG
    pct = val / prev_row * 100 if prev_row > 0 else 0
    if pct < 5:  return MILD_BG, MILD_FG
    if pct < 30: return MOD_BG,  MOD_FG
    return SEV_BG, SEV_FG

def _draw_cell(ax, x, y, w, h, bg, lines, fgs,
               ec=GRID_EDGE, fs1=9, fs2=7.5, w1="bold", w2="normal", ha="center"):
    ax.add_patch(mpatches.Rectangle(
        (x, y), w, h, lw=0.7, edgecolor=ec, facecolor=bg, zorder=2
    ))
    tx = x + (0.10 if ha == "left" else 0.50) * w
    if len(lines) == 1:
        ax.text(tx, y + h / 2, lines[0], ha=ha, va="center",
                fontsize=fs1, color=fgs[0], fontweight=w1, zorder=3)
    else:
        ax.text(tx, y + h * .63, lines[0], ha=ha, va="center",
                fontsize=fs1, color=fgs[0], fontweight=w1, zorder=3)
        ax.text(tx, y + h * .28, lines[1], ha=ha, va="center",
                fontsize=fs2, color=fgs[1], fontweight=w2, zorder=3)


def _draw_heatmap(ax, trans, prev):
    """Draw the full colour-coded transition matrix onto ax."""
    n = len(GRADES)
    col_closing = trans.sum(axis=0)
    RH, CW, HW = 0.82, 1.18, 2.10

    th = (n + 1) * RH
    tw = HW + n * CW

    ax.set_aspect("equal")
    ax.set_xlim(-0.08, tw + 0.08)
    ax.set_ylim(-0.55, th + 0.15)
    ax.axis("off")

    # Column headers
    for ci, g in enumerate(GRADES):
        _draw_cell(ax, HW + ci * CW, n * RH, CW, RH, HEADER_BG,
                   [g, f"Cl: {col_closing[ci]:,.1f}"],
                   [TEXT_DARK, TEXT_MID], fs1=8.5, fs2=7.0)

    for ri in range(n):
        y = (n - 1 - ri) * RH
        _draw_cell(ax, 0, y, HW, RH, ROW_HDR_BG,
                   [GRADES[ri], f"Op: {prev[ri]:,.1f}"],
                   [TEXT_DARK, TEXT_MID], ha="left", fs1=8.5, fs2=7.0)
        for ci in range(n):
            v = trans[ri, ci]
            bg, fg = _cell_colors(v, ri, ci, prev[ri])
            p = v / prev[ri] * 100 if prev[ri] > 0 else 0
            if v == 0 and ri != ci:
                _draw_cell(ax, HW + ci * CW, y, CW, RH, bg,
                           ["—"], [ZERO_FG], fs1=10, w1="normal")
            else:
                _draw_cell(ax, HW + ci * CW, y, CW, RH, bg,
                           [f"{v:,.1f}", f"({p:.1f}%)"], [fg, fg],
                           fs1=8.5, fs2=7.0)

    ax.add_patch(mpatches.Rectangle(
        (0, 0), tw, th, fill=False, lw=1.2, edgecolor=OUTER_EDGE, zorder=4
    ))

    legend_items = [
        (DIAG_BG, "Retained (diagonal)"), (UPG_BG, "Upgrade"),
        (MILD_BG, "Mild downgrade <5%"),  (MOD_BG,  "Moderate 5–30%"),
        (SEV_BG,  "Severe >30%"),          (ZERO_BG, "No flow"),
    ]
    patches = [mpatches.Patch(facecolor=c, edgecolor=GRID_EDGE, lw=.6, label=l)
               for c, l in legend_items]
    ax.legend(handles=patches, loc="lower center",
              bbox_to_anchor=(tw / 2, -0.54), ncol=6,
              fontsize=7.5, frameon=True, fancybox=False,
              edgecolor=GRID_EDGE, columnspacing=1.1,
              handlelength=1.3, borderpad=0.5)


# ── Grade narrative ──────────────────────────────────────────────────────────
def _grade_narrative(grade_idx, d, trans):
    g   = GRADES[grade_idx]
    op  = d["opening"]; cl = d["closing"]
    ret = d["retained"]; up = d["upgraded_out"]; dn = d["downgraded_out"]
    net = cl - op
    ret_pct = ret / op * 100 if op else 0
    up_pct  = up  / op * 100 if op else 0
    dn_pct  = dn  / op * 100 if op else 0
    direction = "expanded" if net >= 0 else "contracted"

    parts = [
        f"The {g} category opened at NPR {op:,.1f} Cr and closed at NPR {cl:,.1f} Cr "
        f"(net {direction}: NPR {abs(net):,.1f} Cr). "
        f"Retention: {ret_pct:.1f}% (NPR {ret:,.1f} Cr) remained in-grade."
    ]
    if up > 0:
        parts.append(f"Upgrade outflow: NPR {up:,.1f} Cr ({up_pct:.1f}%) migrated to better grades.")
    if dn > 0:
        parts.append(f"Downgrade outflow: NPR {dn:,.1f} Cr ({dn_pct:.1f}%) deteriorated to lower grades.")
    inflow = float(sum(trans[r, grade_idx] for r in range(N) if r != grade_idx))
    if inflow > 0:
        parts.append(f"Inflow from other grades: NPR {inflow:,.1f} Cr.")
    if dn_pct > 20:
        parts.append(f"⚠ HIGH ALERT: Downgrade rate {dn_pct:.1f}% exceeds 20% threshold. Immediate review required.")
    elif dn_pct > 10:
        parts.append(f"⚠ CAUTION: Downgrade rate {dn_pct:.1f}% is elevated. Enhanced monitoring advised.")
    return "  ".join(parts)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE BUILDERS
# Content area: x=[0.04, 0.96], y=[0.045, 0.910]  (inside chrome)
# ══════════════════════════════════════════════════════════════════════════════

CONTENT_L = 0.035   # left edge of content area (axes coords)
CONTENT_R = 0.965   # right edge
CONTENT_T = 0.910   # top (just below title rule)
CONTENT_B = 0.042   # bottom (just above footer rule)
CONTENT_W = CONTENT_R - CONTENT_L
CONTENT_H = CONTENT_T - CONTENT_B


def _new_fig():
    fig = plt.figure(figsize=(PAGE_W, PAGE_H))
    fig.patch.set_facecolor(PDF_WHITE)
    return fig


# ── PAGE 1: Cover ────────────────────────────────────────────────────────────
def _page_cover(pdf, stats, period, generated_at):
    fig = _new_fig()
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # Gradient header (top 32%)
    for i in range(120):
        t = i / 119
        c1, c2 = _rgb(PDF_PRIMARY), _rgb(PDF_SECONDARY)
        col = tuple(c1[j]*(1-t) + c2[j]*t for j in range(3))
        ax.add_patch(mpatches.Rectangle(
            (0, 0.68 + i * 0.0027), 1, 0.0027,
            facecolor=col, edgecolor="none", zorder=2
        ))
    ax.text(0.5, 0.87, "LOAN QUALITY TRANSITION MATRIX",
            ha="center", va="center", fontsize=26, fontweight="bold",
            color="white", zorder=5)
    ax.text(0.5, 0.795, "Management Report — Portfolio Migration Analysis",
            ha="center", va="center", fontsize=14, color="#BBDEFB", zorder=5)
    ax.text(0.5, 0.727, f"Reporting Period: {period}",
            ha="center", va="center", fontsize=11, color="#E3F2FD", zorder=5)
    ax.axhline(0.68, color=_rgb(PDF_ACCENT), lw=3, zorder=6)

    # KPI row  (6 cards)
    kpi_data = [
        ("Total Opening",  f"{stats['total_opening']:,.1f} Cr",  "Portfolio Base",        PDF_PRIMARY),
        ("Total Closing",  f"{stats['total_closing']:,.1f} Cr",  "Period-End Value",       PDF_ACCENT),
        ("Retention Rate", f"{stats['retention_pct']:.1f}%",     "Stable Balances",        PDF_SUCCESS),
        ("Upgrade Rate",   f"{stats['upgrade_pct']:.1f}%",       "Credit Improvement",     PDF_SUCCESS),
        ("Downgrade Rate", f"{stats['downgrade_pct']:.1f}%",     "Credit Deterioration",   PDF_ERROR),
        ("Migration Rate", f"{stats['migration_rate']:.1f}%",    "Total Grade Movement",   PDF_WARNING),
    ]
    bw = 0.137; bh = 0.108; gap = (1 - 6*bw) / 7
    for i, (t, v, s, c) in enumerate(kpi_data):
        _kpi(ax, gap + i*(bw+gap), 0.548, bw, bh, t, v, s, c)

    # Info box
    ax.add_patch(mpatches.FancyBboxPatch(
        (0.07, 0.325), 0.86, 0.195,
        boxstyle="round,pad=0.012",
        facecolor=_rgb(PDF_LIGHT_BLUE), edgecolor=_rgb(PDF_BORDER),
        linewidth=1.0, zorder=3
    ))
    ax.text(0.5, 0.505, "Report Overview", ha="center", fontsize=10,
            fontweight="bold", color=_rgb(PDF_PRIMARY), zorder=5)
    overview = (
        f"This report presents a comprehensive analysis of loan portfolio quality migration for the period {period}. "
        "It examines transitions across five NRB-defined credit quality categories (Good, Watchlist, Substandard, "
        "Doubtful, Bad), quantifying the flow of exposures and highlighting credit risk trends. The transition "
        "matrix, key performance indicators, grade-level narratives, and supporting charts are included to support "
        "informed management decision-making and regulatory compliance."
    )
    wy = 0.472
    for line in _wrap(overview, 125):
        ax.text(0.5, wy, line, ha="center", va="top", fontsize=7.8,
                color=_rgb(PDF_DARK), zorder=5)
        wy -= 0.030

    # Contents list (two-column)
    ax.text(0.07, 0.295, "Contents:", fontsize=9.5, fontweight="bold",
            color=_rgb(PDF_PRIMARY))
    contents_l = [
        ("Page 2", "Transition Matrix Heatmap"),
        ("Page 3", "Key Performance Indicators"),
        ("Page 4", "Executive Summary & Grade Table"),
        ("Page 5", "Grade-by-Grade Narratives"),
    ]
    contents_r = [
        ("Page 6", "Opening vs Closing Distribution Chart"),
        ("Page 7", "Retention & Portfolio Composition"),
        ("Page 8", "Upgrade & Downgrade Flow Analysis"),
        ("Page 9", "Percentage Matrix & Conclusions"),
    ]
    for i, (pg, desc) in enumerate(contents_l):
        ax.text(0.08, 0.262 - i*0.024, f"{pg}  —  {desc}",
                fontsize=7.5, color=_rgb(PDF_DARK))
    for i, (pg, desc) in enumerate(contents_r):
        ax.text(0.52, 0.262 - i*0.024, f"{pg}  —  {desc}",
                fontsize=7.5, color=_rgb(PDF_DARK))

    # Footer
    ax.axhline(0.038, color=_rgb(PDF_BORDER), lw=0.8)
    ax.text(0.5, 0.020,
            f"Confidential | Internal Use Only | Generated: {generated_at} | NRB Compliant",
            ha="center", fontsize=7, color=_rgb(PDF_NEUTRAL))

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ── PAGE 2: Heatmap ──────────────────────────────────────────────────────────
def _page_heatmap(pdf, trans, prev, period, generated_at, total_pages):
    fig = _new_fig()

    # Chrome layer
    ax_chrome = fig.add_axes([0, 0, 1, 1])
    _chrome(ax_chrome, "Transition Matrix Heatmap", 2, total_pages, period, generated_at)

    # Heatmap lives in generous content area — left 75% width, full content height
    ax_heat = fig.add_axes([CONTENT_L, CONTENT_B + 0.04,
                            CONTENT_W * 0.74, CONTENT_H - 0.05])
    _draw_heatmap(ax_heat, trans, prev)

    # Right-side legend panel
    ax_leg = fig.add_axes([CONTENT_L + CONTENT_W*0.77, CONTENT_B + 0.12,
                           CONTENT_W * 0.215, CONTENT_H - 0.18])
    ax_leg.set_xlim(0, 1); ax_leg.set_ylim(0, 1); ax_leg.axis("off")

    col_cl = trans.sum(axis=0)
    ax_leg.text(0.5, 0.97, "Grade Summary", ha="center", fontsize=9.5,
                fontweight="bold", color=_rgb(PDF_PRIMARY))
    ax_leg.axhline(0.94, color=_rgb(PDF_BORDER), lw=0.8)

    grade_band = [PDF_SUCCESS, PDF_WARNING, PDF_WARNING, PDF_ERROR, PDF_ERROR]
    y = 0.88
    for i in range(N):
        ax_leg.add_patch(mpatches.Rectangle(
            (0, y - 0.005), 1, 0.090,
            facecolor=_tint(grade_band[i]), edgecolor=_rgb(grade_band[i]),
            linewidth=0.8, zorder=2
        ))
        ax_leg.text(0.06, y + 0.058, GRADES[i],
                    fontsize=8, fontweight="bold", color=_rgb(grade_band[i]))
        ax_leg.text(0.06, y + 0.030, f"Opening: {prev[i]:,.1f} Cr",
                    fontsize=7.2, color=_rgb(PDF_DARK))
        ax_leg.text(0.06, y + 0.008, f"Closing:  {col_cl[i]:,.1f} Cr",
                    fontsize=7.2, color=_rgb(PDF_DARK))
        y -= 0.108

    # Colour legend
    legend_items = [
        (DIAG_BG, "Retained"),     (UPG_BG,  "Upgrade"),
        (MILD_BG, "Mild <5%"),     (MOD_BG,  "Moderate 5–30%"),
        (SEV_BG,  "Severe >30%"),  (ZERO_BG, "No flow"),
    ]
    y = 0.31
    ax_leg.text(0.5, y + 0.03, "Colour Key", ha="center", fontsize=7.5,
                fontweight="bold", color=_rgb(PDF_NEUTRAL))
    y -= 0.005
    for bg, lbl in legend_items:
        ax_leg.add_patch(mpatches.Rectangle(
            (0.04, y - 0.012), 0.14, 0.030,
            facecolor=bg, edgecolor=GRID_EDGE, linewidth=0.6, zorder=3
        ))
        ax_leg.text(0.24, y + 0.003, lbl, fontsize=7, color=_rgb(PDF_DARK))
        y -= 0.042

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ── PAGE 3: KPI Dashboard ────────────────────────────────────────────────────
def _page_kpi(pdf, stats, period, generated_at, total_pages):
    fig = _new_fig()
    ax_chrome = fig.add_axes([0, 0, 1, 1])
    _chrome(ax_chrome, "Key Performance Indicators", 3, total_pages, period, generated_at)

    ax = fig.add_axes([CONTENT_L, CONTENT_B, CONTENT_W, CONTENT_H])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # Row 1 — 4 KPIs
    row1 = [
        ("Total Opening",  f"{stats['total_opening']:,.1f} Cr",  "Portfolio Base",      PDF_PRIMARY),
        ("Total Closing",  f"{stats['total_closing']:,.1f} Cr",  "Period-End Balance",  PDF_ACCENT),
        ("Retained",       f"{stats['retained']:,.1f} Cr",       "Same-Grade Balance",  PDF_SUCCESS),
        ("Upgraded",       f"{stats['upgraded']:,.1f} Cr",       "Moved to Better",     PDF_SUCCESS),
    ]
    # Row 2 — 4 KPIs
    row2 = [
        ("Downgraded",     f"{stats['downgraded']:,.1f} Cr",     "Moved to Worse",      PDF_ERROR),
        ("Retention Rate", f"{stats['retention_pct']:.1f}%",     "Of Closing Portfolio",PDF_PRIMARY),
        ("Migration Rate", f"{stats['migration_rate']:.1f}%",    "Total Movement",      PDF_WARNING),
        ("Concentration",  f"{stats['concentration_ratio']:.1f}%","Top-Grade Share",    PDF_NEUTRAL),
    ]
    bw = 0.225; bh = 0.130; gap = (1 - 4*bw) / 5
    for i, (t, v, s, c) in enumerate(row1):
        _kpi(ax, gap + i*(bw+gap), 0.820, bw, bh, t, v, s, c)
    for i, (t, v, s, c) in enumerate(row2):
        _kpi(ax, gap + i*(bw+gap), 0.660, bw, bh, t, v, s, c)

    # Grade summary table
    ax.text(0, 0.620, "Grade-Level Summary", fontsize=10, fontweight="bold",
            color=_rgb(PDF_PRIMARY))
    ax.axhline(0.605, color=_rgb(PDF_BORDER), lw=0.8)

    col_labels = ["Grade", "Opening (Cr)", "Closing (Cr)", "Retained (Cr)",
                  "Upgraded Out", "Downgraded Out", "Net Change"]
    col_xs     = [0, 0.13, 0.255, 0.385, 0.510, 0.640, 0.790]

    # Header row
    y_hdr = 0.570
    ax.add_patch(mpatches.Rectangle(
        (0, y_hdr - 0.006), 1, 0.038,
        facecolor=_rgb(PDF_LIGHT_BLUE), edgecolor="none", zorder=2
    ))
    for lbl, cx in zip(col_labels, col_xs):
        ax.text(cx, y_hdr + 0.013, lbl, fontsize=7.5, fontweight="bold",
                color=_rgb(PDF_PRIMARY), va="center")

    y_row = y_hdr - 0.012
    row_h = 0.072
    for i, d in enumerate(stats["grade_details"]):
        net = d["closing"] - d["opening"]
        bg  = "#F5F9FF" if i % 2 == 0 else PDF_WHITE
        ax.add_patch(mpatches.Rectangle(
            (0, y_row - row_h + 0.008), 1, row_h,
            facecolor=_rgb(bg), edgecolor="none", zorder=2
        ))
        net_col = PDF_SUCCESS if net >= 0 else PDF_ERROR
        row_vals = [
            d["grade"], f"{d['opening']:,.1f}", f"{d['closing']:,.1f}",
            f"{d['retained']:,.1f}", f"{d['upgraded_out']:,.1f}",
            f"{d['downgraded_out']:,.1f}",
            f"{'▲' if net >= 0 else '▼'} {abs(net):,.1f}"
        ]
        row_colors = [PDF_DARK]*6 + [net_col]
        for val, cx, fc in zip(row_vals, col_xs, row_colors):
            ax.text(cx, y_row - row_h/2 + 0.010, val,
                    fontsize=7.5, color=_rgb(fc), va="center")
        y_row -= row_h

    # Risk summary boxes (bottom strip)
    risk_items = [
        ("Downgrade Signal",
         "LOW RISK" if stats["downgrade_pct"] < 10 else
         "ELEVATED" if stats["downgrade_pct"] < 20 else "CRITICAL",
         PDF_SUCCESS if stats["downgrade_pct"] < 10 else
         PDF_WARNING if stats["downgrade_pct"] < 20 else PDF_ERROR),
        ("Retention Signal",
         "STRONG" if stats["retention_pct"] >= 70 else
         "MODERATE" if stats["retention_pct"] >= 50 else "WEAK",
         PDF_SUCCESS if stats["retention_pct"] >= 70 else
         PDF_WARNING if stats["retention_pct"] >= 50 else PDF_ERROR),
        ("Migration Signal",
         "LOW" if stats["migration_rate"] < 15 else
         "MODERATE" if stats["migration_rate"] < 30 else "HIGH",
         PDF_SUCCESS if stats["migration_rate"] < 15 else
         PDF_WARNING if stats["migration_rate"] < 30 else PDF_ERROR),
        ("Concentration",
         "DIVERSIFIED" if stats["concentration_ratio"] < 60 else
         "MODERATE" if stats["concentration_ratio"] < 80 else "HIGH",
         PDF_SUCCESS if stats["concentration_ratio"] < 60 else PDF_WARNING),
    ]
    bx_w = 0.225; bx_h = 0.075; bx_gap = (1 - 4*bx_w) / 5
    bx_y = y_row - 0.04
    for i, (lbl, val, col) in enumerate(risk_items):
        bx_x = bx_gap + i*(bx_w + bx_gap)
        ax.add_patch(mpatches.FancyBboxPatch(
            (bx_x, bx_y), bx_w, bx_h, boxstyle="round,pad=0.005",
            facecolor=_tint(col), edgecolor=_rgb(col), linewidth=1.2, zorder=3
        ))
        ax.text(bx_x + bx_w/2, bx_y + bx_h*0.68, val,
                ha="center", fontsize=9, fontweight="bold",
                color=_rgb(col), zorder=4)
        ax.text(bx_x + bx_w/2, bx_y + bx_h*0.28, lbl,
                ha="center", fontsize=7, color=_rgb(PDF_NEUTRAL), zorder=4)

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ── PAGE 4: Executive Summary ────────────────────────────────────────────────
def _page_exec_summary(pdf, stats, period, generated_at, total_pages):
    fig = _new_fig()
    ax_chrome = fig.add_axes([0, 0, 1, 1])
    _chrome(ax_chrome, "Executive Summary", 4, total_pages, period, generated_at)

    ax = fig.add_axes([CONTENT_L, CONTENT_B, CONTENT_W, CONTENT_H])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    net_chg   = stats["total_closing"] - stats["total_opening"]
    direction = "increased" if net_chg >= 0 else "decreased"
    risk_signal = (
        "Portfolio quality appears STABLE with manageable migration levels."
        if stats["downgrade_pct"] < 10 else
        "Portfolio quality shows ELEVATED STRESS; enhanced monitoring is required."
        if stats["downgrade_pct"] < 20 else
        "Portfolio quality shows SIGNIFICANT DETERIORATION; immediate management action is required."
    )
    risk_color = (PDF_SUCCESS if stats["downgrade_pct"] < 10 else
                  PDF_WARNING if stats["downgrade_pct"] < 20 else PDF_ERROR)

    paras = [
        (f"During the reporting period ({period}), the total loan portfolio {direction} by "
         f"NPR {abs(net_chg):,.1f} Cr — from NPR {stats['total_opening']:,.1f} Cr to "
         f"NPR {stats['total_closing']:,.1f} Cr. The overall retention rate stood at "
         f"{stats['retention_pct']:.1f}%, meaning NPR {stats['retained']:,.1f} Cr of "
         f"the closing portfolio remained in its original grade classification, "
         f"indicating the degree of portfolio stability over the period."),

        (f"Credit migration was recorded at {stats['migration_rate']:.1f}% of the closing "
         f"portfolio. Upgrades accounted for {stats['upgrade_pct']:.1f}% "
         f"(NPR {stats['upgraded']:,.1f} Cr), reflecting borrower improvement and "
         f"successful rehabilitation efforts. Downgrades accounted for "
         f"{stats['downgrade_pct']:.1f}% (NPR {stats['downgraded']:,.1f} Cr), representing "
         f"credit deterioration that requires supervisory attention and potential provisioning "
         f"adjustments."),

        (f"The portfolio concentration ratio stands at {stats['concentration_ratio']:.1f}%, "
         f"with the dominant grade category holding the highest closing balance. Management "
         f"should review grade-specific narratives on the following page for targeted "
         f"remediation, provisioning, and recovery actions aligned with NRB directives."),
    ]

    y = 0.950
    ax.text(0, y, "Portfolio Overview", fontsize=10.5, fontweight="bold",
            color=_rgb(PDF_PRIMARY)); y -= 0.032
    ax.axhline(y, color=_rgb(PDF_BORDER), lw=0.8); y -= 0.018

    for para in paras:
        for line in _wrap(para, 130):
            ax.text(0, y, line, fontsize=8.2, color=_rgb(PDF_DARK), va="top")
            y -= 0.030
        y -= 0.018

    # Risk signal banner
    y -= 0.010
    ax.add_patch(mpatches.FancyBboxPatch(
        (0, y - 0.010), 1, 0.055,
        boxstyle="round,pad=0.006",
        facecolor=_tint(risk_color, 0.82),
        edgecolor=_rgb(risk_color), linewidth=1.5, zorder=3
    ))
    ax.add_patch(mpatches.Rectangle(
        (0, y - 0.010), 0.006, 0.055,
        facecolor=_rgb(risk_color), edgecolor="none", zorder=4
    ))
    ax.text(0.018, y + 0.018, f"Risk Assessment:  {risk_signal}",
            fontsize=8.5, fontweight="bold", color=_rgb(risk_color),
            va="center", zorder=5)
    y -= 0.080

    # Management recommendations
    y -= 0.010
    ax.text(0, y, "Management Recommendations", fontsize=10.5, fontweight="bold",
            color=_rgb(PDF_PRIMARY)); y -= 0.030
    ax.axhline(y, color=_rgb(PDF_BORDER), lw=0.8); y -= 0.022

    recs = [
        "Review all Substandard, Doubtful, and Bad loans for updated collateral valuations and provisioning adequacy.",
        "Strengthen early-warning monitoring for Watchlist borrowers to prevent further grade deterioration.",
        "Validate recovery actions and legal proceedings for Bad-debt accounts; expedite write-off procedures where applicable.",
        "Assess sectoral and geographic concentrations within each grade category to identify systemic risk clusters.",
        "Benchmark current migration rates against prior periods and industry peers to identify adverse trend deviations.",
        "Ensure NRB provisioning requirements are fully met for all classified categories per current directives.",
    ]
    for j, rec in enumerate(recs):
        bx_y = y - 0.005
        ax.add_patch(mpatches.FancyBboxPatch(
            (0, bx_y - 0.016), 1, 0.044,
            boxstyle="round,pad=0.004",
            facecolor=_rgb(PDF_LIGHT_BLUE),
            edgecolor=_rgb(PDF_BORDER), linewidth=0.7, zorder=2
        ))
        ax.add_patch(mpatches.Rectangle(
            (0, bx_y - 0.016), 0.005, 0.044,
            facecolor=_rgb(PDF_PRIMARY), edgecolor="none", zorder=3
        ))
        ax.text(0.016, bx_y + 0.006, f"{j+1}.  {rec}",
                fontsize=7.8, color=_rgb(PDF_DARK), va="center", zorder=4)
        y -= 0.054

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ── PAGE 5: Grade Narratives ─────────────────────────────────────────────────
def _page_narratives(pdf, stats, trans, period, generated_at, total_pages):
    fig = _new_fig()
    ax_chrome = fig.add_axes([0, 0, 1, 1])
    _chrome(ax_chrome, "Grade-by-Grade Transition Narrative", 5, total_pages, period, generated_at)

    ax = fig.add_axes([CONTENT_L, CONTENT_B, CONTENT_W, CONTENT_H])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    grade_colors = [PDF_SUCCESS, PDF_WARNING, PDF_WARNING, PDF_ERROR, PDF_ERROR]
    bg_colors    = [PDF_LIGHT_GREEN, PDF_LIGHT_AMBER, PDF_LIGHT_AMBER,
                    PDF_LIGHT_RED, PDF_LIGHT_RED]

    # Fixed block height ensures no overlap — divide content area by 5
    block_h = 0.980 / 5   # ≈ 0.196 per grade block

    for i in range(N):
        d   = stats["grade_details"][i]
        top = 0.980 - i * block_h
        bot = top - block_h + 0.008   # small gap between blocks

        # Background panel
        ax.add_patch(mpatches.FancyBboxPatch(
            (0, bot), 1, block_h - 0.010,
            boxstyle="round,pad=0.005",
            facecolor=_rgb(bg_colors[i]),
            edgecolor=_rgb(grade_colors[i]), linewidth=1.0, zorder=2
        ))
        # Accent bar left
        ax.add_patch(mpatches.Rectangle(
            (0, bot), 0.006, block_h - 0.010,
            facecolor=_rgb(grade_colors[i]), edgecolor="none", zorder=3
        ))

        # Grade header line
        hdr_y = top - 0.030
        ax.text(0.014, hdr_y,
                f"{GRADES[i].upper()}  |  Opening: {d['opening']:,.1f} Cr"
                f"  →  Closing: {d['closing']:,.1f} Cr",
                fontsize=9, fontweight="bold",
                color=_rgb(grade_colors[i]), va="center", zorder=4)

        # Definition
        def_y = hdr_y - 0.032
        ax.text(0.014, def_y,
                f"Definition: {GRADE_DESCRIPTIONS[GRADES[i]]}",
                fontsize=7, color=_rgb(PDF_NEUTRAL), style="italic",
                va="top", zorder=4)

        # Narrative — fixed 3 lines max to stay within block
        narrative = _grade_narrative(i, d, trans)
        wrapped   = _wrap(narrative, 138)
        max_lines = 3
        txt_y = def_y - 0.034
        for line in wrapped[:max_lines]:
            ax.text(0.014, txt_y, line, fontsize=7.5,
                    color=_rgb(PDF_DARK), va="top", zorder=4)
            txt_y -= 0.028
        if len(wrapped) > max_lines:
            ax.text(0.014, txt_y, f"[{len(wrapped)-max_lines} more line(s) — see detailed analysis]",
                    fontsize=6.5, color=_rgb(PDF_NEUTRAL), va="top", style="italic", zorder=4)

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ── PAGE 6: Bar Chart ────────────────────────────────────────────────────────
def _page_bar_chart(pdf, stats, period, generated_at, total_pages):
    fig = _new_fig()
    ax_chrome = fig.add_axes([0, 0, 1, 1])
    _chrome(ax_chrome, "Opening vs Closing Distribution by Grade", 6, total_pages, period, generated_at)

    ax = fig.add_axes([CONTENT_L + 0.04, CONTENT_B + 0.06,
                       CONTENT_W - 0.06, CONTENT_H - 0.12])

    openings = [d["opening"]  for d in stats["grade_details"]]
    closings = [d["closing"]  for d in stats["grade_details"]]
    x = np.arange(N); bw = 0.35

    bars1 = ax.bar(x - bw/2, openings, bw, label="Opening",
                   color=_rgb(PDF_ACCENT), alpha=0.85,
                   edgecolor="white", linewidth=0.8)
    bars2 = ax.bar(x + bw/2, closings, bw, label="Closing",
                   color=_rgb(PDF_SUCCESS), alpha=0.85,
                   edgecolor="white", linewidth=0.8)

    all_vals = openings + closings
    max_v = max(all_vals) if all_vals else 1
    for bar in bars1:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + max_v*0.012,
                    f"{h:,.0f}", ha="center", va="bottom", fontsize=8,
                    color=_rgb(PDF_ACCENT), fontweight="600")
    for bar in bars2:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + max_v*0.012,
                    f"{h:,.0f}", ha="center", va="bottom", fontsize=8,
                    color=_rgb(PDF_SUCCESS), fontweight="600")

    ax.set_xticks(x)
    ax.set_xticklabels([f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
                       fontsize=10, fontweight="600")
    ax.set_ylabel("Amount (NPR Crore)", fontsize=10, color=_rgb(PDF_NEUTRAL))
    ax.set_title(f"Portfolio Grade Distribution — {period}",
                 fontsize=12, fontweight="bold", color=_rgb(PDF_DARK), pad=14)
    ax.legend(fontsize=10, framealpha=0.9, loc="upper right")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.spines[["top","right"]].set_visible(False)
    ax.set_facecolor(PDF_BG)
    ax.grid(axis="y", linestyle="--", alpha=0.45, color=_rgb(PDF_BORDER))
    ax.set_ylim(0, max_v * 1.15)

    # Annotation footer
    net_txt = (f"Net portfolio change: NPR {stats['total_closing']-stats['total_opening']:+,.1f} Cr"
               f"  |  Retention: {stats['retention_pct']:.1f}%"
               f"  |  Downgrade rate: {stats['downgrade_pct']:.1f}%")
    ax.text(0.5, -0.115, net_txt, ha="center", va="center", fontsize=8.5,
            color=_rgb(PDF_NEUTRAL), transform=ax.transAxes,
            bbox=dict(boxstyle="round,pad=0.4",
                      facecolor=_rgb(PDF_LIGHT_BLUE),
                      edgecolor=_rgb(PDF_BORDER), linewidth=0.8))

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ── PAGE 7: Retention + Pie ──────────────────────────────────────────────────
def _page_retention_pie(pdf, stats, period, generated_at, total_pages):
    fig = _new_fig()
    ax_chrome = fig.add_axes([0, 0, 1, 1])
    _chrome(ax_chrome, "Retention & Portfolio Composition", 7, total_pages, period, generated_at)

    gs  = gridspec.GridSpec(1, 2, figure=fig,
                            left  = CONTENT_L + 0.02,
                            right = CONTENT_R - 0.02,
                            top   = CONTENT_T - 0.02,
                            bottom= CONTENT_B + 0.06,
                            wspace=0.35)
    ax1 = fig.add_subplot(gs[0])
    ax2 = fig.add_subplot(gs[1])

    # Retention bar chart
    ret_rates = [
        d["retained"] / d["opening"] * 100 if d["opening"] > 0 else 0
        for d in stats["grade_details"]
    ]
    bar_colors = [
        _rgb(PDF_SUCCESS) if r >= 70 else
        _rgb(PDF_WARNING) if r >= 40 else
        _rgb(PDF_ERROR)
        for r in ret_rates
    ]
    bars = ax1.barh(GRADES, ret_rates, color=bar_colors,
                    edgecolor="white", linewidth=0.8, height=0.55)
    for bar, rate in zip(bars, ret_rates):
        ax1.text(min(rate + 1.5, 107), bar.get_y() + bar.get_height()/2,
                 f"{rate:.1f}%", va="center", fontsize=9.5,
                 color=_rgb(PDF_DARK), fontweight="600")

    ax1.set_xlim(0, 115)
    ax1.set_xlabel("Retention Rate (%)", fontsize=10)
    ax1.set_title("Retention Rate by Grade", fontsize=11, fontweight="bold",
                  color=_rgb(PDF_DARK), pad=10)
    ax1.axvline(70, color=_rgb(PDF_WARNING), lw=1.4,
                linestyle="--", label="70% threshold")
    ax1.legend(fontsize=9); ax1.spines[["top","right"]].set_visible(False)
    ax1.set_facecolor(PDF_BG)
    ax1.grid(axis="x", linestyle="--", alpha=0.4)

    # Closing portfolio pie
    closings = [d["closing"] for d in stats["grade_details"]]
    non_zero = [(c, g, i) for i, (c, g) in enumerate(zip(closings, GRADES)) if c > 0]
    if non_zero:
        vals, lbls, idxs = zip(*non_zero)
        pie_palette = [PDF_SUCCESS, "#66BB6A", PDF_WARNING, PDF_ERROR, "#B71C1C"]
        pie_colors  = [_rgb(pie_palette[i % len(pie_palette)]) for i in idxs]
        wedges, texts, autotexts = ax2.pie(
            vals, labels=lbls, autopct="%1.1f%%",
            colors=pie_colors, startangle=90, pctdistance=0.78,
            wedgeprops=dict(edgecolor="white", linewidth=1.5)
        )
        for t in texts:     t.set_fontsize(9.5)
        for t in autotexts: t.set_fontsize(8.5); t.set_fontweight("bold")
        ax2.set_title("Closing Portfolio Composition",
                      fontsize=11, fontweight="bold", color=_rgb(PDF_DARK), pad=10)

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ── PAGE 8: Flow Analysis ────────────────────────────────────────────────────
def _page_flow(pdf, stats, prev, period, generated_at, total_pages):
    fig = _new_fig()
    ax_chrome = fig.add_axes([0, 0, 1, 1])
    _chrome(ax_chrome, "Upgrade & Downgrade Flow Analysis", 8, total_pages, period, generated_at)

    gs  = gridspec.GridSpec(2, 1, figure=fig,
                            left  = CONTENT_L + 0.05,
                            right = CONTENT_R - 0.03,
                            top   = CONTENT_T - 0.01,
                            bottom= CONTENT_B + 0.04,
                            hspace=0.48)
    ax1 = fig.add_subplot(gs[0])
    ax2 = fig.add_subplot(gs[1])

    ups = [d["upgraded_out"]   for d in stats["grade_details"]]
    dns = [d["downgraded_out"] for d in stats["grade_details"]]
    x   = np.arange(N); bw = 0.32

    # Absolute
    ax1.bar(x-bw/2, ups, bw, label="Upgraded Out",
            color=_rgb(PDF_SUCCESS), alpha=0.85, edgecolor="white", linewidth=0.8)
    ax1.bar(x+bw/2, dns, bw, label="Downgraded Out",
            color=_rgb(PDF_ERROR),   alpha=0.85, edgecolor="white", linewidth=0.8)
    ax1.set_xticks(x); ax1.set_xticklabels(GRADES, fontsize=9.5)
    ax1.set_ylabel("NPR Crore", fontsize=9); ax1.legend(fontsize=9)
    ax1.set_title("Absolute Upgrade & Downgrade Flows",
                  fontsize=10.5, fontweight="bold", color=_rgb(PDF_DARK))
    ax1.spines[["top","right"]].set_visible(False)
    ax1.set_facecolor(PDF_BG)
    ax1.grid(axis="y", linestyle="--", alpha=0.4)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))

    # Percentage
    up_pcts = [u/o*100 if o > 0 else 0 for u, o in zip(ups, prev)]
    dn_pcts = [d/o*100 if o > 0 else 0 for d, o in zip(dns, prev)]
    ax2.bar(x-bw/2, up_pcts, bw, label="Upgrade %",
            color=_rgb(PDF_SUCCESS), alpha=0.85, edgecolor="white", linewidth=0.8)
    ax2.bar(x+bw/2, dn_pcts, bw, label="Downgrade %",
            color=_rgb(PDF_ERROR),   alpha=0.85, edgecolor="white", linewidth=0.8)
    ax2.axhline(10, color=_rgb(PDF_WARNING), lw=1.3, linestyle="--",
                label="10% Alert")
    ax2.axhline(20, color=_rgb(PDF_ERROR),   lw=1.0, linestyle=":",
                label="20% Critical")
    ax2.set_xticks(x); ax2.set_xticklabels(GRADES, fontsize=9.5)
    ax2.set_ylabel("% of Opening Balance", fontsize=9); ax2.legend(fontsize=8.5)
    ax2.set_title("Migration Rates (% of Opening Balance)",
                  fontsize=10.5, fontweight="bold", color=_rgb(PDF_DARK))
    ax2.spines[["top","right"]].set_visible(False)
    ax2.set_facecolor(PDF_BG)
    ax2.grid(axis="y", linestyle="--", alpha=0.4)
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.1f}%"))

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ── PAGE 9: Percentage Matrix + Conclusions ──────────────────────────────────
def _page_pct_and_conclusions(pdf, stats, trans, prev, period, generated_at, total_pages):
    fig = _new_fig()
    ax_chrome = fig.add_axes([0, 0, 1, 1])
    _chrome(ax_chrome, "Percentage Transition Matrix & Conclusions",
            9, total_pages, period, generated_at)

    # Split page: left = pct matrix (60%), right = conclusions (40%)
    ax_mat = fig.add_axes([CONTENT_L, CONTENT_B + 0.04,
                           CONTENT_W * 0.56, CONTENT_H - 0.06])
    ax_con = fig.add_axes([CONTENT_L + CONTENT_W*0.60, CONTENT_B + 0.02,
                           CONTENT_W * 0.40, CONTENT_H - 0.04])

    # Percentage heatmap
    pct = np.zeros((N, N))
    for r in range(N):
        for c in range(N):
            pct[r, c] = trans[r, c] / prev[r] * 100 if prev[r] > 0 else 0

    im = ax_mat.imshow(pct, cmap="RdYlGn_r", vmin=0, vmax=100, aspect="auto")
    for i in range(N):
        ax_mat.add_patch(plt.Rectangle(
            (i-0.5, i-0.5), 1, 1,
            facecolor=_rgb(DIAG_BG), edgecolor="white", lw=2, zorder=3
        ))
        ax_mat.text(i, i, f"{pct[i,i]:.1f}%", ha="center", va="center",
                    fontsize=12, fontweight="bold", color=_rgb(DIAG_FG), zorder=4)
    for r in range(N):
        for c in range(N):
            if r != c:
                v = pct[r, c]
                col = "white" if v > 55 else _rgb(PDF_DARK)
                ax_mat.text(c, r, f"{v:.1f}%", ha="center", va="center",
                            fontsize=10, color=col, zorder=4)

    ax_mat.set_xticks(range(N))
    ax_mat.set_yticks(range(N))
    ax_mat.set_xticklabels([f"→ {g}" for g in GRADES], fontsize=8.5, fontweight="600")
    ax_mat.set_yticklabels([f"{g} →" for g in GRADES], fontsize=8.5, fontweight="600")
    ax_mat.set_title("Transition Probabilities (% of opening)\nDiagonal = retained",
                     fontsize=9.5, fontweight="bold", color=_rgb(PDF_DARK), pad=10)
    cbar = plt.colorbar(im, ax=ax_mat, fraction=0.035, pad=0.02)
    cbar.set_label("% of Opening", fontsize=8)
    cbar.ax.tick_params(labelsize=7)
    ax_mat.spines[:].set_visible(False)

    # Conclusions panel
    ax_con.set_xlim(0, 1); ax_con.set_ylim(0, 1); ax_con.axis("off")

    ax_con.text(0.5, 0.980, "Key Findings", ha="center", fontsize=10,
                fontweight="bold", color=_rgb(PDF_PRIMARY))
    ax_con.axhline(0.960, color=_rgb(PDF_BORDER), lw=0.8)

    dnp = stats["downgrade_pct"]
    findings = [
        (f"Portfolio {'EXPANDED' if stats['total_closing'] >= stats['total_opening'] else 'CONTRACTED'} "
         f"by {abs((stats['total_closing']/stats['total_opening']-1)*100) if stats['total_opening'] else 0:.1f}%",
         PDF_PRIMARY),
        (f"Retention: {stats['retention_pct']:.1f}% — "
         f"{'Strong stability' if stats['retention_pct'] >= 70 else 'Moderate; review drivers'}",
         PDF_SUCCESS if stats["retention_pct"] >= 70 else PDF_WARNING),
        (f"Downgrades: {dnp:.1f}% — "
         f"{'Acceptable (<10%)' if dnp<10 else 'Elevated (10–20%)' if dnp<20 else 'CRITICAL (>20%)'}",
         PDF_SUCCESS if dnp < 10 else PDF_WARNING if dnp < 20 else PDF_ERROR),
        (f"Migration: {stats['migration_rate']:.1f}% — "
         f"{'Low' if stats['migration_rate']<15 else 'Moderate' if stats['migration_rate']<30 else 'High — investigate'}",
         PDF_SUCCESS if stats["migration_rate"] < 15 else PDF_WARNING),
        (f"Concentration: {stats['concentration_ratio']:.1f}% top grade — "
         f"{'Diversified' if stats['concentration_ratio']<60 else 'Concentrated'}",
         PDF_SUCCESS if stats["concentration_ratio"] < 60 else PDF_WARNING),
    ]

    y = 0.930
    for text, col in findings:
        ax_con.add_patch(mpatches.FancyBboxPatch(
            (0, y - 0.010), 1, 0.052,
            boxstyle="round,pad=0.004",
            facecolor=_tint(col, 0.84), edgecolor=_rgb(col),
            linewidth=0.9, zorder=2
        ))
        ax_con.add_patch(mpatches.Rectangle(
            (0, y - 0.010), 0.008, 0.052,
            facecolor=_rgb(col), edgecolor="none", zorder=3
        ))
        for li, ln in enumerate(_wrap(text, 48)[:2]):
            ax_con.text(0.020, y + 0.026 - li*0.022, ln,
                        fontsize=7.5, color=_rgb(PDF_DARK),
                        va="center", zorder=4)
        y -= 0.072

    # Disclaimer
    ax_con.add_patch(mpatches.FancyBboxPatch(
        (0, 0.005), 1, 0.055,
        boxstyle="round,pad=0.005",
        facecolor=_rgb(PDF_LIGHT_BLUE),
        edgecolor=_rgb(PDF_BORDER), linewidth=0.7, zorder=2
    ))
    ax_con.text(0.5, 0.032,
                "Disclaimer: Internal management use only.\n"
                "Figures in NPR Crore. NRB directives apply.",
                ha="center", va="center", fontsize=6.5,
                color=_rgb(PDF_NEUTRAL), style="italic", zorder=3)

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)


# ══════════════════════════════════════════════════════════════════════════════
# MASTER BUILDER — public entry point
# ══════════════════════════════════════════════════════════════════════════════
def build_pdf_report(trans: np.ndarray, prev: np.ndarray,
                     stats: dict, period: str) -> bytes:
    """Assemble all 9 pages into a landscape PDF and return bytes."""
    buf          = io.BytesIO()
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    TOTAL_PAGES  = 9

    plt.rcParams.update({"font.family": "DejaVu Sans", "figure.dpi": 140})

    with PdfPages(buf) as pdf:
        d = pdf.infodict()
        d["Title"]       = f"Loan Transition Matrix Report — {period}"
        d["Author"]      = "Loan Matrix Dashboard"
        d["Subject"]     = "Portfolio Migration Analysis"
        d["Keywords"]    = "NRB, Loan, Transition Matrix, Credit Risk"
        d["CreationDate"]= datetime.now()

        _page_cover            (pdf, stats, period, generated_at)
        _page_heatmap          (pdf, trans, prev, period, generated_at, TOTAL_PAGES)
        _page_kpi              (pdf, stats, period, generated_at, TOTAL_PAGES)
        _page_exec_summary     (pdf, stats, period, generated_at, TOTAL_PAGES)
        _page_narratives       (pdf, stats, trans, period, generated_at, TOTAL_PAGES)
        _page_bar_chart        (pdf, stats, period, generated_at, TOTAL_PAGES)
        _page_retention_pie    (pdf, stats, period, generated_at, TOTAL_PAGES)
        _page_flow             (pdf, stats, prev, period, generated_at, TOTAL_PAGES)
        _page_pct_and_conclusions(pdf, stats, trans, prev, period, generated_at, TOTAL_PAGES)

    buf.seek(0)
    return buf.read()



# ══════════════════════════════════════════════════════════════════════════════
# ── 8. SIDEBAR ───────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        '<div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;">'
        '<span style="font-size:32px;">🏦</span>'
        '<div><div style="font-weight:700;font-size:18px;">Loan Matrix Tool</div>'
        '<div style="font-size:12px;color:var(--neutral);">v4.0 | Redesigned PDF</div>'
        '</div></div>', unsafe_allow_html=True
    )
    st.divider()
    st.session_state.period = st.text_input(
        "📅 Reporting Period", value=st.session_state.period,
        placeholder="e.g., Poush 2081"
    )
    with st.expander("🖼️ Export Settings", expanded=False):
        export_dpi = st.select_slider("Chart DPI", options=[100, 150, 220, 300], value=150)
    st.divider()
    if st.session_state.generated:
        if st.button("🔄 Upload New File", type="secondary", use_container_width=True):
            st.session_state.update({
                "prev": None, "matrix": None, "generated": False,
                "stats_cache": None, "last_fig": None
            })
            st.rerun()

# ── 9. PAGE HEADER ───────────────────────────────────────────────────────────
st.markdown(
    '<div class="page-header">'
    '<h1>🏦 Loan Quality Transition Matrix</h1>'
    '<p>Upload Excel template → Analyze loan migration → '
    '<span style="color:var(--primary);font-weight:600;">Export PDF Report</span>'
    '<span class="badge">NRB Compliant</span>'
    '<span class="badge">PDF Report</span></p>'
    '</div>', unsafe_allow_html=True
)

tab_upload, tab_dash, tab_analytics, tab_report, tab_guide = st.tabs(
    ["📂 Upload", "📊 Dashboard", "📈 Analytics", "📄 PDF Report", "ℹ️ Guide"]
)

# ── TAB 1: UPLOAD ────────────────────────────────────────────────────────────
with tab_upload:
    if not st.session_state.generated:
        uploaded = st.file_uploader(
            "Upload Excel Template (.xlsx)", type=["xlsx"],
            label_visibility="collapsed"
        )
        if uploaded:
            file_bytes = uploaded.read()
            with st.spinner("🔍 Parsing & cleaning data..."):
                try:
                    prev, trans = parse_template_cached(
                        compute_file_hash(file_bytes), file_bytes
                    )
                    st.session_state.update({
                        "prev": prev, "matrix": trans,
                        "filename": uploaded.name, "upload_error": None
                    })
                    st.success(f"✅ Loaded: {uploaded.name} ({int(prev.sum()):,} cr opening)")
                except Exception as e:
                    st.error(f"❌ Parse Error: {e}")
                    st.session_state.update({"prev": None, "matrix": None})

    if st.session_state.prev is not None:
        prev_arr, trans_arr = st.session_state.prev, st.session_state.matrix
        st.divider(); st.markdown("### 📊 Data Preview")
        cols = st.columns(4)
        items = [
            (cols[0], "Opening",   prev_arr.sum(),             "Total",  "var(--primary)"),
            (cols[1], "Closing",   trans_arr.sum(),            "Total",  "var(--primary)"),
            (cols[2], "Retention", sum(trans_arr[i,i] for i in range(N))/trans_arr.sum()*100, "%", "var(--success)"),
            (cols[3], "Delta",     abs(trans_arr.sum()-prev_arr.sum()), "Chg", "var(--warning)"),
        ]
        for col, title, val, sub, color in items:
            with col:
                val_str = format_currency(val) if title != "Delta" else f"{val:.2f}"
                if title == "Retention": val_str = f"{val:.1f}%"
                st.markdown(
                    f'<div class="metric-card">'
                    f'<div class="metric-title">{title}</div>'
                    f'<div class="metric-value">{val_str}</div>'
                    f'<div class="metric-sub" style="color:{color}">{sub}</div>'
                    f'</div>', unsafe_allow_html=True
                )
        st.markdown("#### Transition Matrix Preview")
        st.markdown(render_matrix_html(trans_arr, prev_arr), unsafe_allow_html=True)
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            if st.button("🚀 Generate Dashboard", type="primary", use_container_width=True):
                stats = compute_statistics_cached(
                    tuple(map(tuple, trans_arr)), tuple(prev_arr)
                )
                st.session_state.update({"generated": True, "stats_cache": stats})
                st.rerun()
    else:
        st.markdown(
            '<div class="upload-zone">'
            '<div style="font-size:42px;">📁</div>'
            '<div style="font-weight:600;">Upload Template</div>'
            '<div style="color:var(--neutral);">Supports .xlsx with Grade Rows/Cols</div>'
            '</div>', unsafe_allow_html=True
        )

# ── TAB 2: DASHBOARD ─────────────────────────────────────────────────────────
with tab_dash:
    if not st.session_state.generated:
        st.info("👆 Upload and generate a matrix first")
    else:
        stats = st.session_state.stats_cache or compute_statistics_cached(
            tuple(map(tuple, st.session_state.matrix)), tuple(st.session_state.prev)
        )
        st.markdown("#### 🎯 Key Performance Indicators")
        kpi_cols = st.columns(4)
        kpi_data = [
            ("Total Portfolio",  format_currency(stats["total_closing"]),
             "Value at Risk",                                         "var(--primary)"),
            ("Retention Rate",   format_percentage(stats["retention_pct"]),
             f"{format_currency(stats['retained'])} Stable",         "var(--success)"),
            ("Migration Rate",   format_percentage(stats["migration_rate"]),
             f"{format_percentage(stats['upgrade_pct'])} Upgraded",  "var(--warning)"),
            ("Concentration",    format_percentage(stats["concentration_ratio"]),
             "Highest Grade",                                         "var(--neutral)"),
        ]
        for col, (title, value, sub, color) in zip(kpi_cols, kpi_data):
            col.markdown(
                f'<div class="metric-card">'
                f'<div class="metric-title">{title}</div>'
                f'<div class="metric-value">{value}</div>'
                f'<div class="metric-sub" style="color:{color}">{sub}</div>'
                f'</div>', unsafe_allow_html=True
            )

        st.divider()
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown("##### Opening vs Closing Distribution")
            st.bar_chart(
                pd.DataFrame({
                    "Grade":   [f"{ICONS[i]} {g}" for i, g in enumerate(GRADES)],
                    "Opening": [stats["grade_details"][i]["opening"]  for i in range(N)],
                    "Closing": [stats["grade_details"][i]["closing"]  for i in range(N)],
                }).set_index("Grade"),
                color=["#64B5F6", "#4DB6AC"], height=300
            )
        with col_c2:
            st.markdown("##### Retention by Grade")
            st.bar_chart(
                pd.DataFrame({
                    "Grade":  [f"{ICONS[i]} {g}" for i, g in enumerate(GRADES)],
                    "Rate %": [
                        (stats["grade_details"][i]["retained"] /
                         stats["grade_details"][i]["opening"] * 100)
                        if stats["grade_details"][i]["opening"] > 0 else 0
                        for i in range(N)
                    ],
                }).set_index("Grade"),
                color="#81C784", height=300
            )

        st.divider()
        st.markdown("#### 🔥 Transition Flow Matrix")
        with st.spinner("Rendering professional matrix..."):
            fig = build_figure(
                GRADES, st.session_state.matrix,
                st.session_state.prev, st.session_state.period
            )
            st.session_state["last_fig"] = fig
        st.pyplot(fig, use_container_width=True)

        col_exp1, col_exp2, col_exp3 = st.columns(3)
        with col_exp2:
            st.download_button(
                "⬇️ Download PNG", fig_to_bytes(fig, "png", export_dpi),
                "nrb_matrix.png", "image/png", use_container_width=True
            )
        with col_exp3:
            st.download_button(
                "⬇️ Download SVG", fig_to_bytes(fig, "svg", export_dpi),
                "nrb_matrix.svg", "image/svg+xml", use_container_width=True
            )
        plt.close(fig)

# ── TAB 3: ANALYTICS ─────────────────────────────────────────────────────────
with tab_analytics:
    if not st.session_state.generated:
        st.info("👆 Generate a matrix first")
    else:
        st.markdown("#### 📋 Transition Amounts (NPR Crore)")
        main_df = pd.DataFrame(
            np.round(st.session_state.matrix, 2),
            index  =[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES
        )
        main_df["Opening"] = np.round(st.session_state.prev, 2)
        closing_row = (list(np.round(st.session_state.matrix.sum(axis=0), 2)) +
                       [np.round(st.session_state.matrix.sum(), 2)])
        st.dataframe(
            pd.concat([
                main_df,
                pd.DataFrame([closing_row], index=["Closing"], columns=main_df.columns)
            ]).style.format("{:.2f}"),
            use_container_width=True
        )

        with st.expander("📊 View Percentages", expanded=False):
            pct_df = pd.DataFrame(
                [[(st.session_state.matrix[r, c] / st.session_state.prev[r] * 100)
                  if st.session_state.prev[r] > 0 else 0
                  for c in range(N)] for r in range(N)],
                index  =[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
                columns=GRADES
            ).round(1)
            st.dataframe(
                pct_df.style.format("{:.1f}%").background_gradient(
                    cmap="RdYlGn_r", axis=None
                ),
                use_container_width=True
            )

        st.markdown("#### 🔍 Grade Details")
        stats = st.session_state.stats_cache
        for i in range(N):
            with st.expander(f"{ICONS[i]} {GRADES[i]} Analysis"):
                d = stats["grade_details"][i]
                m1, m2, m3, m4, m5, m6 = st.columns(6)
                m1.metric("Opening",    format_currency(d["opening"]))
                m2.metric("Closing",    format_currency(d["closing"]))
                m3.metric("Retained",   format_currency(d["retained"]))
                m4.metric("Upgraded",   format_currency(d["upgraded_out"]))
                m5.metric("Downgraded", format_currency(d["downgraded_out"]))
                m6.metric("Net Flow",   format_currency(d["closing"] - d["opening"]))

# ── TAB 4: PDF REPORT ────────────────────────────────────────────────────────
with tab_report:
    if not st.session_state.generated:
        st.info("👆 Generate a matrix first, then come here to export the PDF report.")
    else:
        stats = st.session_state.stats_cache

        st.markdown("### 📄 Management PDF Report — Redesigned")
        st.markdown(
            "Click **Generate PDF** to build a full 9-page landscape management report:\n"
            "- **Cover** — portfolio KPIs, overview, contents\n"
            "- **Heatmap** — full-page colour-coded transition matrix\n"
            "- **KPI Dashboard** — 8 cards, grade table, risk signals\n"
            "- **Executive Summary** — written narrative + recommendations\n"
            "- **Grade Narratives** — per-grade analysis with risk alerts\n"
            "- **Bar Chart** — opening vs closing distribution\n"
            "- **Retention & Pie** — retention bars + portfolio composition\n"
            "- **Flow Analysis** — upgrade/downgrade absolute & % charts\n"
            "- **Pct Matrix & Conclusions** — probability matrix + findings"
        )
        st.divider()

        col_g1, col_g2, col_g3 = st.columns([1, 2, 1])
        with col_g2:
            if st.button("📄 Generate PDF Report", type="primary",
                         use_container_width=True):
                with st.spinner("Building 9-page PDF — please wait..."):
                    try:
                        pdf_bytes = build_pdf_report(
                            st.session_state.matrix,
                            st.session_state.prev,
                            stats,
                            st.session_state.period or "N/A"
                        )
                        period_safe = (st.session_state.period or "report").replace(" ", "_")
                        st.success("✅ PDF ready! Click below to download.")
                        st.download_button(
                            label="⬇️ Download Management Report PDF",
                            data=pdf_bytes,
                            file_name=f"Loan_Transition_Matrix_{period_safe}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ PDF generation failed: {e}")

# ── TAB 5: GUIDE ─────────────────────────────────────────────────────────────
with tab_guide:
    st.markdown("## ℹ️ User Guide")
    st.markdown(
        "1. **Prepare Template**: Excel with Row 1 = Grade headers, Column A = Row labels, 5×5 grid.\n"
        "2. **Upload**: Drag & drop your `.xlsx` in the **Upload** tab.\n"
        "3. **Generate**: Click *Generate Dashboard*.\n"
        "4. **PDF Report**: Go to **PDF Report** tab → *Generate PDF Report*.\n"
        "5. **Download**: PNG, SVG, or PDF exports available."
    )
    st.dataframe(
        pd.DataFrame({
            "Grade":   GRADES,
            "Status":  ["Performing", "Special Mention", "Classified", "Impaired", "Loss"],
            "Overdue": ["Current", "1-3m", "3-6m", "6-12m", "12m+"],
        }),
        use_container_width=True, hide_index=True
    )

# ── DEPENDENCY CHECK ─────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    st.error("❌ Missing: `pip install openpyxl`")
