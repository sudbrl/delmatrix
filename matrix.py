# ══════════════════════════════════════════════════════════════════════════════
# Loan Transition Matrix Dashboard
# Original Heatmap Restored + Robust Excel Parser + PDF Report Generator
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

# PDF Report Color Palette
PDF_PRIMARY    = "#1565C0"
PDF_SECONDARY  = "#0D47A1"
PDF_ACCENT     = "#1976D2"
PDF_LIGHT_BLUE = "#E3F2FD"
PDF_SUCCESS    = "#2E7D32"
PDF_WARNING    = "#ED6C02"
PDF_ERROR      = "#C62828"
PDF_NEUTRAL    = "#546E7A"
PDF_BG         = "#FAFAFA"
PDF_WHITE      = "#FFFFFF"
PDF_DARK       = "#1A1A18"
PDF_BORDER     = "#CFD8DC"

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

# ── 6. HEATMAP FUNCTIONS ─────────────────────────────────────────────────────
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

# ══════════════════════════════════════════════════════════════════════════════
# ── 7. PDF REPORT GENERATOR ──────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

# ── Transition narrative descriptions ────────────────────────────────────────
GRADE_DESCRIPTIONS = {
    "Good": (
        "Performing loans where borrowers are current on all payments. "
        "These represent the highest credit quality in the portfolio."
    ),
    "Watchlist": (
        "Loans showing early warning signals such as minor payment delays "
        "(1-3 months overdue) or adverse business conditions. "
        "Enhanced monitoring is required."
    ),
    "Substandard": (
        "Classified loans with well-defined weaknesses, typically 3-6 months "
        "overdue. Full repayment is in doubt; provisioning is mandatory."
    ),
    "Doubtful": (
        "Highly impaired loans (6-12 months overdue) where collection in full "
        "is improbable. Significant provisioning and recovery action required."
    ),
    "Bad": (
        "Loss loans (12+ months overdue) with negligible recovery prospects. "
        "Full provisioning applies and write-off procedures are initiated."
    ),
}

def _narrative_for_grade(grade_idx: int, d: dict, stats: dict, trans: np.ndarray) -> str:
    """Generate a plain-text transition narrative for one grade row."""
    g      = GRADES[grade_idx]
    op     = d["opening"]; cl = d["closing"]
    ret    = d["retained"]; up = d["upgraded_out"]; dn = d["downgraded_out"]
    net    = cl - op
    ret_pct= ret / op * 100 if op else 0
    up_pct = up  / op * 100 if op else 0
    dn_pct = dn  / op * 100 if op else 0
    net_dir= "improved" if net >= 0 else "contracted"

    lines = []
    lines.append(
        f"The {g} category opened the period with NPR {op:,.1f} Cr and "
        f"closed at NPR {cl:,.1f} Cr, a net {net_dir} of NPR {abs(net):,.1f} Cr."
    )
    lines.append(
        f"Retention: NPR {ret:,.1f} Cr ({ret_pct:.1f}%) of opening balances "
        f"remained in the same grade, reflecting portfolio stability."
    )
    if up > 0:
        lines.append(
            f"Upgrade flow: NPR {up:,.1f} Cr ({up_pct:.1f}%) migrated to "
            f"better-quality grades, signalling credit improvement."
        )
    if dn > 0:
        lines.append(
            f"Downgrade flow: NPR {dn:,.1f} Cr ({dn_pct:.1f}%) deteriorated "
            f"to lower-quality grades, requiring supervisory attention."
        )

    # Inflow from other grades
    inflow = float(sum(trans[r, grade_idx] for r in range(N) if r != grade_idx))
    if inflow > 0:
        lines.append(
            f"Inflow: NPR {inflow:,.1f} Cr migrated into {g} from other "
            f"categories during the period."
        )

    if dn_pct > 20:
        lines.append(
            f"⚠ HIGH ALERT: Downgrade rate of {dn_pct:.1f}% exceeds the 20% "
            f"risk threshold. Immediate portfolio review recommended."
        )
    elif dn_pct > 10:
        lines.append(
            f"⚠ CAUTION: Downgrade rate of {dn_pct:.1f}% is elevated. "
            f"Enhanced monitoring advised."
        )

    return "  ".join(lines)

# ── Low-level PDF drawing helpers ────────────────────────────────────────────
def _hex_to_rgb(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))

def _draw_page_frame(ax, title: str, page_num: int, total_pages: int,
                     period: str, generated_at: str):
    """Draw page border, header bar, footer."""
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # Outer border
    for spine in ["top", "bottom", "left", "right"]:
        ax.spines[spine].set_visible(False)

    # Header bar
    ax.add_patch(mpatches.FancyBboxPatch(
        (0, 0.955), 1, 0.045, boxstyle="square,pad=0",
        facecolor=_hex_to_rgb(PDF_PRIMARY), edgecolor="none",
        transform=ax.transAxes, clip_on=False, zorder=5
    ))
    ax.text(0.012, 0.978, "🏦  Loan Quality Transition Matrix — Management Report",
            transform=ax.transAxes, fontsize=8.5, color="white",
            fontweight="bold", va="center", zorder=6)
    ax.text(0.988, 0.978, f"Period: {period}",
            transform=ax.transAxes, fontsize=7.5, color="#BBDEFB",
            va="center", ha="right", zorder=6)

    # Page title
    ax.text(0.012, 0.938, title,
            transform=ax.transAxes, fontsize=11, color=_hex_to_rgb(PDF_DARK),
            fontweight="bold", va="top", zorder=6)
    ax.axhline(0.926, color=_hex_to_rgb(PDF_BORDER), lw=0.8, xmin=0, xmax=1)

    # Footer
    ax.axhline(0.032, color=_hex_to_rgb(PDF_BORDER), lw=0.8, xmin=0, xmax=1)
    ax.text(0.012, 0.016, f"Generated: {generated_at}  |  NRB Compliant",
            transform=ax.transAxes, fontsize=6.5, color=_hex_to_rgb(PDF_NEUTRAL),
            va="center")
    ax.text(0.988, 0.016, f"Page {page_num} / {total_pages}",
            transform=ax.transAxes, fontsize=6.5, color=_hex_to_rgb(PDF_NEUTRAL),
            va="center", ha="right")

def _kpi_box(ax, x, y, w, h, title, value, subtitle, color_hex):
    """Draw a single KPI card on the axes."""
    c = _hex_to_rgb(color_hex)
    bg = tuple(min(1, v + 0.88 * (1 - v)) for v in c)   # very light tint
    ax.add_patch(mpatches.FancyBboxPatch(
        (x, y), w, h, boxstyle="round,pad=0.008",
        facecolor=bg, edgecolor=c, linewidth=1.2, zorder=3
    ))
    ax.add_patch(mpatches.Rectangle(
        (x, y + h - 0.006), w, 0.006,
        facecolor=c, edgecolor="none", zorder=4
    ))
    ax.text(x + w / 2, y + h * 0.72, value,
            ha="center", va="center", fontsize=13, fontweight="bold",
            color=_hex_to_rgb(PDF_DARK), zorder=5)
    ax.text(x + w / 2, y + h * 0.38, title,
            ha="center", va="center", fontsize=7, color=_hex_to_rgb(PDF_NEUTRAL),
            fontweight="600", zorder=5)
    ax.text(x + w / 2, y + h * 0.14, subtitle,
            ha="center", va="center", fontsize=6, color=c, zorder=5)

def _wrap_text(text: str, chars_per_line: int = 100):
    """Wrap text into lines of ≤ chars_per_line characters."""
    words = text.split(); lines = []; cur = ""
    for w in words:
        if len(cur) + len(w) + 1 > chars_per_line:
            lines.append(cur.strip()); cur = w
        else:
            cur += (" " if cur else "") + w
    if cur: lines.append(cur.strip())
    return lines

# ── Page 1 — Cover ────────────────────────────────────────────────────────────
def _page_cover(pdf: PdfPages, stats: dict, period: str, generated_at: str):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)
    ax  = fig.add_axes([0, 0, 1, 1]); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # Full-width gradient header strip
    grad_colors = [_hex_to_rgb(PDF_PRIMARY), _hex_to_rgb(PDF_SECONDARY)]
    for i in range(100):
        t = i / 99
        c = tuple(grad_colors[0][j] * (1 - t) + grad_colors[1][j] * t for j in range(3))
        ax.add_patch(mpatches.Rectangle((0, 0.72 + i * 0.0028), 1, 0.0028,
                                        facecolor=c, edgecolor="none", zorder=2))

    ax.text(0.5, 0.88, "LOAN QUALITY TRANSITION MATRIX",
            ha="center", va="center", fontsize=22, fontweight="bold",
            color="white", zorder=5)
    ax.text(0.5, 0.80, "Management Report — Portfolio Migration Analysis",
            ha="center", va="center", fontsize=13, color="#BBDEFB", zorder=5)
    ax.text(0.5, 0.74, f"Reporting Period: {period}",
            ha="center", va="center", fontsize=10, color="#E3F2FD", zorder=5)

    # Decorative accent line
    ax.axhline(0.72, color=_hex_to_rgb(PDF_ACCENT), lw=3, zorder=6)

    # Summary KPI boxes
    kpi_data = [
        ("Total Opening", f"{stats['total_opening']:,.1f} Cr",  "Portfolio Base",    PDF_PRIMARY),
        ("Total Closing", f"{stats['total_closing']:,.1f} Cr",  "Period End Value",  PDF_ACCENT),
        ("Retention Rate",f"{stats['retention_pct']:.1f}%",     "Stable Balances",   PDF_SUCCESS),
        ("Upgrade Rate",  f"{stats['upgrade_pct']:.1f}%",       "Credit Improvement",PDF_SUCCESS),
        ("Downgrade Rate",f"{stats['downgrade_pct']:.1f}%",     "Credit Deterioration",PDF_ERROR),
        ("Migration Rate",f"{stats['migration_rate']:.1f}%",    "Total Movement",    PDF_WARNING),
    ]
    bw = 0.14; bh = 0.11; gap = (1 - 6 * bw) / 7
    for i, (title, val, sub, col) in enumerate(kpi_data):
        _kpi_box(ax, gap + i * (bw + gap), 0.555, bw, bh, title, val, sub, col)

    # Report info box
    ax.add_patch(mpatches.FancyBboxPatch(
        (0.1, 0.32), 0.8, 0.19, boxstyle="round,pad=0.01",
        facecolor=_hex_to_rgb(PDF_LIGHT_BLUE), edgecolor=_hex_to_rgb(PDF_BORDER),
        linewidth=1, zorder=3
    ))
    ax.text(0.5, 0.495, "Report Overview",
            ha="center", va="center", fontsize=9, fontweight="bold",
            color=_hex_to_rgb(PDF_PRIMARY), zorder=5)
    overview = (
        "This report presents a comprehensive analysis of loan portfolio quality migration "
        f"for the period {period}. It examines transitions across five NRB-defined credit "
        "quality categories (Good, Watchlist, Substandard, Doubtful, Bad), quantifying the "
        "flow of exposures and highlighting credit risk trends. The transition matrix, key "
        "performance indicators, grade-level narratives, and supporting charts are included "
        "to support informed management decision-making and regulatory compliance."
    )
    wrapped = _wrap_text(overview, 110)
    for j, line in enumerate(wrapped):
        ax.text(0.5, 0.455 - j * 0.028, line,
                ha="center", va="center", fontsize=7.2,
                color=_hex_to_rgb(PDF_DARK), zorder=5)

    # Contents table
    ax.text(0.1, 0.275, "Contents:",
            fontsize=8.5, fontweight="bold", color=_hex_to_rgb(PDF_PRIMARY))
    contents = [
        ("Page 2", "Transition Matrix Heatmap"),
        ("Page 3", "Key Performance Indicators & Executive Summary"),
        ("Page 4", "Grade-by-Grade Transition Narrative"),
        ("Page 5", "Opening vs Closing Distribution Chart"),
        ("Page 6", "Retention & Migration Analysis Chart"),
        ("Page 7", "Flow Analysis — Upgrades & Downgrades"),
        ("Page 8", "Percentage Transition Matrix"),
        ("Page 9", "Risk Concentration & Conclusions"),
    ]
    for i, (pg, desc) in enumerate(contents):
        ax.text(0.12, 0.245 - i * 0.022, f"{pg}  —  {desc}",
                fontsize=7.2, color=_hex_to_rgb(PDF_DARK))

    # Footer
    ax.axhline(0.04, color=_hex_to_rgb(PDF_BORDER), lw=0.8)
    ax.text(0.5, 0.022,
            f"Confidential | Internal Use Only | Generated: {generated_at} | NRB Compliant",
            ha="center", va="center", fontsize=6.5, color=_hex_to_rgb(PDF_NEUTRAL))

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Page 2 — Heatmap ─────────────────────────────────────────────────────────
def _page_heatmap(pdf: PdfPages, trans: np.ndarray, prev: np.ndarray,
                  period: str, generated_at: str, total_pages: int):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)

    # Frame axes (full page)
    ax_frame = fig.add_axes([0, 0, 1, 1])
    _draw_page_frame(ax_frame, "Transition Matrix Heatmap", 2, total_pages,
                     period, generated_at)

    # Embed heatmap
    ax_heat = fig.add_axes([0.04, 0.09, 0.92, 0.82])
    ax_heat.axis("off")

    n = len(GRADES); col_closing = trans.sum(axis=0)
    RH, CW, HW = 0.85, 1.25, 2.2
    th = (n + 1) * RH; tw = HW + n * CW

    ax_heat.set_aspect("equal")
    ax_heat.set_xlim(-0.1, tw + 0.1); ax_heat.set_ylim(-0.5, th + 0.3)

    # Column headers
    for ci, g in enumerate(GRADES):
        draw_cell(ax_heat, HW + ci * CW, n * RH, CW, RH, HEADER_BG,
                  [g, f"Cl: {col_closing[ci]:,.1f}"],
                  [TEXT_DARK, TEXT_MID], fs1=8.5, fs2=7.0)

    for ri in range(n):
        y = (n - 1 - ri) * RH
        draw_cell(ax_heat, 0, y, HW, RH, ROW_HDR_BG,
                  [GRADES[ri], f"Op: {prev[ri]:,.1f}"],
                  [TEXT_DARK, TEXT_MID], ha="left", fs1=8.5, fs2=7.0)
        for ci in range(n):
            v = trans[ri, ci]
            bg, fg = cell_colors(v, ri, ci, prev)
            p = v / prev[ri] * 100 if prev[ri] > 0 else 0
            if v == 0 and ri != ci:
                draw_cell(ax_heat, HW + ci * CW, y, CW, RH, bg,
                          ["—"], [ZERO_FG], fs1=9, w1="normal")
            else:
                draw_cell(ax_heat, HW + ci * CW, y, CW, RH, bg,
                          [f"{v:,.1f}", f"({p:.1f}%)"], [fg, fg],
                          fs1=8.5, fs2=7.0)

    ax_heat.add_patch(mpatches.Rectangle((0, 0), tw, th,
                                         fill=False, lw=1.2,
                                         edgecolor=OUTER_EDGE, zorder=4))

    legend_items = [
        (DIAG_BG, "Retained"), (UPG_BG, "Upgrade"),
        (MILD_BG, "Mild dngrade<5%"), (MOD_BG, "Moderate 5-30%"),
        (SEV_BG, "Severe >30%"),      (ZERO_BG, "No flow"),
    ]
    patches = [mpatches.Patch(facecolor=c, edgecolor=GRID_EDGE, lw=.6, label=l)
               for c, l in legend_items]
    ax_heat.legend(handles=patches, loc="lower center",
                   bbox_to_anchor=(tw / 2, -0.48), ncol=6,
                   fontsize=7, frameon=True, fancybox=False,
                   edgecolor=GRID_EDGE, columnspacing=1.0,
                   handlelength=1.3, borderpad=0.5)

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Page 3 — KPI + Executive Summary ─────────────────────────────────────────
def _page_kpi_summary(pdf: PdfPages, stats: dict, trans: np.ndarray,
                      prev: np.ndarray, period: str,
                      generated_at: str, total_pages: int):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)
    ax_frame = fig.add_axes([0, 0, 1, 1])
    _draw_page_frame(ax_frame, "Key Performance Indicators & Executive Summary",
                     3, total_pages, period, generated_at)

    ax = fig.add_axes([0.04, 0.08, 0.92, 0.83])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # KPI row 1
    kpi_row1 = [
        ("Total Opening",  f"{stats['total_opening']:,.1f} Cr",  "Base Portfolio",     PDF_PRIMARY),
        ("Total Closing",  f"{stats['total_closing']:,.1f} Cr",  "Period-End Balance",  PDF_ACCENT),
        ("Retained",       f"{stats['retained']:,.1f} Cr",       "Same-Grade Balance",  PDF_SUCCESS),
        ("Upgraded",       f"{stats['upgraded']:,.1f} Cr",       "Moved to Better Grade",PDF_SUCCESS),
    ]
    kpi_row2 = [
        ("Downgraded",     f"{stats['downgraded']:,.1f} Cr",     "Moved to Worse Grade", PDF_ERROR),
        ("Retention Rate", f"{stats['retention_pct']:.1f}%",     "Of Closing Portfolio", PDF_PRIMARY),
        ("Migration Rate", f"{stats['migration_rate']:.1f}%",    "Total Grade Movement", PDF_WARNING),
        ("Concentration",  f"{stats['concentration_ratio']:.1f}%","Top-Grade Share",     PDF_NEUTRAL),
    ]
    bw = 0.22; bh = 0.115; gap = (1 - 4 * bw) / 5
    for i, (t, v, s, c) in enumerate(kpi_row1):
        _kpi_box(ax, gap + i * (bw + gap), 0.835, bw, bh, t, v, s, c)
    for i, (t, v, s, c) in enumerate(kpi_row2):
        _kpi_box(ax, gap + i * (bw + gap), 0.700, bw, bh, t, v, s, c)

    # Executive Summary text
    ax.text(0, 0.658, "Executive Summary", fontsize=10, fontweight="bold",
            color=_hex_to_rgb(PDF_PRIMARY))
    ax.axhline(0.645, color=_hex_to_rgb(PDF_BORDER), lw=0.7)

    net_chg   = stats["total_closing"] - stats["total_opening"]
    direction = "increased" if net_chg >= 0 else "decreased"
    risk_signal = (
        "Portfolio quality appears STABLE with manageable migration."
        if stats["downgrade_pct"] < 10 else
        "Portfolio quality shows ELEVATED STRESS; enhanced monitoring required."
        if stats["downgrade_pct"] < 20 else
        "Portfolio quality shows SIGNIFICANT DETERIORATION; immediate action required."
    )

    exec_text = (
        f"During the reporting period ({period}), the total loan portfolio "
        f"{direction} by NPR {abs(net_chg):,.1f} Cr, from NPR "
        f"{stats['total_opening']:,.1f} Cr to NPR {stats['total_closing']:,.1f} Cr. "
        f"The overall retention rate stood at {stats['retention_pct']:.1f}%, indicating "
        f"that {stats['retained']:,.1f} Cr of the closing portfolio remained in its "
        f"original grade classification. "
        f"\n\n"
        f"Credit migration was recorded at {stats['migration_rate']:.1f}% of the "
        f"closing portfolio. Upgrades accounted for {stats['upgrade_pct']:.1f}% "
        f"(NPR {stats['upgraded']:,.1f} Cr), reflecting borrower improvement, while "
        f"downgrades accounted for {stats['downgrade_pct']:.1f}% "
        f"(NPR {stats['downgraded']:,.1f} Cr), representing credit deterioration. "
        f"\n\n"
        f"Risk Assessment: {risk_signal} "
        f"The portfolio concentration ratio is {stats['concentration_ratio']:.1f}%, "
        f"with the dominant category holding the highest closing balance. "
        f"Management should review grade-specific narratives on the following pages "
        f"for targeted remediation and provisioning actions."
    )

    y_pos = 0.625
    for para in exec_text.split("\n\n"):
        wrapped = _wrap_text(para.strip(), 115)
        for line in wrapped:
            ax.text(0, y_pos, line, fontsize=7.5, color=_hex_to_rgb(PDF_DARK),
                    va="top")
            y_pos -= 0.031
        y_pos -= 0.018  # paragraph spacing

    # Grade summary table
    y_tbl = y_pos - 0.015
    ax.text(0, y_tbl, "Grade-Level Summary", fontsize=9.5, fontweight="bold",
            color=_hex_to_rgb(PDF_PRIMARY))
    ax.axhline(y_tbl - 0.018, color=_hex_to_rgb(PDF_BORDER), lw=0.7)

    col_labels = ["Grade", "Opening (Cr)", "Closing (Cr)", "Retained (Cr)",
                  "Upgraded Out", "Downgraded Out", "Net Change"]
    col_xs     = [0, 0.14, 0.27, 0.40, 0.53, 0.66, 0.82]
    y_hdr = y_tbl - 0.045
    ax.add_patch(mpatches.Rectangle((0, y_hdr - 0.004), 1, 0.034,
                                    facecolor=_hex_to_rgb(PDF_LIGHT_BLUE),
                                    edgecolor="none", zorder=2))
    for lbl, cx in zip(col_labels, col_xs):
        ax.text(cx, y_hdr + 0.012, lbl, fontsize=7, fontweight="bold",
                color=_hex_to_rgb(PDF_PRIMARY), va="center")

    y_row = y_hdr - 0.010
    for i, d in enumerate(stats["grade_details"]):
        net = d["closing"] - d["opening"]
        bg_clr = "#F5F9FF" if i % 2 == 0 else PDF_WHITE
        ax.add_patch(mpatches.Rectangle((0, y_row - 0.006), 1, 0.030,
                                        facecolor=_hex_to_rgb(bg_clr),
                                        edgecolor="none", zorder=2))
        row_vals = [
            d["grade"], f"{d['opening']:,.1f}", f"{d['closing']:,.1f}",
            f"{d['retained']:,.1f}", f"{d['upgraded_out']:,.1f}",
            f"{d['downgraded_out']:,.1f}",
            f"{'▲' if net >= 0 else '▼'} {abs(net):,.1f}"
        ]
        row_colors = [PDF_DARK] * 6 + [PDF_SUCCESS if net >= 0 else PDF_ERROR]
        for val, cx, fc in zip(row_vals, col_xs, row_colors):
            ax.text(cx, y_row + 0.009, val, fontsize=7,
                    color=_hex_to_rgb(fc), va="center")
        y_row -= 0.030

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Page 4 — Grade Narratives ─────────────────────────────────────────────────
def _page_narratives(pdf: PdfPages, stats: dict, trans: np.ndarray,
                     prev: np.ndarray, period: str,
                     generated_at: str, total_pages: int):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)
    ax_frame = fig.add_axes([0, 0, 1, 1])
    _draw_page_frame(ax_frame, "Grade-by-Grade Transition Narrative",
                     4, total_pages, period, generated_at)

    ax = fig.add_axes([0.04, 0.07, 0.92, 0.845])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    grade_colors = [PDF_SUCCESS, PDF_WARNING, PDF_WARNING, PDF_ERROR, PDF_ERROR]
    bg_colors    = ["#E8F5E9", "#FFF8E1", "#FFF3E0", "#FFEBEE", "#FCE4EC"]
    y_pos = 0.975

    for i in range(N):
        d = stats["grade_details"][i]

        # Grade header band
        ax.add_patch(mpatches.FancyBboxPatch(
            (0, y_pos - 0.035), 1, 0.038,
            boxstyle="round,pad=0.004",
            facecolor=_hex_to_rgb(bg_colors[i]),
            edgecolor=_hex_to_rgb(grade_colors[i]),
            linewidth=1.0, zorder=3
        ))
        ax.add_patch(mpatches.Rectangle(
            (0, y_pos - 0.035), 0.006, 0.038,
            facecolor=_hex_to_rgb(grade_colors[i]),
            edgecolor="none", zorder=4
        ))
        ax.text(0.012, y_pos - 0.014,
                f"{GRADES[i].upper()}  |  Opening: {d['opening']:,.1f} Cr  "
                f"→  Closing: {d['closing']:,.1f} Cr",
                fontsize=8.5, fontweight="bold",
                color=_hex_to_rgb(grade_colors[i]), va="center", zorder=5)

        y_pos -= 0.044

        # Definition line
        ax.text(0.008, y_pos,
                f"Definition: {GRADE_DESCRIPTIONS[GRADES[i]]}",
                fontsize=6.8, color=_hex_to_rgb(PDF_NEUTRAL),
                style="italic", va="top")
        y_pos -= 0.030

        # Narrative
        narrative = _narrative_for_grade(i, d, stats, trans)
        wrapped   = _wrap_text(narrative, 120)
        for line in wrapped:
            ax.text(0.008, y_pos, line, fontsize=7.2,
                    color=_hex_to_rgb(PDF_DARK), va="top")
            y_pos -= 0.026
        y_pos -= 0.018   # spacing between grades

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Page 5 — Opening vs Closing Bar Chart ────────────────────────────────────
def _page_bar_chart(pdf: PdfPages, stats: dict, period: str,
                    generated_at: str, total_pages: int):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)
    ax_frame = fig.add_axes([0, 0, 1, 1])
    _draw_page_frame(ax_frame, "Opening vs Closing Distribution by Grade",
                     5, total_pages, period, generated_at)

    ax = fig.add_axes([0.10, 0.13, 0.85, 0.75])
    openings = [d["opening"]  for d in stats["grade_details"]]
    closings = [d["closing"]  for d in stats["grade_details"]]
    x = np.arange(N); bw = 0.35

    bars1 = ax.bar(x - bw/2, openings, bw, label="Opening",
                   color=_hex_to_rgb(PDF_ACCENT), alpha=0.85,
                   edgecolor="white", linewidth=0.8)
    bars2 = ax.bar(x + bw/2, closings, bw, label="Closing",
                   color=_hex_to_rgb(PDF_SUCCESS), alpha=0.85,
                   edgecolor="white", linewidth=0.8)

    for bar in bars1:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + max(openings+closings)*0.01,
                    f"{h:,.0f}", ha="center", va="bottom", fontsize=7.5,
                    color=_hex_to_rgb(PDF_ACCENT), fontweight="600")
    for bar in bars2:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + max(openings+closings)*0.01,
                    f"{h:,.0f}", ha="center", va="bottom", fontsize=7.5,
                    color=_hex_to_rgb(PDF_SUCCESS), fontweight="600")

    ax.set_xticks(x)
    ax.set_xticklabels(GRADES, fontsize=9, fontweight="600")
    ax.set_ylabel("Amount (NPR Crore)", fontsize=9, color=_hex_to_rgb(PDF_NEUTRAL))
    ax.set_title(f"Portfolio Distribution — {period}", fontsize=11,
                 fontweight="bold", color=_hex_to_rgb(PDF_DARK), pad=12)
    ax.legend(fontsize=9, framealpha=0.9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_facecolor(PDF_BG)
    ax.grid(axis="y", linestyle="--", alpha=0.5, color=_hex_to_rgb(PDF_BORDER))

    # Annotation box
    net_txt = (f"Net portfolio change: "
               f"NPR {stats['total_closing'] - stats['total_opening']:+,.1f} Cr  |  "
               f"Retention: {stats['retention_pct']:.1f}%  |  "
               f"Downgrade rate: {stats['downgrade_pct']:.1f}%")
    ax.text(0.5, -0.14, net_txt, ha="center", va="center",
            fontsize=8, color=_hex_to_rgb(PDF_NEUTRAL),
            transform=ax.transAxes,
            bbox=dict(boxstyle="round,pad=0.4", facecolor=_hex_to_rgb(PDF_LIGHT_BLUE),
                      edgecolor=_hex_to_rgb(PDF_BORDER), linewidth=0.8))

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Page 6 — Retention & Pie Chart ───────────────────────────────────────────
def _page_retention_pie(pdf: PdfPages, stats: dict, period: str,
                        generated_at: str, total_pages: int):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)
    ax_frame = fig.add_axes([0, 0, 1, 1])
    _draw_page_frame(ax_frame, "Retention & Migration Analysis",
                     6, total_pages, period, generated_at)

    gs  = gridspec.GridSpec(1, 2, figure=fig, left=0.06, right=0.96,
                            top=0.88, bottom=0.12, wspace=0.35)
    ax1 = fig.add_subplot(gs[0])   # Retention bar
    ax2 = fig.add_subplot(gs[1])   # Pie

    # Retention rate per grade
    ret_rates = []
    for d in stats["grade_details"]:
        ret_rates.append(d["retained"] / d["opening"] * 100 if d["opening"] > 0 else 0)

    bar_colors = [_hex_to_rgb(PDF_SUCCESS) if r >= 70
                  else _hex_to_rgb(PDF_WARNING) if r >= 40
                  else _hex_to_rgb(PDF_ERROR) for r in ret_rates]
    bars = ax1.barh(GRADES, ret_rates, color=bar_colors, edgecolor="white",
                    linewidth=0.8, height=0.55)
    for bar, rate in zip(bars, ret_rates):
        ax1.text(min(rate + 1.5, 102), bar.get_y() + bar.get_height()/2,
                 f"{rate:.1f}%", va="center", fontsize=8.5,
                 color=_hex_to_rgb(PDF_DARK), fontweight="600")

    ax1.set_xlim(0, 115); ax1.set_xlabel("Retention Rate (%)", fontsize=8.5)
    ax1.set_title("Retention Rate by Grade", fontsize=10, fontweight="bold",
                  color=_hex_to_rgb(PDF_DARK))
    ax1.axvline(70, color=_hex_to_rgb(PDF_WARNING), lw=1.2,
                linestyle="--", label="70% threshold")
    ax1.legend(fontsize=7.5); ax1.spines[["top", "right"]].set_visible(False)
    ax1.set_facecolor(PDF_BG)
    ax1.grid(axis="x", linestyle="--", alpha=0.4)

    # Pie — portfolio composition (closing)
    closings = [d["closing"] for d in stats["grade_details"]]
    non_zero = [(c, g) for c, g in zip(closings, GRADES) if c > 0]
    if non_zero:
        vals, lbls = zip(*non_zero)
        pie_colors = [_hex_to_rgb(c) for c in
                      [PDF_SUCCESS, "#66BB6A", PDF_WARNING, PDF_ERROR, "#B71C1C"]][:len(vals)]
        wedges, texts, autotexts = ax2.pie(
            vals, labels=lbls, autopct="%1.1f%%",
            colors=pie_colors[:len(vals)],
            startangle=90, pctdistance=0.78,
            wedgeprops=dict(edgecolor="white", linewidth=1.5)
        )
        for t in texts:     t.set_fontsize(8.5)
        for t in autotexts: t.set_fontsize(7.5); t.set_fontweight("bold")
        ax2.set_title("Closing Portfolio Composition", fontsize=10,
                      fontweight="bold", color=_hex_to_rgb(PDF_DARK))

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Page 7 — Flow Analysis ────────────────────────────────────────────────────
def _page_flow_analysis(pdf: PdfPages, stats: dict, trans: np.ndarray,
                        prev: np.ndarray, period: str,
                        generated_at: str, total_pages: int):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)
    ax_frame = fig.add_axes([0, 0, 1, 1])
    _draw_page_frame(ax_frame, "Flow Analysis — Upgrades & Downgrades by Grade",
                     7, total_pages, period, generated_at)

    gs   = gridspec.GridSpec(2, 1, figure=fig, left=0.10, right=0.94,
                             top=0.87, bottom=0.09, hspace=0.45)
    ax1  = fig.add_subplot(gs[0])
    ax2  = fig.add_subplot(gs[1])

    ups  = [d["upgraded_out"]   for d in stats["grade_details"]]
    dns  = [d["downgraded_out"] for d in stats["grade_details"]]
    x    = np.arange(N); bw = 0.32

    # Absolute flows
    ax1.bar(x - bw/2, ups, bw, label="Upgraded Out",
            color=_hex_to_rgb(PDF_SUCCESS), alpha=0.85,
            edgecolor="white", linewidth=0.8)
    ax1.bar(x + bw/2, dns, bw, label="Downgraded Out",
            color=_hex_to_rgb(PDF_ERROR), alpha=0.85,
            edgecolor="white", linewidth=0.8)
    ax1.set_xticks(x); ax1.set_xticklabels(GRADES, fontsize=8.5)
    ax1.set_ylabel("NPR Crore", fontsize=8); ax1.legend(fontsize=8)
    ax1.set_title("Absolute Upgrade & Downgrade Flows", fontsize=9.5,
                  fontweight="bold", color=_hex_to_rgb(PDF_DARK))
    ax1.spines[["top", "right"]].set_visible(False)
    ax1.set_facecolor(PDF_BG)
    ax1.grid(axis="y", linestyle="--", alpha=0.4)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))

    # Percentage flows
    up_pcts = [u / o * 100 if o > 0 else 0 for u, o in zip(ups, prev)]
    dn_pcts = [d / o * 100 if o > 0 else 0 for d, o in zip(dns, prev)]
    ax2.bar(x - bw/2, up_pcts, bw, label="Upgrade %",
            color=_hex_to_rgb(PDF_SUCCESS), alpha=0.85,
            edgecolor="white", linewidth=0.8)
    ax2.bar(x + bw/2, dn_pcts, bw, label="Downgrade %",
            color=_hex_to_rgb(PDF_ERROR), alpha=0.85,
            edgecolor="white", linewidth=0.8)
    ax2.axhline(10, color=_hex_to_rgb(PDF_WARNING), lw=1.2,
                linestyle="--", label="10% Alert Threshold")
    ax2.axhline(20, color=_hex_to_rgb(PDF_ERROR), lw=1.0,
                linestyle=":", label="20% Critical Threshold")
    ax2.set_xticks(x); ax2.set_xticklabels(GRADES, fontsize=8.5)
    ax2.set_ylabel("% of Opening Balance", fontsize=8); ax2.legend(fontsize=7.5)
    ax2.set_title("Migration Rates (% of Opening Balance)", fontsize=9.5,
                  fontweight="bold", color=_hex_to_rgb(PDF_DARK))
    ax2.spines[["top", "right"]].set_visible(False)
    ax2.set_facecolor(PDF_BG)
    ax2.grid(axis="y", linestyle="--", alpha=0.4)
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.1f}%"))

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Page 8 — Percentage Matrix ───────────────────────────────────────────────
def _page_pct_matrix(pdf: PdfPages, trans: np.ndarray, prev: np.ndarray,
                     period: str, generated_at: str, total_pages: int):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)
    ax_frame = fig.add_axes([0, 0, 1, 1])
    _draw_page_frame(ax_frame, "Percentage Transition Matrix (% of Opening Balance)",
                     8, total_pages, period, generated_at)

    ax = fig.add_axes([0.07, 0.10, 0.88, 0.78])
    pct_matrix = np.zeros((N, N))
    for r in range(N):
        for c in range(N):
            pct_matrix[r, c] = trans[r, c] / prev[r] * 100 if prev[r] > 0 else 0

    # Custom colormap: green diagonal, blue upgrades, red downgrades
    display = np.copy(pct_matrix)
    im = ax.imshow(display, cmap="RdYlGn_r", vmin=0, vmax=100, aspect="auto")

    # Diagonal overlay
    for i in range(N):
        ax.add_patch(plt.Rectangle((i - 0.5, i - 0.5), 1, 1,
                                   fill=True, facecolor=_hex_to_rgb(DIAG_BG),
                                   edgecolor="white", lw=2, zorder=3))
        ax.text(i, i, f"{pct_matrix[i,i]:.1f}%",
                ha="center", va="center", fontsize=11,
                fontweight="bold", color=_hex_to_rgb(DIAG_FG), zorder=4)

    for r in range(N):
        for c in range(N):
            if r != c:
                v = pct_matrix[r, c]
                color = "white" if v > 50 else _hex_to_rgb(PDF_DARK)
                ax.text(c, r, f"{v:.1f}%",
                        ha="center", va="center", fontsize=9.5,
                        color=color, zorder=4)

    ax.set_xticks(range(N)); ax.set_yticks(range(N))
    ax.set_xticklabels([f"To: {g}" for g in GRADES], fontsize=9, fontweight="600")
    ax.set_yticklabels([f"From: {g}" for g in GRADES], fontsize=9, fontweight="600")
    ax.set_title(f"Transition Probability Matrix — {period}\n"
                 "(Each row sums to 100% of opening balance; diagonal = retained)",
                 fontsize=10, fontweight="bold", color=_hex_to_rgb(PDF_DARK), pad=10)

    cbar = plt.colorbar(im, ax=ax, orientation="vertical",
                        fraction=0.03, pad=0.02)
    cbar.set_label("% of Opening Balance", fontsize=8)
    cbar.ax.tick_params(labelsize=7)

    ax.spines[:].set_visible(False)

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Page 9 — Conclusions ─────────────────────────────────────────────────────
def _page_conclusions(pdf: PdfPages, stats: dict, trans: np.ndarray,
                      prev: np.ndarray, period: str,
                      generated_at: str, total_pages: int):
    fig = plt.figure(figsize=(11, 8.5)); fig.patch.set_facecolor(PDF_WHITE)
    ax_frame = fig.add_axes([0, 0, 1, 1])
    _draw_page_frame(ax_frame, "Risk Concentration & Conclusions",
                     9, total_pages, period, generated_at)

    ax = fig.add_axes([0.05, 0.08, 0.90, 0.845])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # --- Concentration chart (inline horizontal bars) ---
    ax.text(0, 0.97, "Grade Concentration — Closing Portfolio",
            fontsize=10, fontweight="bold", color=_hex_to_rgb(PDF_PRIMARY))
    ax.axhline(0.955, color=_hex_to_rgb(PDF_BORDER), lw=0.7)

    total_cl = stats["total_closing"]
    bar_ys   = [0.895, 0.845, 0.795, 0.745, 0.695]
    bar_cols = [PDF_SUCCESS, "#A5D6A7", PDF_WARNING, PDF_ERROR, "#B71C1C"]
    for i, (d, y, bc) in enumerate(zip(stats["grade_details"], bar_ys, bar_cols)):
        pct  = d["closing"] / total_cl * 100 if total_cl > 0 else 0
        blen = pct / 100 * 0.72
        ax.text(0, y + 0.008, GRADES[i], fontsize=8, fontweight="600",
                color=_hex_to_rgb(PDF_DARK))
        ax.add_patch(mpatches.FancyBboxPatch(
            (0.16, y - 0.002), blen, 0.025,
            boxstyle="round,pad=0.002",
            facecolor=_hex_to_rgb(bc), edgecolor="none", zorder=3
        ))
        ax.add_patch(mpatches.Rectangle(
            (0.16, y - 0.002), 0.72, 0.025,
            facecolor="none",
            edgecolor=_hex_to_rgb(PDF_BORDER), linewidth=0.7, zorder=2
        ))
        ax.text(0.16 + blen + 0.01, y + 0.008,
                f"{pct:.1f}%  ({d['closing']:,.1f} Cr)",
                fontsize=8, color=_hex_to_rgb(PDF_DARK))

    # --- Key Findings --------------------------------------------------------
    ax.text(0, 0.635, "Key Findings & Risk Indicators",
            fontsize=10, fontweight="bold", color=_hex_to_rgb(PDF_PRIMARY))
    ax.axhline(0.620, color=_hex_to_rgb(PDF_BORDER), lw=0.7)

    dngrade_pct = stats["downgrade_pct"]
    findings = [
        (f"Portfolio {'EXPANDED' if stats['total_closing'] >= stats['total_opening'] else 'CONTRACTED'} "
         f"by NPR {abs(stats['total_closing'] - stats['total_opening']):,.1f} Cr "
         f"({abs((stats['total_closing']/stats['total_opening'] - 1)*100) if stats['total_opening'] else 0:.1f}%) "
         f"during the period.",
         PDF_PRIMARY),
        (f"Retention Rate: {stats['retention_pct']:.1f}% — "
         f"{'Strong stability; portfolio quality maintained.' if stats['retention_pct'] >= 70 else 'Moderate; review migration drivers.'}",
         PDF_SUCCESS if stats['retention_pct'] >= 70 else PDF_WARNING),
        (f"Downgrade Rate: {dngrade_pct:.1f}% — "
         f"{'Within acceptable limits (<10%).' if dngrade_pct < 10 else 'ELEVATED (10-20%): Enhanced monitoring required.' if dngrade_pct < 20 else 'CRITICAL (>20%): Immediate management intervention needed.'}",
         PDF_SUCCESS if dngrade_pct < 10 else PDF_WARNING if dngrade_pct < 20 else PDF_ERROR),
        (f"Migration Rate: {stats['migration_rate']:.1f}% — "
         f"{'Low portfolio mobility.' if stats['migration_rate'] < 15 else 'Moderate migration; track trending.' if stats['migration_rate'] < 30 else 'High migration; investigate systemic causes.'}",
         PDF_SUCCESS if stats['migration_rate'] < 15 else PDF_WARNING),
        (f"Concentration Risk: {stats['concentration_ratio']:.1f}% in top grade — "
         f"{'Well diversified.' if stats['concentration_ratio'] < 60 else 'Moderate concentration.' if stats['concentration_ratio'] < 80 else 'HIGH concentration risk.'}",
         PDF_SUCCESS if stats['concentration_ratio'] < 60 else PDF_WARNING),
    ]
    y_f = 0.595
    for text, color in findings:
        ax.add_patch(mpatches.FancyBboxPatch(
            (0, y_f - 0.005), 1, 0.032,
            boxstyle="round,pad=0.004",
            facecolor=tuple(min(1, v + 0.88 * (1 - v))
                            for v in _hex_to_rgb(color)),
            edgecolor=_hex_to_rgb(color), linewidth=0.8, zorder=2
        ))
        ax.add_patch(mpatches.Rectangle(
            (0, y_f - 0.005), 0.005, 0.032,
            facecolor=_hex_to_rgb(color), edgecolor="none", zorder=3
        ))
        wrapped = _wrap_text(text, 115)
        ax.text(0.012, y_f + 0.011, wrapped[0],
                fontsize=7.5, color=_hex_to_rgb(PDF_DARK),
                va="center", zorder=4)
        y_f -= 0.042

    # --- Recommendations -----------------------------------------------------
    ax.text(0, y_f - 0.005, "Management Recommendations",
            fontsize=10, fontweight="bold", color=_hex_to_rgb(PDF_PRIMARY))
    ax.axhline(y_f - 0.020, color=_hex_to_rgb(PDF_BORDER), lw=0.7)
    y_r = y_f - 0.045

    recs = [
        "Review all Substandard, Doubtful and Bad loans for updated collateral valuations and provisioning.",
        "Strengthen early-warning systems for Watchlist borrowers to prevent further downgrade.",
        "Validate recovery actions and legal proceedings for Bad-debt accounts.",
        "Assess sectoral and geographic concentrations within each grade.",
        "Benchmark current migration rates against prior periods and industry peers.",
        "Ensure NRB provisioning requirements are fully met for all classified categories.",
    ]
    for j, rec in enumerate(recs):
        ax.text(0.01, y_r,
                f"{j+1}.  {rec}",
                fontsize=7.5, color=_hex_to_rgb(PDF_DARK), va="top")
        y_r -= 0.035

    # Disclaimer
    ax.add_patch(mpatches.FancyBboxPatch(
        (0, 0.005), 1, 0.042,
        boxstyle="round,pad=0.005",
        facecolor=_hex_to_rgb(PDF_LIGHT_BLUE),
        edgecolor=_hex_to_rgb(PDF_BORDER), linewidth=0.8, zorder=2
    ))
    ax.text(0.5, 0.026,
            "Disclaimer: This report is generated from uploaded portfolio data and is intended "
            "for internal management purposes only. "
            "Figures are in NPR Crore. All classifications follow NRB directives.",
            ha="center", va="center", fontsize=6.8,
            color=_hex_to_rgb(PDF_NEUTRAL), style="italic", zorder=3)

    pdf.savefig(fig, bbox_inches="tight"); plt.close(fig)

# ── Master PDF Builder ────────────────────────────────────────────────────────
def build_pdf_report(trans: np.ndarray, prev: np.ndarray,
                     stats: dict, period: str) -> bytes:
    """Assemble all 9 pages into a single PDF and return bytes."""
    buf          = io.BytesIO()
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    TOTAL_PAGES  = 9

    with PdfPages(buf) as pdf:
        # PDF metadata
        d = pdf.infodict()
        d["Title"]   = f"Loan Transition Matrix Report — {period}"
        d["Author"]  = "Loan Matrix Dashboard"
        d["Subject"] = "Portfolio Migration Analysis"
        d["Keywords"]= "NRB, Loan, Transition Matrix, Credit Risk"
        d["CreationDate"] = datetime.now()

        _page_cover(pdf, stats, period, generated_at)
        _page_heatmap(pdf, trans, prev, period, generated_at, TOTAL_PAGES)
        _page_kpi_summary(pdf, stats, trans, prev, period, generated_at, TOTAL_PAGES)
        _page_narratives(pdf, stats, trans, prev, period, generated_at, TOTAL_PAGES)
        _page_bar_chart(pdf, stats, period, generated_at, TOTAL_PAGES)
        _page_retention_pie(pdf, stats, period, generated_at, TOTAL_PAGES)
        _page_flow_analysis(pdf, stats, trans, prev, period, generated_at, TOTAL_PAGES)
        _page_pct_matrix(pdf, trans, prev, period, generated_at, TOTAL_PAGES)
        _page_conclusions(pdf, stats, trans, prev, period, generated_at, TOTAL_PAGES)

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
        '<div style="font-size:12px;color:var(--neutral);">v3.0 | + PDF Report</div>'
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

        st.markdown("### 📄 Management PDF Report")
        st.markdown(
            "Click **Generate PDF** to build a full 9-page management report containing:\n"
            "- Cover page with portfolio summary\n"
            "- Transition matrix heatmap\n"
            "- KPI dashboard & executive summary\n"
            "- **Grade-by-grade transition narratives** with risk alerts\n"
            "- Opening vs Closing bar chart\n"
            "- Retention & portfolio composition pie chart\n"
            "- Upgrade / Downgrade flow analysis\n"
            "- Percentage transition matrix (heatmap)\n"
            "- Risk concentration bars & management recommendations"
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

        # Preview of what's in each page
        st.divider()
        st.markdown("#### 📋 Report Contents Preview")
        pages_info = [
            ("1", "Cover Page",
             "Portfolio KPIs, report overview, period details, contents table"),
            ("2", "Transition Matrix Heatmap",
             "Color-coded matrix (retained/upgrade/downgrade) with legend"),
            ("3", "KPI Dashboard & Executive Summary",
             "8 KPI cards, grade-level table, written executive narrative"),
            ("4", "Grade Narratives",
             "Opening→Closing, retention, upgrade/downgrade, inflow, risk alerts per grade"),
            ("5", "Opening vs Closing Chart",
             "Grouped bar chart comparing portfolio size across grades"),
            ("6", "Retention & Composition",
             "Horizontal retention rate bars + closing portfolio pie chart"),
            ("7", "Flow Analysis",
             "Absolute and % upgrade/downgrade bars with alert thresholds"),
            ("8", "Percentage Matrix",
             "Heatmap of transition probabilities (% of opening balance)"),
            ("9", "Conclusions",
             "Concentration bars, 5 key findings, 6 management recommendations, disclaimer"),
        ]
        for pg, title, desc in pages_info:
            with st.expander(f"Page {pg} — {title}"):
                st.markdown(f"**Content:** {desc}")

# ── TAB 5: GUIDE ─────────────────────────────────────────────────────────────
with tab_guide:
    st.markdown("## ℹ️ User Guide")
    st.markdown(
        "1. **Prepare Template**: Ensure your Excel file has: "
        "Row 1 = Headers (`Good`, `Watchlist`, etc.), Column A = Row Labels, "
        "5x5 Data Grid = Numbers only.\n"
        "2. **Upload**: Drag & drop your `.xlsx` file in the **Upload** tab.\n"
        "3. **Generate**: Click *Generate Dashboard* to compute all metrics.\n"
        "4. **PDF Report**: Go to the **PDF Report** tab and click *Generate PDF Report*.\n"
        "5. **Download**: Use the download buttons for PNG, SVG, or PDF exports."
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
