import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
import io
import json

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="NRB Loan Transition Matrix",
    page_icon="🏦",
    layout="wide"
)

# ── Constants ─────────────────────────────────────────────────────────────────
GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
N = 5

# ── Session State ─────────────────────────────────────────────────────────────
defaults = {
    "prev": None,
    "matrix": None,
    "period": "",
    "generated": False
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── Excel Loader (FIXED) ──────────────────────────────────────────────────────
def load_excel(file_bytes):
    try:
        return pd.read_excel(
            io.BytesIO(file_bytes),
            header=None,
            engine="openpyxl"  # critical fix
        )
    except ImportError:
        raise ValueError(
            "Missing dependency: install openpyxl → pip install openpyxl"
        )

# ── Parser ────────────────────────────────────────────────────────────────────
def parse_template(file_bytes):
    df = load_excel(file_bytes)

    def norm(x):
        return str(x).strip().lower()

    grade_norms = [norm(g) for g in GRADES]

    # Detect header
    header_idx = None
    for i, row in df.iterrows():
        matches = sum(1 for c in row if norm(c) in grade_norms)
        if matches >= 4:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Header row not found")

    header = df.iloc[header_idx]

    col_map = {}
    for ci, val in enumerate(header):
        if norm(val) in grade_norms:
            col_map[norm(val)] = ci

    if set(col_map.keys()) != set(grade_norms):
        raise ValueError("Missing grade columns")

    # Detect rows
    rows = {}
    for i in range(header_idx + 1, len(df)):
        row = df.iloc[i]
        first = next((c for c in row if pd.notna(c)), None)
        if norm(first) in grade_norms:
            rows[norm(first)] = row

    if len(rows) != 5:
        raise ValueError("Expected 5 grade rows")

    trans = np.zeros((5, 5))

    for r, g_from in enumerate(grade_norms):
        for c, g_to in enumerate(grade_norms):
            val = pd.to_numeric(
                rows[g_from].iloc[col_map[g_to]],
                errors="coerce"
            )
            trans[r, c] = 0 if pd.isna(val) else val

    prev = trans.sum(axis=1)
    return prev, trans

# ── Stats ─────────────────────────────────────────────────────────────────────
def compute_stats(trans, prev):
    total = trans.sum()
    retained = np.trace(trans)
    upgraded = np.tril(trans, -1).sum()
    downgraded = np.triu(trans, 1).sum()

    return {
        "opening": prev.sum(),
        "closing": total,
        "retained": retained,
        "upgraded": upgraded,
        "downgraded": downgraded,
        "retention_pct": retained / total * 100 if total else 0
    }

# ── Plot ──────────────────────────────────────────────────────────────────────
def build_figure(trans, prev):
    fig, ax = plt.subplots()

    ax.imshow(trans)

    for i in range(N):
        for j in range(N):
            base = prev[i]
            pct = (trans[i, j] / base * 100) if base > 0 else 0
            ax.text(j, i, f"{trans[i,j]:.1f}\n({pct:.1f}%)",
                    ha="center", va="center", fontsize=8)

    ax.set_xticks(range(N))
    ax.set_yticks(range(N))
    ax.set_xticklabels(GRADES)
    ax.set_yticklabels(GRADES)

    return fig

# ── UI ─────────────────────────────────────────────────────────────────────────

st.title("NRB Loan Transition Matrix")

# Sidebar
with st.sidebar:
    st.session_state.period = st.text_input("Period Label")
    dpi = st.select_slider("Export DPI", [100,150,220,300], value=220)

# Upload
uploaded = st.file_uploader("Upload Excel (.xlsx)", key="file")

if uploaded:
    try:
        prev, trans = parse_template(uploaded.read())
        st.session_state.prev = prev
        st.session_state.matrix = trans
        st.success("Parsed successfully")
    except Exception as e:
        st.error(str(e))

# Display
if st.session_state.prev is not None:

    prev = st.session_state.prev
    trans = st.session_state.matrix

    stats = compute_stats(trans, prev)
    st.write(stats)

    # Balance check FIXED
    diff = abs(prev.sum() - trans.sum())
    balanced = np.isclose(prev.sum(), trans.sum(), atol=1e-2)

    st.write("Balanced:", balanced, "| Diff:", diff)

    if st.button("Generate"):
        st.session_state.generated = True

# Plot
if st.session_state.generated:

    fig = build_figure(trans, prev)
    st.pyplot(fig)

    buf = io.BytesIO()
    fig.savefig(buf, dpi=dpi, format="png")
    buf.seek(0)

    st.download_button("Download PNG", buf, "matrix.png")

    plt.close(fig)

# Export JSON
if st.session_state.generated:
    st.download_button(
        "Download JSON",
        json.dumps({
            "opening": prev.tolist(),
            "matrix": trans.tolist()
        }, indent=2),
        "data.json"
    )
