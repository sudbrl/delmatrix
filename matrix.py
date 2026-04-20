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

# ── Constants ─────────────────────────────────────────────────────────────────
GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
N = 5

# ── Session State ─────────────────────────────────────────────────────────────
for key, default in {
    "prev": None,
    "matrix": None,
    "period": "",
    "generated": False,
    "upload_error": None,
    "filename": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ── Helper Functions ──────────────────────────────────────────────────────────

def parse_template(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)

    def norm(s):
        return str(s).strip().lower()

    grade_norms = [norm(g) for g in GRADES]

    # Find header row
    header_row_idx = None
    for i, row in df.iterrows():
        matches = sum(1 for cell in row if norm(cell) in grade_norms)
        if matches >= 4:
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("Header row with grade names not found.")

    header = df.iloc[header_row_idx]

    col_map = {}
    for ci, cell in enumerate(header):
        n = norm(cell)
        if n in grade_norms:
            col_map[n] = ci

    if set(col_map.keys()) != set(grade_norms):
        missing = set(grade_norms) - set(col_map.keys())
        raise ValueError(f"Missing grade columns: {missing}")

    # Detect data rows robustly
    data_rows = {}
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]

        first_cell = next((cell for cell in row if pd.notna(cell)), None)
        first_val = norm(first_cell)

        if first_val in grade_norms:
            data_rows[first_val] = row

    if len(data_rows) < 5:
        raise ValueError("Not all 5 grade rows found.")

    trans = np.zeros((5, 5), dtype=float)

    for ri, g_from in enumerate(grade_norms):
        row = data_rows[g_from]
        for ci, g_to in enumerate(grade_norms):
            val = pd.to_numeric(row.iloc[col_map[g_to]], errors="coerce")
            trans[ri, ci] = 0.0 if pd.isna(val) else val

    prev = trans.sum(axis=1)
    return prev, trans


def compute_stats(trans, prev):
    ret = sum(trans[i, i] for i in range(N))
    up = sum(trans[r, c] for r in range(N) for c in range(r))
    dn = sum(trans[r, c] for r in range(N) for c in range(r + 1, N))
    tot = trans.sum()

    return {
        "total_opening": prev.sum(),
        "total_closing": tot,
        "retained": ret,
        "upgraded": up,
        "downgraded": dn,
        "retention_pct": ret / tot * 100 if tot else 0,
        "upgrade_pct": up / tot * 100 if tot else 0,
        "downgrade_pct": dn / tot * 100 if tot else 0,
        "col_closing": trans.sum(axis=0),
    }


def build_figure(grades, trans, prev, period):
    n = len(grades)
    col_closing = trans.sum(axis=0)

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.imshow(trans, cmap="coolwarm")

    for i in range(n):
        for j in range(n):
            row_total = prev[i]
            pct = (trans[i, j] / row_total * 100) if row_total > 0 else 0
            ax.text(j, i, f"{trans[i,j]:.1f}\n({pct:.1f}%)",
                    ha="center", va="center", fontsize=8)

    ax.set_xticks(range(n))
    ax.set_yticks(range(n))
    ax.set_xticklabels(grades)
    ax.set_yticklabels(grades)

    ax.set_title(f"Transition Matrix — {period}")
    plt.tight_layout()
    return fig


def fig_to_bytes(fig, fmt="png", dpi=220):
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, dpi=dpi, bbox_inches="tight")
    buf.seek(0)
    return buf.read()


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.session_state.period = st.text_input("Period Label", value=st.session_state.period)

    export_dpi = st.select_slider("Export DPI", [100,150,220,300], value=220)

    if st.session_state.generated:
        if st.button("Reset"):
            for k in ["prev","matrix","generated","filename"]:
                st.session_state[k] = None if k!="generated" else False
            st.rerun()


# ── Main ──────────────────────────────────────────────────────────────────────

st.title("NRB Loan Transition Matrix")

uploaded = st.file_uploader(
    "Upload Excel Template",
    type=["xlsx","xls"],
    key="file_uploader"
)

if uploaded:
    try:
        prev, trans = parse_template(uploaded.read())
        st.session_state.prev = prev
        st.session_state.matrix = trans
        st.session_state.generated = False

        st.success("File parsed successfully")

    except Exception as e:
        st.error(f"Parse error: {e}")

if st.session_state.prev is not None:

    prev_arr = st.session_state.prev
    trans_arr = st.session_state.matrix

    stats = compute_stats(trans_arr, prev_arr)

    st.subheader("Summary")
    st.write(stats)

    diff = abs(trans_arr.sum() - prev_arr.sum())
    balanced = np.isclose(trans_arr.sum(), prev_arr.sum(), atol=1e-2)

    st.write("Balanced:", balanced, "Difference:", diff)

    if st.button("Generate Matrix"):
        st.session_state.generated = True

if st.session_state.generated:

    fig = build_figure(GRADES, trans_arr, prev_arr, st.session_state.period)

    st.pyplot(fig)
    plt.close(fig)

    st.download_button(
        "Download PNG",
        fig_to_bytes(fig, "png", export_dpi),
        "matrix.png"
    )

    st.download_button(
        "Download JSON",
        json.dumps({
            "opening": prev_arr.tolist(),
            "matrix": trans_arr.tolist()
        }, indent=2),
        "data.json"
    )
