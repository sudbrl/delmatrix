# ══════════════════════════════════════════════════════════════════════════════
# Loan Transition Matrix Dashboard  —  v5.1 (Final Clean Version)
# ══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import numpy as np
import io
import hashlib
import json
import urllib.request
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# ── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Loan Transition Matrix", page_icon="🏦", layout="wide")

GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = len(GRADES)

# ── SESSION STATE ────────────────────────────────────────────────────────────
def init_session_state():
    defaults = {
        "prev": None,
        "matrix": None,
        "generated": False,
        "stats_cache": None,
        "period": "",
        "ai_summary": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_session_state()

# ── UTILITIES ────────────────────────────────────────────────────────────────
def compute_file_hash(b): return hashlib.md5(b).hexdigest()
def format_currency(v): return f"{v:,.1f} cr"
def format_pct(v): return f"{v:.1f}%"

# ── PARSER ───────────────────────────────────────────────────────────────────
def clean_numeric(val):
    try:
        return float(str(val).replace(",", "").strip())
    except:
        return 0.0

@st.cache_data
def parse_template_cached(_, file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    grades_lower = [g.lower() for g in GRADES]

    header_idx = None
    for i in range(10):
        row = [str(x).lower() for x in df.iloc[i]]
        if sum(g in row for g in grades_lower) >= 4:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Header not found")

    col_map = {}
    for i, val in enumerate(df.iloc[header_idx]):
        if str(val).lower() in grades_lower:
            col_map[str(val).lower()] = i

    data = {}
    for i in range(header_idx + 1, len(df)):
        key = str(df.iloc[i, 0]).lower()
        if key in grades_lower:
            data[key] = df.iloc[i]

    trans = np.zeros((N, N))
    for i, g1 in enumerate(grades_lower):
        for j, g2 in enumerate(grades_lower):
            trans[i, j] = clean_numeric(data[g1][col_map[g2]])

    return trans.sum(axis=1), trans

# ── STATS ────────────────────────────────────────────────────────────────────
def compute_stats(trans, prev):
    total = trans.sum()
    retained = np.trace(trans)
    upgraded = sum(trans[r,c] for r in range(N) for c in range(r))
    downgraded = sum(trans[r,c] for r in range(N) for c in range(r+1, N))

    return {
        "total_opening": prev.sum(),
        "total_closing": total,
        "retention_pct": retained / total * 100 if total else 0,
        "upgrade_pct": upgraded / total * 100 if total else 0,
        "downgrade_pct": downgraded / total * 100 if total else 0,
        "migration_rate": (upgraded+downgraded)/total*100 if total else 0,
    }

# ── GEMINI PROMPT ────────────────────────────────────────────────────────────
def build_prompt(trans, prev, stats, period):
    return f"""
Loan Portfolio Analysis for {period}

Opening: {stats['total_opening']}
Closing: {stats['total_closing']}
Retention: {stats['retention_pct']:.1f}%
Upgrade: {stats['upgrade_pct']:.1f}%
Downgrade: {stats['downgrade_pct']:.1f}%

Provide:
1. Executive summary
2. Migration insights
3. Risk analysis
4. Recommendations
"""

def call_gemini(prompt, api_key):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}]
    }).encode()

    req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"})
    res = urllib.request.urlopen(req)
    data = json.loads(res.read())

    try:
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except:
        return json.dumps(data, indent=2)

# ── UI ───────────────────────────────────────────────────────────────────────
st.title("🏦 Loan Transition Matrix")

tab1, tab2, tab3 = st.tabs(["Upload", "Dashboard", "AI Summary"])

# ── UPLOAD ───────────────────────────────────────────────────────────────────
with tab1:
    file = st.file_uploader("Upload Excel", type=["xlsx"])

    if file:
        prev, trans = parse_template_cached(compute_file_hash(file.read()), file.getvalue())
        st.session_state.prev = prev
        st.session_state.matrix = trans
        st.success("File loaded")

        if st.button("Generate Dashboard"):
            st.session_state.generated = True
            st.session_state.stats_cache = compute_stats(trans, prev)
            st.rerun()

# ── DASHBOARD ────────────────────────────────────────────────────────────────
with tab2:
    if not st.session_state.generated:
        st.info("Upload data first")
    else:
        stats = st.session_state.stats_cache
        st.metric("Portfolio", format_currency(stats["total_closing"]))
        st.metric("Retention", format_pct(stats["retention_pct"]))
        st.metric("Migration", format_pct(stats["migration_rate"]))

        st.dataframe(pd.DataFrame(st.session_state.matrix, columns=GRADES))

# ── AI SUMMARY ───────────────────────────────────────────────────────────────
with tab3:
    if not st.session_state.generated:
        st.info("Generate dashboard first")
    else:
        api_key = st.secrets.get("GEMINI_API_KEY", None)

        if not api_key:
            st.warning("Add GEMINI_API_KEY in secrets")
        else:
            if st.button("Generate AI Summary") or st.session_state.ai_summary:
                if not st.session_state.ai_summary:
                    prompt = build_prompt(
                        st.session_state.matrix,
                        st.session_state.prev,
                        st.session_state.stats_cache,
                        st.session_state.period or "N/A"
                    )
                    with st.spinner("Generating..."):
                        st.session_state.ai_summary = call_gemini(prompt, api_key)

                st.markdown(st.session_state.ai_summary)

# ── GUIDE ────────────────────────────────────────────────────────────────────
st.sidebar.markdown("## Guide")
st.sidebar.markdown(
    """
1. Upload Excel file  
2. Generate dashboard  
3. Review KPIs  
4. Generate AI summary  
5. Download insights  
"""
)
