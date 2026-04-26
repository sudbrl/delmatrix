# ══════════════════════════════════════════════════════════════════════════════
# Loan Transition Matrix Dashboard — v6.0 (Gemini Explanation Focus)
# ══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import numpy as np
import io
import hashlib
import json
import urllib.request

# ── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Loan Transition Matrix", page_icon="🏦", layout="wide")

GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = len(GRADES)

# ── SESSION STATE ────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "prev": None,
        "matrix": None,
        "generated": False,
        "period": "",
        "ai_explanation": None
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ── UTILITIES ────────────────────────────────────────────────────────────────
def hash_file(b): return hashlib.md5(b).hexdigest()

def clean_numeric(val):
    try:
        return float(str(val).replace(",", "").strip())
    except:
        return 0.0

# ── PARSER ───────────────────────────────────────────────────────────────────
@st.cache_data
def parse_excel(_, file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    g = [x.lower() for x in GRADES]

    header_idx = None
    for i in range(10):
        row = [str(x).lower() for x in df.iloc[i]]
        if sum(x in row for x in g) >= 4:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Header row not found")

    col_map = {}
    for i, val in enumerate(df.iloc[header_idx]):
        v = str(val).lower()
        if v in g:
            col_map[v] = i

    rows = {}
    for i in range(header_idx + 1, len(df)):
        key = str(df.iloc[i, 0]).lower()
        if key in g:
            rows[key] = df.iloc[i]

    trans = np.zeros((N, N))
    for i, r in enumerate(g):
        for j, c in enumerate(g):
            trans[i, j] = clean_numeric(rows[r][col_map[c]])

    return trans.sum(axis=1), trans

# ── GEMINI PROMPT (FOCUSED ON AMOUNT + %) ─────────────────────────────────────
def build_prompt(trans, prev, period):
    pct = np.zeros_like(trans)

    for r in range(N):
        if prev[r] > 0:
            pct[r] = (trans[r] / prev[r]) * 100

    lines = []
    lines.append(f"Loan transition matrix explanation for period: {period}")
    lines.append("")
    lines.append("Rows = FROM grade, Columns = TO grade")
    lines.append("")

    lines.append("AMOUNTS (NPR Crore):")
    for i in range(N):
        row = ", ".join(f"{trans[i,j]:.1f}" for j in range(N))
        lines.append(f"{GRADES[i]} -> [{row}] (Opening {prev[i]:.1f})")

    lines.append("")
    lines.append("PERCENTAGES (% of row):")
    for i in range(N):
        row = ", ".join(f"{pct[i,j]:.1f}%" for j in range(N))
        lines.append(f"{GRADES[i]} -> [{row}]")

    lines.append("")
    lines.append("TASK:")
    lines.append(
        "Explain clearly:\n"
        "- What the amounts mean\n"
        "- What the percentages indicate\n"
        "- Where upgrades are happening\n"
        "- Where downgrades are concentrated\n"
        "- Which movements are significant\n\n"
        "Keep it concise, structured, and analytical (not long report)."
    )

    return "\n".join(lines)

# ── GEMINI CALL ──────────────────────────────────────────────────────────────
def call_gemini(prompt, api_key):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"

    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 2048
        }
    }).encode()

    req = urllib.request.Request(
        url,
        data=payload,
        headers={"Content-Type": "application/json"}
    )

    with urllib.request.urlopen(req, timeout=60) as res:
        data = json.loads(res.read().decode())

    try:
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except:
        return json.dumps(data, indent=2)

# ── UI ───────────────────────────────────────────────────────────────────────
st.title("🏦 Loan Transition Matrix")

tab1, tab2, tab3 = st.tabs(["Upload", "Matrix", "AI Explanation"])

# ── UPLOAD ───────────────────────────────────────────────────────────────────
with tab1:
    file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

    if file:
        bytes_data = file.read()
        prev, trans = parse_excel(hash_file(bytes_data), bytes_data)

        st.session_state.prev = prev
        st.session_state.matrix = trans
        st.session_state.generated = True

        st.success("File loaded successfully")

# ── MATRIX ───────────────────────────────────────────────────────────────────
with tab2:
    if not st.session_state.generated:
        st.info("Upload data first")
    else:
        trans = st.session_state.matrix
        prev = st.session_state.prev

        df = pd.DataFrame(
            np.round(trans, 2),
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES
        )
        df["Opening"] = np.round(prev, 2)

        st.markdown("### Transition Amounts (NPR Crore)")
        st.dataframe(df, use_container_width=True)

        pct_df = pd.DataFrame(
            [
                [(trans[r,c]/prev[r]*100) if prev[r] > 0 else 0 for c in range(N)]
                for r in range(N)
            ],
            index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)],
            columns=GRADES
        )

        st.markdown("### Transition Percentages (%)")
        st.dataframe(pct_df.round(1), use_container_width=True)

# ── AI EXPLANATION ───────────────────────────────────────────────────────────
with tab3:
    if not st.session_state.generated:
        st.info("Upload data first")
    else:
        api_key = st.secrets.get("GEMINI_API_KEY", None)

        if not api_key:
            st.warning("Add GEMINI_API_KEY in Streamlit secrets")
        else:
            if st.button("Generate Explanation") or st.session_state.ai_explanation:
                if not st.session_state.ai_explanation:
                    prompt = build_prompt(
                        st.session_state.matrix,
                        st.session_state.prev,
                        st.session_state.period or "N/A"
                    )
                    with st.spinner("Analyzing transitions..."):
                        st.session_state.ai_explanation = call_gemini(prompt, api_key)

                st.markdown("### Gemini Explanation")
                st.markdown(
                    f'<div style="background:#F5F9FF;padding:20px;border-radius:10px;">'
                    f'{st.session_state.ai_explanation.replace(chr(10), "<br>")}'
                    f'</div>',
                    unsafe_allow_html=True
                )
