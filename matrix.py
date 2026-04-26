# ══════════════════════════════════════════════════════════════════════════════
# Loan Transition Matrix — v6.1 (Flow-Level Gemini Interpretation)
# ══════════════════════════════════════════════════════════════════════════════

import streamlit as st
import pandas as pd
import numpy as np
import io
import hashlib
import json
import urllib.request

st.set_page_config(page_title="Loan Transition Matrix", page_icon="🏦", layout="wide")

GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
ICONS  = ["🟢", "🟡", "🟠", "🔴", "⛔"]
N = len(GRADES)

# ── STATE ────────────────────────────────────────────────────────────────────
def init_state():
    for k, v in {
        "prev": None,
        "matrix": None,
        "generated": False,
        "ai_explanation": None
    }.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ── UTIL ─────────────────────────────────────────────────────────────────────
def clean(x):
    try: return float(str(x).replace(",", "").strip())
    except: return 0.0

def file_hash(b): return hashlib.md5(b).hexdigest()

# ── PARSER ───────────────────────────────────────────────────────────────────
@st.cache_data
def parse_excel(_, b):
    df = pd.read_excel(io.BytesIO(b), header=None)
    g = [x.lower() for x in GRADES]

    header = None
    for i in range(10):
        row = [str(x).lower() for x in df.iloc[i]]
        if sum(v in row for v in g) >= 4:
            header = i; break

    if header is None:
        raise ValueError("Header not found")

    col = {str(df.iloc[header, i]).lower(): i for i in range(len(df.columns)) if str(df.iloc[header, i]).lower() in g}

    rows = {}
    for i in range(header+1, len(df)):
        k = str(df.iloc[i,0]).lower()
        if k in g:
            rows[k] = df.iloc[i]

    trans = np.zeros((N,N))
    for i, r in enumerate(g):
        for j, c in enumerate(g):
            trans[i,j] = clean(rows[r][col[c]])

    return trans.sum(axis=1), trans

# ── PROMPT (CRITICAL FIX) ─────────────────────────────────────────────────────
def build_prompt(trans, prev):
    flows = []

    for r in range(N):
        for c in range(N):
            val = trans[r,c]
            if val <= 0: continue

            pct = (val / prev[r] * 100) if prev[r] > 0 else 0

            if r == c:
                movement = "retained"
            elif c < r:
                movement = "upgrade"
            else:
                movement = "downgrade"

            flows.append({
                "from": GRADES[r],
                "to": GRADES[c],
                "amount": round(val,2),
                "pct": round(pct,1),
                "type": movement
            })

    # sort by biggest movements
    flows = sorted(flows, key=lambda x: x["amount"], reverse=True)

    lines = []
    lines.append("You are a senior credit risk analyst.")
    lines.append("Interpret the loan transition flows between credit grades.")
    lines.append("")
    lines.append("Each entry shows: From → To, Amount (NPR Cr), % of that row, Movement type.")
    lines.append("")

    for f in flows:
        lines.append(f"{f['from']} → {f['to']} | {f['amount']} Cr | {f['pct']}% | {f['type']}")

    lines.append("")
    lines.append("TASK:")
    lines.append(
        "Provide a structured analysis:\n"
        "1. Explain the largest movements (top 5–10 flows) clearly\n"
        "2. Interpret what each major flow means (e.g., credit improvement or deterioration)\n"
        "3. Highlight where downgrades are concentrated\n"
        "4. Highlight where upgrades are occurring\n"
        "5. Comment on any unusual or high-percentage transitions\n\n"
        "Be precise. Focus on interpretation of flows, not generic summary."
    )

    return "\n".join(lines)

# ── GEMINI ───────────────────────────────────────────────────────────────────
def call_gemini(prompt, key):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={key}"

    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.2, "maxOutputTokens": 2000}
    }).encode()

    req = urllib.request.Request(url, data=payload, headers={"Content-Type":"application/json"})
    res = urllib.request.urlopen(req, timeout=60)
    data = json.loads(res.read().decode())

    try:
        return data["candidates"][0]["content"]["parts"][0]["text"]
    except:
        return json.dumps(data, indent=2)

# ── UI ───────────────────────────────────────────────────────────────────────
st.title("🏦 Loan Transition Matrix")

tab1, tab2, tab3 = st.tabs(["Upload", "Matrix", "AI Flow Analysis"])

# ── UPLOAD ───────────────────────────────────────────────────────────────────
with tab1:
    f = st.file_uploader("Upload Excel", type=["xlsx"])
    if f:
        b = f.read()
        prev, trans = parse_excel(file_hash(b), b)
        st.session_state.prev = prev
        st.session_state.matrix = trans
        st.session_state.generated = True
        st.success("Loaded")

# ── MATRIX ───────────────────────────────────────────────────────────────────
with tab2:
    if not st.session_state.generated:
        st.info("Upload file first")
    else:
        trans = st.session_state.matrix
        prev  = st.session_state.prev

        df = pd.DataFrame(trans, columns=GRADES,
                          index=[f"{ICONS[i]} {GRADES[i]}" for i in range(N)])
        df["Opening"] = prev

        st.markdown("### Transition Amounts (NPR Cr)")
        st.dataframe(df.round(2), use_container_width=True)

        pct = pd.DataFrame([
            [(trans[r,c]/prev[r]*100) if prev[r]>0 else 0 for c in range(N)]
            for r in range(N)
        ], columns=GRADES, index=df.index)

        st.markdown("### Transition Percentages (%)")
        st.dataframe(pct.round(1), use_container_width=True)

# ── AI ───────────────────────────────────────────────────────────────────────
with tab3:
    if not st.session_state.generated:
        st.info("Upload data first")
    else:
        key = st.secrets.get("GEMINI_API_KEY")

        if not key:
            st.warning("Add GEMINI_API_KEY in secrets")
        else:
            if st.button("Generate Flow Analysis") or st.session_state.ai_explanation:
                if not st.session_state.ai_explanation:
                    prompt = build_prompt(
                        st.session_state.matrix,
                        st.session_state.prev
                    )
                    with st.spinner("Analyzing transitions..."):
                        st.session_state.ai_explanation = call_gemini(prompt, key)

                st.markdown("### Detailed Flow Interpretation")
                st.markdown(
                    f'<div style="background:#F5F9FF;padding:20px;border-radius:10px;">'
                    f'{st.session_state.ai_explanation.replace(chr(10), "<br>")}'
                    f'</div>',
                    unsafe_allow_html=True
                )
