import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import io

st.set_page_config(
    page_title="Loan Quality Transition Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
    .metric-card {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 14px 16px;
        margin-bottom: 4px;
    }
    .metric-label { font-size: 11px; color: #6c757d; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
    .metric-value { font-size: 22px; font-weight: 600; color: #212529; line-height: 1; }
    .metric-sub { font-size: 11px; margin-top: 4px; }
    .up { color: #dc2626; }
    .dn { color: #059669; }
    .nu { color: #6c757d; }
    .badge {
        display: inline-block;
        font-size: 11px;
        padding: 3px 10px;
        border-radius: 6px;
        font-weight: 500;
        margin-right: 6px;
    }
    .badge-red   { background: #fee2e2; color: #991b1b; }
    .badge-amber { background: #fef3c7; color: #92400e; }
    .badge-green { background: #d1fae5; color: #065f46; }
    .badge-blue  { background: #dbeafe; color: #1e40af; }
    .section-header {
        font-size: 11px; font-weight: 600; color: #6c757d;
        text-transform: uppercase; letter-spacing: 0.6px;
        margin-bottom: 10px; margin-top: 4px;
    }
    .ew-card {
        border-left: 3px solid;
        padding: 8px 12px;
        margin-bottom: 8px;
        border-radius: 0 6px 6px 0;
        background: #f8f9fa;
    }
    .ew-title  { font-size: 13px; font-weight: 600; color: #212529; }
    .ew-desc   { font-size: 12px; color: #6c757d; margin-top: 2px; }
    div[data-testid="stMetric"] { background: #f8f9fa; border-radius: 8px; padding: 12px 14px; }
</style>
""", unsafe_allow_html=True)

GRADES = ["Good", "Watchlist", "Substandard", "Doubtful", "Bad"]
GRADE_COLORS = ["#185fa5", "#1d9e75", "#d97706", "#f97316", "#dc2626"]
GRADE_BG = ["#eaf3de", "#d1fae5", "#fef3c7", "#ffedd5", "#fee2e2"]

DEFAULT_DATA = np.array([
    [3394.788, 363.653,  12.612,   0.000,   0.000],
    [ 230.898, 425.444,  84.323,   0.000,   0.000],
    [  24.602,  20.287,  23.810,  52.986,   0.000],
    [   2.840,   2.335,   2.906,  44.073,  10.472],
    [  20.561,   0.376,   0.101,   0.317, 247.836],
])


def parse_upload(file) -> np.ndarray | None:
    try:
        df = pd.read_excel(file, header=None)
        nums = df.select_dtypes(include=[np.number])
        arr = nums.values
        if arr.shape[0] >= 5 and arr.shape[1] >= 5:
            return arr[:5, :5].astype(float)
    except Exception:
        pass
    return None


def compute_metrics(raw: np.ndarray):
    row_totals = raw.sum(axis=1)
    col_totals = raw.sum(axis=0)
    grand_total = raw.sum()

    pct = np.where(
        row_totals[:, None] > 0,
        raw / row_totals[:, None] * 100,
        0.0,
    )

    retention = np.diag(pct)

    roll_fwd = np.array([
        sum(pct[i, j] for j in range(len(GRADES)) if j > i)
        for i in range(len(GRADES))
    ])
    upgrade = np.array([
        sum(raw[i, j] for j in range(len(GRADES)) if j < i)
        for i in range(len(GRADES))
    ])
    downgrade = np.array([
        sum(raw[i, j] for j in range(len(GRADES)) if j > i)
        for i in range(len(GRADES))
    ])

    npl_balance = col_totals[3] + col_totals[4]
    npl_ratio = npl_balance / grand_total * 100

    return dict(
        raw=raw,
        pct=pct,
        row_totals=row_totals,
        col_totals=col_totals,
        grand_total=grand_total,
        retention=retention,
        roll_fwd=roll_fwd,
        upgrade=upgrade,
        downgrade=downgrade,
        npl_balance=npl_balance,
        npl_ratio=npl_ratio,
    )


def color_cell(val: float, row: int, col: int) -> str:
    if row == col:
        intensity = val / 100
        r = int(37 + (59 - 37) * intensity)
        g = int(99 + (130 - 99) * intensity)
        b = int(235 - (235 - 246) * intensity)
        return f"rgba({r},{g},{b},{0.12 + intensity * 0.45})"
    if col > row and val > 2:
        intensity = min(val / 60, 1)
        return f"rgba(220,38,38,{0.08 + intensity * 0.38})"
    if col < row and val > 2:
        return "rgba(5,150,105,0.12)"
    return "rgba(0,0,0,0)"


def fmt(val: float, decimals: int = 1) -> str:
    return f"{val:,.{decimals}f}"


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Data Input")
    uploaded = st.file_uploader(
        "Upload transition matrix (.xlsx)",
        type=["xlsx"],
        help="5×5 matrix: rows = from-grade, columns = to-grade, values in $M",
    )
    st.markdown("---")
    st.markdown("**Grade labels**")
    custom_grades = []
    for g in GRADES:
        custom_grades.append(st.text_input(g, value=g, key=f"grade_{g}"))
    st.markdown("---")
    st.markdown("**Thresholds**")
    npl_threshold = st.slider("NPL alert threshold (%)", 1.0, 15.0, 5.0, 0.5)
    roll_threshold = st.slider("Roll-fwd alert threshold (%)", 5.0, 40.0, 10.0, 1.0)
    retention_floor = st.slider("Retention floor (%)", 40.0, 95.0, 80.0, 1.0)
    st.markdown("---")
    st.markdown("**Export**")
    export_btn = st.button("Download summary CSV")

raw_data = DEFAULT_DATA
source_label = "Default sample data (Template.xlsx)"
if uploaded:
    parsed = parse_upload(uploaded)
    if parsed is not None:
        raw_data = parsed
        source_label = f"Uploaded: {uploaded.name}"
    else:
        st.sidebar.error("Could not parse file — using default data.")

m = compute_metrics(raw_data)
G = custom_grades  # potentially renamed grades

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("## Loan Quality Transition Dashboard")
st.caption(f"Source: {source_label} &nbsp;|&nbsp; Total Book: **${m['grand_total']:,.0f}M**")

badges = []
if m["npl_ratio"] > npl_threshold:
    badges.append(f'<span class="badge badge-red">NPL {m["npl_ratio"]:.1f}% — Breach</span>')
else:
    badges.append(f'<span class="badge badge-green">NPL {m["npl_ratio"]:.1f}% — OK</span>')
if m["roll_fwd"][1] > roll_threshold:
    badges.append(f'<span class="badge badge-amber">Watch Roll {m["roll_fwd"][1]:.1f}% — Elevated</span>')
good_pct = m["col_totals"][0] / m["grand_total"] * 100
badges.append(f'<span class="badge badge-green">Good {good_pct:.1f}%</span>')
st.markdown(" ".join(badges), unsafe_allow_html=True)

st.markdown("---")

# ── KPI row ───────────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5, k6 = st.columns(6)
kpi_cols = [k1, k2, k3, k4, k5, k6]
kpis = [
    ("Total Book", f"${m['grand_total']:,.0f}M", "", "nu"),
    ("NPL Balance", f"${m['npl_balance']:,.1f}M", "Doubtful + Bad", "up" if m["npl_ratio"] > npl_threshold else "dn"),
    ("NPL Ratio", f"{m['npl_ratio']:.2f}%", f"{'Above' if m['npl_ratio'] > npl_threshold else 'Within'} {npl_threshold}% threshold", "up" if m["npl_ratio"] > npl_threshold else "dn"),
    ("Watch Ratio", f"{m['col_totals'][1]/m['grand_total']*100:.2f}%", "Leading indicator", "nu"),
    ("Sub-std Ratio", f"{m['col_totals'][2]/m['grand_total']*100:.2f}%", "Stage 2 proxy", "up"),
    ("Good Retention", f"{m['retention'][0]:.1f}%", "Stay in Good", "dn"),
]
for col, (label, val, sub, cls) in zip(kpi_cols, kpis):
    col.metric(label=label, value=val, delta=sub)

st.markdown("")

# ── Stage composition bar ─────────────────────────────────────────────────────
st.markdown('<div class="section-header">Book composition by grade ($M)</div>', unsafe_allow_html=True)
fig_comp = go.Figure()
for i, g in enumerate(G):
    pct_v = m["col_totals"][i] / m["grand_total"] * 100
    fig_comp.add_trace(go.Bar(
        name=g,
        x=[m["col_totals"][i]],
        y=["Portfolio"],
        orientation="h",
        marker_color=GRADE_COLORS[i],
        hovertemplate=f"<b>{g}</b><br>${m['col_totals'][i]:.1f}M ({pct_v:.1f}%)<extra></extra>",
        text=f"{g}<br>${m['col_totals'][i]:.0f}M",
        textposition="inside",
        insidetextanchor="middle",
        textfont=dict(size=11, color="white"),
    ))
fig_comp.update_layout(
    barmode="stack", height=80, margin=dict(l=0, r=0, t=0, b=0),
    showlegend=False, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
    yaxis=dict(showgrid=False, showticklabels=False),
)
st.plotly_chart(fig_comp, use_container_width=True, config={"displayModeBar": False})

# ── Transition Matrix ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">Transition matrix — row = from-grade, column = to-grade (% of opening balance)</div>', unsafe_allow_html=True)

header_cols = st.columns([1.8] + [1] * 5 + [0.9])
header_cols[0].markdown("**From \\ To**")
for j, g in enumerate(G):
    header_cols[j + 1].markdown(f"**{g}**")
header_cols[6].markdown("**Total**")

for i, g_from in enumerate(G):
    row_cols = st.columns([1.8] + [1] * 5 + [0.9])
    row_cols[0].markdown(f"**{g_from}**")
    for j in range(5):
        v = m["pct"][i, j]
        bg = color_cell(v, i, j)
        is_diag = i == j
        fw = "700" if is_diag else "400"
        if i == j:
            fc = "#1e3a8a"
        elif j > i and v > 2:
            fc = "#991b1b"
        elif j < i and v > 2:
            fc = "#065f46"
        else:
            fc = "#212529"
        row_cols[j + 1].markdown(
            f'<div style="background:{bg};border-radius:4px;padding:5px 6px;text-align:center;font-size:13px;font-weight:{fw};color:{fc}">'
            f'{"—" if v == 0 else f"{v:.1f}%"}</div>',
            unsafe_allow_html=True,
        )
    row_cols[6].markdown(
        f'<div style="color:#6c757d;font-size:12px;padding:5px 0">${m["row_totals"][i]:,.0f}M</div>',
        unsafe_allow_html=True,
    )

st.markdown("")

# ── Migration flows + Retention/Roll rates ────────────────────────────────────
col_mig, col_rates = st.columns(2)

with col_mig:
    st.markdown('<div class="section-header">Migration flows by grade ($M)</div>', unsafe_allow_html=True)
    fig_mig = go.Figure()
    fig_mig.add_trace(go.Bar(
        name="Downgrade", x=G, y=[-v for v in m["downgrade"]],
        marker_color="rgba(220,38,38,0.75)", marker_line_width=0,
        hovertemplate="<b>%{x}</b><br>Downgrade: $%{customdata:.1f}M<extra></extra>",
        customdata=m["downgrade"],
    ))
    fig_mig.add_trace(go.Bar(
        name="Upgrade", x=G, y=m["upgrade"],
        marker_color="rgba(5,150,105,0.75)", marker_line_width=0,
        hovertemplate="<b>%{x}</b><br>Upgrade: $%{y:.1f}M<extra></extra>",
    ))
    fig_mig.update_layout(
        barmode="overlay", height=240, margin=dict(l=10, r=10, t=10, b=30),
        legend=dict(orientation="h", y=1.12, x=0, font_size=11),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(gridcolor="rgba(0,0,0,0.05)"),
        yaxis=dict(gridcolor="rgba(0,0,0,0.05)", title="$M",
                   tickformat="$,.0f", tickprefix=""),
    )
    st.plotly_chart(fig_mig, use_container_width=True, config={"displayModeBar": False})

with col_rates:
    st.markdown('<div class="section-header">Grade retention rate</div>', unsafe_allow_html=True)
    for i, g in enumerate(G):
        r = m["retention"][i]
        col = "#059669" if r >= retention_floor else ("#d97706" if r >= 50 else "#dc2626")
        pct_bar = int(r)
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:5px">'
            f'<span style="width:90px;font-size:12px;color:#6c757d">{g}</span>'
            f'<div style="flex:1;background:#f1f3f5;border-radius:3px;height:13px;overflow:hidden">'
            f'<div style="width:{pct_bar}%;height:100%;background:{col};border-radius:3px"></div></div>'
            f'<span style="width:42px;text-align:right;font-size:12px;font-weight:600;color:{col}">{r:.1f}%</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
    st.markdown('<div class="section-header" style="margin-top:12px">Roll-forward rate (to worse grade)</div>', unsafe_allow_html=True)
    for i, g in enumerate(G):
        r = m["roll_fwd"][i]
        col = "#059669" if r < 5 else ("#d97706" if r < roll_threshold else "#dc2626")
        pct_bar = min(int(r), 100)
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:5px">'
            f'<span style="width:90px;font-size:12px;color:#6c757d">{g}</span>'
            f'<div style="flex:1;background:#f1f3f5;border-radius:3px;height:13px;overflow:hidden">'
            f'<div style="width:{pct_bar}%;height:100%;background:{col};border-radius:3px"></div></div>'
            f'<span style="width:42px;text-align:right;font-size:12px;font-weight:600;color:{col}">{r:.1f}%</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

# ── Heatmap of % matrix ───────────────────────────────────────────────────────
st.markdown('<div class="section-header">Transition probability heatmap (%)</div>', unsafe_allow_html=True)
z = m["pct"]
text_vals = [[f"{v:.1f}%" if v > 0 else "—" for v in row] for row in z]
fig_hm = go.Figure(go.Heatmap(
    z=z, x=G, y=G,
    text=text_vals, texttemplate="%{text}",
    colorscale=[
        [0.0,  "rgba(255,255,255,0)"],
        [0.3,  "rgba(59,130,246,0.2)"],
        [0.6,  "rgba(37,99,235,0.5)"],
        [1.0,  "rgba(30,58,138,0.9)"],
    ],
    showscale=True,
    colorbar=dict(title="%", thickness=12, len=0.8),
    hovertemplate="<b>%{y} → %{x}</b><br>%{z:.1f}%<extra></extra>",
))
fig_hm.update_layout(
    height=300, margin=dict(l=80, r=40, t=10, b=60),
    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(title="To grade", side="bottom"),
    yaxis=dict(title="From grade", autorange="reversed"),
)
st.plotly_chart(fig_hm, use_container_width=True, config={"displayModeBar": False})

# ── Sankey ────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">Migration flow — Sankey diagram ($M)</div>', unsafe_allow_html=True)
n = len(G)
sankey_colors = GRADE_COLORS
node_labels = [f"{g} (from)" for g in G] + [f"{g} (to)" for g in G]
sources, targets, values, link_colors = [], [], [], []
for i in range(n):
    for j in range(n):
        v = m["raw"][i, j]
        if v > 0.5:
            sources.append(i)
            targets.append(n + j)
            values.append(round(v, 2))
            link_colors.append(sankey_colors[i].replace("#", "") )

node_colors = GRADE_COLORS * 2
fig_sk = go.Figure(go.Sankey(
    node=dict(
        pad=15, thickness=18, line=dict(color="rgba(0,0,0,0.1)", width=0.5),
        label=node_labels, color=node_colors,
    ),
    link=dict(
        source=sources, target=targets, value=values,
        color=[f"rgba({int(sankey_colors[s][1:3],16)},{int(sankey_colors[s][3:5],16)},{int(sankey_colors[s][5:7],16)},0.25)" for s in sources],
        hovertemplate="<b>%{source.label} → %{target.label}</b><br>$%{value:.1f}M<extra></extra>",
    ),
))
fig_sk.update_layout(
    height=320, margin=dict(l=10, r=10, t=10, b=10),
    paper_bgcolor="rgba(0,0,0,0)", font_size=11,
)
st.plotly_chart(fig_sk, use_container_width=True, config={"displayModeBar": False})

# ── Donut chart ───────────────────────────────────────────────────────────────
col_d1, col_d2 = st.columns(2)

with col_d1:
    st.markdown('<div class="section-header">Closing balance by grade ($M)</div>', unsafe_allow_html=True)
    fig_pie = go.Figure(go.Pie(
        labels=G, values=[round(v, 2) for v in m["col_totals"]],
        marker_colors=GRADE_COLORS,
        hole=0.55,
        hovertemplate="<b>%{label}</b><br>$%{value:.1f}M (%{percent})<extra></extra>",
        textinfo="percent",
        textfont_size=12,
    ))
    fig_pie.update_layout(
        height=260, margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="v", x=1.0, font_size=11),
        annotations=[dict(text=f"${m['grand_total']:,.0f}M", x=0.5, y=0.5,
                          font_size=14, showarrow=False, font_color="#212529")],
    )
    st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})

with col_d2:
    st.markdown('<div class="section-header">Net migration per grade ($M, downgrade is negative)</div>', unsafe_allow_html=True)
    net = m["upgrade"] - m["downgrade"]
    bar_colors = ["#059669" if v >= 0 else "#dc2626" for v in net]
    fig_net = go.Figure(go.Bar(
        x=G, y=[round(v, 2) for v in net],
        marker_color=bar_colors, marker_line_width=0,
        hovertemplate="<b>%{x}</b><br>Net: $%{y:.1f}M<extra></extra>",
        text=[f"${v:.1f}M" for v in net],
        textposition="outside",
        textfont_size=11,
    ))
    fig_net.add_hline(y=0, line_width=1, line_color="rgba(0,0,0,0.2)")
    fig_net.update_layout(
        height=260, margin=dict(l=10, r=10, t=30, b=30),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        yaxis=dict(gridcolor="rgba(0,0,0,0.05)", tickprefix="$", ticksuffix="M"),
        xaxis=dict(gridcolor="rgba(0,0,0,0)"),
        showlegend=False,
    )
    st.plotly_chart(fig_net, use_container_width=True, config={"displayModeBar": False})

# ── Early Warning Indicators ──────────────────────────────────────────────────
st.markdown('<div class="section-header">Early warning indicators</div>', unsafe_allow_html=True)

pct = m["pct"]
raw = m["raw"]
col_t = m["col_totals"]
row_t = m["row_totals"]

ews = [
    (
        "#dc2626" if m["npl_ratio"] > npl_threshold else "#d97706",
        f"NPL ratio: {m['npl_ratio']:.2f}% of total book",
        f"Doubtful ${col_t[3]:.1f}M + Bad ${col_t[4]:.1f}M vs total ${m['grand_total']:,.0f}M — {'BREACH: above' if m['npl_ratio'] > npl_threshold else 'Within'} {npl_threshold}% threshold",
    ),
    (
        "#dc2626" if m["roll_fwd"][1] > roll_threshold else "#d97706",
        f"Watchlist roll-forward rate: {m['roll_fwd'][1]:.1f}%",
        f"${m['downgrade'][1]:.1f}M migrating to Substandard or below — key leading indicator for NPL build-up",
    ),
    (
        "#dc2626" if pct[2, 3] > 35 else "#d97706",
        f"Substandard → Doubtful migration: {pct[2, 3]:.1f}%",
        f"${raw[2, 3]:.1f}M rolling into Doubtful — provision shortfall risk if cure rates are low",
    ),
    (
        "#dc2626" if m["retention"][3] < retention_floor else "#d97706",
        f"Doubtful retention rate: {m['retention'][3]:.1f}%",
        f"Only {pct[3, 0]:.1f}% curing to Good; {pct[3, 4]:.1f}% rolling to Bad — resolution pipeline is thin",
    ),
    (
        "#dc2626" if pct[3, 4] > 10 else "#d97706",
        f"Doubtful → Bad migration: {pct[3, 4]:.1f}%",
        f"${raw[3, 4]:.1f}M converting to Bad — write-off pressure building; charge-off ratio at risk",
    ),
    (
        "#d97706" if m["retention"][4] > 80 else "#059669",
        f"Bad grade retention: {m['retention'][4]:.1f}%",
        f"${raw[4, 4]:.1f}M remaining in Bad — recovery strategy and resolution effectiveness needs review",
    ),
    (
        "#059669",
        f"Good grade retention: {m['retention'][0]:.1f}%",
        f"${raw[0, 0]:.1f}M of performing book staying in Good — core portfolio stable",
    ),
]

ew_cols = st.columns(2)
for idx, (color, title, desc) in enumerate(ews):
    with ew_cols[idx % 2]:
        st.markdown(
            f'<div class="ew-card" style="border-left-color:{color}">'
            f'<div class="ew-title">{title}</div>'
            f'<div class="ew-desc">{desc}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

# ── Raw data table ─────────────────────────────────────────────────────────────
with st.expander("View raw transition matrix ($M)"):
    df_raw = pd.DataFrame(m["raw"], index=G, columns=G)
    df_raw["Row Total"] = m["row_totals"]
    st.dataframe(df_raw.style.format("${:,.3f}M").background_gradient(cmap="Blues", axis=None), use_container_width=True)

with st.expander("View % transition matrix"):
    df_pct = pd.DataFrame(m["pct"], index=G, columns=G)
    st.dataframe(df_pct.style.format("{:.1f}%").background_gradient(cmap="RdYlGn_r", axis=None), use_container_width=True)

# ── Export CSV ────────────────────────────────────────────────────────────────
if export_btn:
    buf = io.StringIO()
    summary = pd.DataFrame({
        "Grade": G,
        "Opening Balance ($M)": [round(v, 3) for v in m["row_totals"]],
        "Closing Balance ($M)": [round(v, 3) for v in m["col_totals"]],
        "Retention Rate (%)": [round(v, 2) for v in m["retention"]],
        "Roll-Forward Rate (%)": [round(v, 2) for v in m["roll_fwd"]],
        "Downgrade Flow ($M)": [round(v, 3) for v in m["downgrade"]],
        "Upgrade Flow ($M)": [round(v, 3) for v in m["upgrade"]],
        "Net Migration ($M)": [round(v, 3) for v in (m["upgrade"] - m["downgrade"])],
    })
    summary.to_csv(buf, index=False)
    st.download_button("Download CSV", buf.getvalue(), "loan_transition_summary.csv", "text/csv")

st.markdown("---")
st.caption("Loan Quality Transition Dashboard · Built with Streamlit + Plotly · Data: Template.xlsx")
