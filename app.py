import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings('ignore')

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sell-Out Dashboard 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #F8F4F0;
    border-right: 1px solid #E8DDD5;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stRadio label,
[data-testid="stSidebar"] p {
    color: #3D2B1F !important;
    font-size: 0.82rem;
    font-weight: 500;
    letter-spacing: 0.03em;
    text-transform: uppercase;
}

/* Main background */
.stApp { background: #FAFAF8; }

/* Metric cards */
[data-testid="metric-container"] {
    background: #FFFFFF;
    border: 1px solid #EDE8E3;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
[data-testid="stMetricValue"] {
    font-family: 'DM Mono', monospace !important;
    font-size: 1.6rem !important;
    color: #1A1A1A !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
    color: #8B7355 !important;
}
[data-testid="stMetricDelta"] { font-size: 0.85rem !important; }

/* Section headers */
.section-header {
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #8B7355;
    border-bottom: 1px solid #EDE8E3;
    padding-bottom: 6px;
    margin: 24px 0 16px 0;
}

/* Page title */
.page-title {
    font-size: 1.6rem;
    font-weight: 700;
    color: #1A1A1A;
    letter-spacing: -0.02em;
    margin-bottom: 2px;
}
.page-subtitle {
    font-size: 0.85rem;
    color: #8B7355;
    font-weight: 400;
    margin-bottom: 24px;
}

/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: #F0EBE5;
    padding: 4px;
    border-radius: 10px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 7px;
    font-size: 0.82rem;
    font-weight: 600;
    letter-spacing: 0.02em;
    color: #8B7355;
    padding: 6px 16px;
}
.stTabs [aria-selected="true"] {
    background: #FFFFFF !important;
    color: #1A1A1A !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}

/* Divider */
hr { border-color: #EDE8E3; }

/* Plotly chart border */
.js-plotly-plot {
    border-radius: 12px;
    border: 1px solid #EDE8E3;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)

# ── Colours ────────────────────────────────────────────────────────────────────
C = {
    "brand":   "#C75B7A",
    "accent":  "#D4845A",
    "green":   "#3D9B8A",
    "red":     "#C0392B",
    "warm":    "#8B7355",
    "grey":    "#B0A090",
    "bg":      "#FFFFFF",
    "card":    "#FAFAF8",
    "grid":    "#EDE8E3",
    "text":    "#1A1A1A",
    "text2":   "#8B7355",
}

AX = dict(
    gridcolor=C["grid"], gridwidth=0.8,
    zerolinecolor=C["grid"],
    tickfont=dict(color=C["text2"], size=10),
    showline=True, linecolor=C["grid"],
)

CURRENT_WEEK = 18
CURRENT_YEAR = 2026
PREV_YEAR    = 2025

# ── Data loading ───────────────────────────────────────────────────────────────
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name="data", engine="openpyxl")
    df.columns = ["LL","Store","Family","Group","Year","Week","Revenue","Units","WeekNum2","LL2"]
    df = df[df["Year"].notna() & df["Revenue"].notna()].copy()
    df["Year"]  = df["Year"].astype(int)
    df["Week"]  = df["Week"].astype(int)
    return df

def ytd(data, year, week=CURRENT_WEEK):
    return data[(data["Year"] == year) & (data["Week"] <= week)]

# ── Chart helpers ──────────────────────────────────────────────────────────────
def apply_layout(fig, title="", height=420):
    fig.update_xaxes(**AX)
    fig.update_yaxes(**AX)
    fig.update_layout(
        title={"text": title, "font": {"size": 13, "color": C["text"], "family": "DM Sans"},
               "x": 0.01, "pad": {"b": 10}} if title else {},
        paper_bgcolor=C["bg"],
        plot_bgcolor=C["card"],
        font={"family": "DM Sans", "color": C["text"]},
        height=height,
        margin={"l": 10, "r": 20, "t": 40 if title else 20, "b": 10},
        legend={"font": {"size": 11, "color": C["text2"]}, "bgcolor": "rgba(0,0,0,0)",
                "orientation": "h", "x": 0, "y": 1.12},
        hoverlabel={"bgcolor": "#fff", "bordercolor": C["grid"],
                    "font_size": 12, "font_family": "DM Sans"},
    )
    return fig

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 📊 Sell-Out Dashboard")
    st.markdown("---")

    uploaded = st.file_uploader(
        "ΑΝΕΒΑΣΕ ΤΟ ΑΡΧΕΙΟ .XLSX",
        type=["xlsx"],
        help="ΕΒΔΟΜΑΔΙΑΙΟ_SELL_OUT_ΠΡΟΪΟΝΤΩΝ_2026_18η.xlsx",
    )

    if uploaded:
        st.success("✅ Φορτώθηκε!")
        st.markdown("---")
        st.markdown("**ΦΙΛΤΡΑ**")

if not uploaded:
    st.markdown('<div class="page-title">📊 Sell-Out Dashboard 2026</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">Ανέβασε το αρχείο Excel από το sidebar για να ξεκινήσεις.</div>', unsafe_allow_html=True)
    st.info("👈 Χρησιμοποίησε το **sidebar** για να ανεβάσεις το αρχείο ΕΒΔΟΜΑΔΙΑΙΟ_SELL_OUT_ΠΡΟΪΟΝΤΩΝ_2026_18η.xlsx")
    st.stop()

# ── Load ───────────────────────────────────────────────────────────────────────
df = load_data(uploaded)

all_stores  = sorted(df["Store"].unique())
all_families = sorted(df["Family"].unique())

with st.sidebar:
    view_mode = st.radio(
        "ΠΡΟΒΟΛΗ",
        ["🌐 Σύνολο Δικτύου", "🏪 Ανά Κατάστημα"],
        index=0,
    )
    st.markdown("---")

    if view_mode == "🏪 Ανά Κατάστημα":
        selected_store = st.selectbox(
            "ΚΑΤΑΣΤΗΜΑ",
            all_stores,
            index=0,
        )
    else:
        selected_store = None

    selected_families = st.multiselect(
        "ΟΙΚΟΓΕΝΕΙΕΣ ΕΙΔΩΝ",
        all_families,
        default=all_families,
    )
    st.markdown("---")
    st.caption(f"Εβδομάδα **{CURRENT_WEEK}** / 2026")
    st.caption(f"Καταστήματα: **{df['Store'].nunique()}**")
    st.caption(f"Εγγραφές: **{len(df):,}**")

# ── Filter by family ───────────────────────────────────────────────────────────
if selected_families:
    dff = df[df["Family"].isin(selected_families)]
else:
    dff = df.copy()

cy_all = ytd(dff, CURRENT_YEAR)
py_all = ytd(dff, PREV_YEAR)

# ══════════════════════════════════════════════════════════════════════════════
# TAB LAYOUT
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["🌐 Σύνολο Δικτύου", "🏪 Ανά Κατάστημα"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — NETWORK OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown(f'<div class="page-title">Σύνολο Δικτύου</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">YTD Εβδομάδα {CURRENT_WEEK} · 2026 vs 2025</div>', unsafe_allow_html=True)

    # KPIs
    rev_cy   = cy_all["Revenue"].sum()
    rev_py   = py_all["Revenue"].sum()
    units_cy = cy_all["Units"].sum()
    units_py = py_all["Units"].sum()
    stores_n = cy_all["Store"].nunique()
    avg_store= rev_cy / stores_n if stores_n else 0
    avg_py   = rev_py / py_all["Store"].nunique() if py_all["Store"].nunique() else 0

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Τζίρος YTD 2026",    f"€{rev_cy:,.0f}",   f"{(rev_cy-rev_py)/rev_py*100:+.1f}% vs 2025")
    k2.metric("Τεμάχια YTD 2026",   f"{units_cy:,.0f}",  f"{(units_cy-units_py)/units_py*100:+.1f}% vs 2025")
    k3.metric("Ενεργά Καταστήματα", f"{stores_n}",        "")
    k4.metric("Μέσος Τζίρος/Store", f"€{avg_store:,.0f}", f"{(avg_store-avg_py)/avg_py*100:+.1f}% vs 2025")

    st.markdown('<div class="section-header">Εβδομαδιαία Εξέλιξη Τζίρου</div>', unsafe_allow_html=True)

    # Weekly evolution
    weekly = dff[dff["Year"].isin([2025,2026])].groupby(["Year","Week"])["Revenue"].sum().reset_index()
    w25 = weekly[weekly["Year"]==2025].sort_values("Week")
    w26 = weekly[weekly["Year"]==2026].sort_values("Week")

    fig_weekly = go.Figure()
    fig_weekly.add_trace(go.Scatter(
        x=w25["Week"], y=w25["Revenue"], name="2025",
        mode="lines", line={"color": C["grey"], "width": 1.8, "dash": "dot"},
        hovertemplate="Εβδ. %{x}<br>€%{y:,.0f}<extra>2025</extra>",
    ))
    fig_weekly.add_trace(go.Scatter(
        x=w26["Week"], y=w26["Revenue"], name="2026",
        mode="lines+markers",
        line={"color": C["brand"], "width": 2.5},
        marker={"size": 7, "color": C["brand"], "line": {"color": "#fff", "width": 1.5}},
        fill="tozeroy", fillcolor="rgba(199,91,122,0.08)",
        hovertemplate="Εβδ. %{x}<br>€%{y:,.0f}<extra>2026</extra>",
    ))
    apply_layout(fig_weekly, height=320)
    fig_weekly.update_layout(
        yaxis=dict(tickprefix="€", tickformat=",.0f", **AX),
        xaxis=dict(title="Εβδομάδα", **AX),
    )
    st.plotly_chart(fig_weekly, use_container_width=True)

    col_a, col_b = st.columns([1.4, 1])

    with col_a:
        st.markdown('<div class="section-header">Τζίρος ανά Οικογένεια Ειδών</div>', unsafe_allow_html=True)
        fam_cy = cy_all.groupby("Family")["Revenue"].sum().sort_values(ascending=False)
        fam_py = py_all.groupby("Family")["Revenue"].sum().reindex(fam_cy.index).fillna(0)

        fig_fam = go.Figure()
        fig_fam.add_trace(go.Bar(
            name="2025", x=fam_cy.index, y=fam_py.values,
            marker_color=C["grey"], opacity=0.45,
            hovertemplate="%{x}<br>€%{y:,.0f}<extra>2025</extra>",
        ))
        fig_fam.add_trace(go.Bar(
            name="2026", x=fam_cy.index, y=fam_cy.values,
            marker_color=C["brand"],
            hovertemplate="%{x}<br>€%{y:,.0f}<extra>2026</extra>",
        ))
        apply_layout(fig_fam, height=320)
        fig_fam.update_layout(
            barmode="group",
            yaxis=dict(tickprefix="€", tickformat=",.0f", **AX),
            xaxis=dict(tickangle=-30, **AX),
        )
        st.plotly_chart(fig_fam, use_container_width=True)

    with col_b:
        st.markdown('<div class="section-header">L/L Mix</div>', unsafe_allow_html=True)
        ll_mix = cy_all.groupby("LL")["Revenue"].sum()
        fig_pie = go.Figure(go.Pie(
            labels=ll_mix.index, values=ll_mix.values,
            hole=0.60,
            marker={"colors": [C["green"], C["grey"]]},
            textinfo="percent+label",
            textfont={"size": 12, "color": C["text"]},
            hovertemplate="%{label}<br>€%{value:,.0f}<br>%{percent}<extra></extra>",
        ))
        apply_layout(fig_pie, height=320)
        st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown('<div class="section-header">Top 15 Καταστήματα · Τζίρος YTD 2026</div>', unsafe_allow_html=True)

    col_c, col_d = st.columns(2)

    with col_c:
        top15 = cy_all.groupby("Store")["Revenue"].sum().sort_values(ascending=True).tail(15).reset_index()
        fig_top = go.Figure(go.Bar(
            x=top15["Revenue"], y=top15["Store"],
            orientation="h",
            marker=dict(
                color=top15["Revenue"],
                colorscale=[[0, "#F5DDD9"], [1, C["brand"]]],
                showscale=False,
            ),
            text=top15["Revenue"].apply(lambda x: f"€{x:,.0f}"),
            textposition="outside",
            textfont={"size": 9, "color": C["text"]},
            hovertemplate="%{y}<br>€%{x:,.0f}<extra></extra>",
        ))
        apply_layout(fig_top, height=420)
        fig_top.update_layout(
            xaxis=dict(tickprefix="€", tickformat=",.0f", **AX),
            yaxis=dict(**AX),
            margin={"l": 160, "r": 80, "t": 10, "b": 10},
        )
        st.plotly_chart(fig_top, use_container_width=True)

    with col_d:
        st.markdown('<div class="section-header" style="margin-top:0">YoY % Δ · Top 20 (by τζίρο 2026)</div>', unsafe_allow_html=True)
        st_cy2 = cy_all.groupby("Store")["Revenue"].sum()
        st_py2 = py_all.groupby("Store")["Revenue"].sum()
        cmp = pd.DataFrame({"CY": st_cy2, "PY": st_py2}).dropna()
        cmp["Delta"] = (cmp["CY"] - cmp["PY"]) / cmp["PY"] * 100
        top20 = cmp.nlargest(20, "CY").sort_values("Delta")
        colors_d = [C["green"] if v >= 0 else C["red"] for v in top20["Delta"]]

        fig_yoy = go.Figure(go.Bar(
            x=top20["Delta"], y=top20.index,
            orientation="h",
            marker_color=colors_d,
            text=top20["Delta"].apply(lambda x: f"{x:+.1f}%"),
            textposition="outside",
            textfont={"size": 9, "color": C["text"]},
            hovertemplate="%{y}<br>%{x:+.1f}%<extra></extra>",
        ))
        apply_layout(fig_yoy, height=420)
        fig_yoy.update_layout(
            xaxis=dict(ticksuffix="%", **AX),
            yaxis=dict(**AX),
            margin={"l": 160, "r": 70, "t": 10, "b": 10},
        )
        fig_yoy.add_vline(x=0, line_color=C["grid"], line_width=1.5)
        st.plotly_chart(fig_yoy, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PER STORE
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    # Store selector at top of tab
    col_sel, col_info = st.columns([2, 3])
    with col_sel:
        store = st.selectbox(
            "Επίλεξε κατάστημα",
            all_stores,
            key="tab2_store",
            label_visibility="collapsed",
        )

    s     = dff[dff["Store"] == store]
    s_cy  = ytd(s, CURRENT_YEAR)
    s_py  = ytd(s, PREV_YEAR)

    rev_s_cy  = s_cy["Revenue"].sum()
    rev_s_py  = s_py["Revenue"].sum()
    unit_s_cy = s_cy["Units"].sum()
    unit_s_py = s_py["Units"].sum()
    chg_rev   = (rev_s_cy - rev_s_py) / rev_s_py * 100 if rev_s_py else 0
    chg_u     = (unit_s_cy - unit_s_py) / unit_s_py * 100 if unit_s_py else 0
    contrib   = rev_s_cy / cy_all["Revenue"].sum() * 100 if cy_all["Revenue"].sum() else 0

    # Rank
    rank_df = cy_all.groupby("Store")["Revenue"].sum().sort_values(ascending=False).reset_index()
    rank_df["Rank"] = range(1, len(rank_df)+1)
    store_rank = rank_df[rank_df["Store"]==store]["Rank"].values
    rank_str = f"#{store_rank[0]} / {len(rank_df)}" if len(store_rank) else "—"

    with col_info:
        st.markdown(f'<div class="page-title">{store}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="page-subtitle">YTD Εβδομάδα {CURRENT_WEEK} · Κατάταξη {rank_str} στο δίκτυο</div>', unsafe_allow_html=True)

    # KPIs
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Τζίρος YTD 2026",  f"€{rev_s_cy:,.0f}",  f"{chg_rev:+.1f}% vs 2025")
    k2.metric("Τεμάχια YTD 2026", f"{unit_s_cy:,.0f}",   f"{chg_u:+.1f}% vs 2025")
    k3.metric("Contribution %",   f"{contrib:.2f}%",      "επί του δικτύου")
    k4.metric("Κατάταξη Δικτύου", rank_str,               "")

    st.markdown('<div class="section-header">Εβδομαδιαία Εξέλιξη</div>', unsafe_allow_html=True)

    col_w, col_cum = st.columns(2)

    with col_w:
        w_s25 = s[s["Year"]==2025].groupby("Week")["Revenue"].sum().reset_index()
        w_s26 = s[s["Year"]==2026].groupby("Week")["Revenue"].sum().reset_index()

        fig_sw = go.Figure()
        fig_sw.add_trace(go.Bar(
            x=w_s25["Week"], y=w_s25["Revenue"], name="2025",
            marker_color=C["grey"], opacity=0.4,
            hovertemplate="Εβδ. %{x}<br>€%{y:,.0f}<extra>2025</extra>",
        ))
        fig_sw.add_trace(go.Scatter(
            x=w_s26["Week"], y=w_s26["Revenue"], name="2026",
            mode="lines+markers",
            line={"color": C["brand"], "width": 2.5},
            marker={"size": 7, "color": C["brand"], "line": {"color": "#fff", "width": 1.5}},
            hovertemplate="Εβδ. %{x}<br>€%{y:,.0f}<extra>2026</extra>",
        ))
        apply_layout(fig_sw, "Εβδομαδιαίος Τζίρος", height=300)
        fig_sw.update_layout(
            yaxis=dict(tickprefix="€", tickformat=",.0f", **AX),
            xaxis=dict(title="Εβδομάδα", **AX),
        )
        st.plotly_chart(fig_sw, use_container_width=True)

    with col_cum:
        cum25 = w_s25.set_index("Week")["Revenue"].cumsum().reset_index()
        cum26 = w_s26.set_index("Week")["Revenue"].cumsum().reset_index()

        fig_cum = go.Figure()
        fig_cum.add_trace(go.Scatter(
            x=cum25["Week"], y=cum25["Revenue"], name="2025",
            mode="lines", line={"color": C["grey"], "width": 1.8, "dash": "dot"},
            hovertemplate="Εβδ. %{x}<br>€%{y:,.0f}<extra>2025</extra>",
        ))
        fig_cum.add_trace(go.Scatter(
            x=cum26["Week"], y=cum26["Revenue"], name="2026",
            mode="lines+markers",
            line={"color": C["green"], "width": 2.5},
            marker={"size": 7, "color": C["green"], "line": {"color": "#fff", "width": 1.5}},
            fill="tozeroy", fillcolor="rgba(61,155,138,0.08)",
            hovertemplate="Εβδ. %{x}<br>€%{y:,.0f}<extra>2026</extra>",
        ))
        apply_layout(fig_cum, "Σωρευτικός Τζίρος", height=300)
        fig_cum.update_layout(
            yaxis=dict(tickprefix="€", tickformat=",.0f", **AX),
            xaxis=dict(title="Εβδομάδα", **AX),
        )
        st.plotly_chart(fig_cum, use_container_width=True)

    st.markdown('<div class="section-header">Ανάλυση ανά Οικογένεια & Ομάδα</div>', unsafe_allow_html=True)

    col_f, col_g = st.columns(2)

    with col_f:
        fam_s_cy = s_cy.groupby("Family")["Revenue"].sum().sort_values(ascending=True)
        fam_s_py = s_py.groupby("Family")["Revenue"].sum().reindex(fam_s_cy.index).fillna(0)

        fig_fams = go.Figure()
        fig_fams.add_trace(go.Bar(
            x=fam_s_py.values, y=fam_s_py.index,
            orientation="h", name="2025",
            marker_color=C["grey"], opacity=0.45,
            hovertemplate="%{y}<br>€%{x:,.0f}<extra>2025</extra>",
        ))
        fig_fams.add_trace(go.Bar(
            x=fam_s_cy.values, y=fam_s_cy.index,
            orientation="h", name="2026",
            marker_color=C["brand"],
            text=fam_s_cy.values.astype(int),
            texttemplate="€%{text:,}",
            textposition="outside",
            textfont={"size": 9},
            hovertemplate="%{y}<br>€%{x:,.0f}<extra>2026</extra>",
        ))
        apply_layout(fig_fams, "Τζίρος ανά Οικογένεια", height=340)
        fig_fams.update_layout(
            barmode="overlay",
            xaxis=dict(tickprefix="€", tickformat=",.0f", **AX),
            yaxis=dict(**AX),
            margin={"l": 180, "r": 90, "t": 40, "b": 10},
        )
        st.plotly_chart(fig_fams, use_container_width=True)

    with col_g:
        grp_s = s_cy.groupby("Group")["Revenue"].sum().sort_values(ascending=True).tail(10)
        fig_grp = go.Figure(go.Bar(
            x=grp_s.values, y=grp_s.index,
            orientation="h",
            marker=dict(
                color=grp_s.values,
                colorscale=[[0,"#FAE5D3"],[1, C["accent"]]],
                showscale=False,
            ),
            text=grp_s.values.astype(int),
            texttemplate="€%{text:,}",
            textposition="outside",
            textfont={"size": 9},
            hovertemplate="%{y}<br>€%{x:,.0f}<extra></extra>",
        ))
        apply_layout(fig_grp, "Top 10 Ομάδες Ειδών", height=340)
        fig_grp.update_layout(
            xaxis=dict(tickprefix="€", tickformat=",.0f", **AX),
            yaxis=dict(**AX),
            margin={"l": 180, "r": 90, "t": 40, "b": 10},
        )
        st.plotly_chart(fig_grp, use_container_width=True)

    # Units weekly
    st.markdown('<div class="section-header">Τεμάχια ανά Εβδομάδα</div>', unsafe_allow_html=True)
    wu25 = s[s["Year"]==2025].groupby("Week")["Units"].sum().reset_index()
    wu26 = s[s["Year"]==2026].groupby("Week")["Units"].sum().reset_index()

    fig_units = go.Figure()
    fig_units.add_trace(go.Scatter(
        x=wu25["Week"], y=wu25["Units"], name="2025",
        mode="lines", line={"color": C["grey"], "width": 1.5, "dash": "dot"},
    ))
    fig_units.add_trace(go.Bar(
        x=wu26["Week"], y=wu26["Units"], name="2026",
        marker_color=C["accent"], opacity=0.85,
    ))
    apply_layout(fig_units, height=240)
    fig_units.update_layout(
        barmode="overlay",
        xaxis=dict(title="Εβδομάδα", **AX),
        yaxis=dict(title="Τεμάχια", **AX),
    )
    st.plotly_chart(fig_units, use_container_width=True)
