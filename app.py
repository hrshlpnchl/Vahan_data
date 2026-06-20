#!/usr/bin/env python3
"""
app.py — India EV & Vehicle Registration Dashboard
Reads master.parquet (compiled from VAHAN xlsx files).
Runs on Streamlit Cloud. Fast because it's one file, not 216.

Run locally:  streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import json
import urllib.request
from pathlib import Path
from datetime import datetime

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG  (must be first Streamlit call)
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="India EV Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════
# THEME — auto-adapts to light/dark, single source of truth
# ══════════════════════════════════════════════════════════════
# Streamlit exposes the active theme (light/dark) via st.context.theme.
# We use that to pick ONE matching palette, then push it everywhere:
# CSS variables (for native widgets) AND a plain dict (for Plotly,
# which can't read CSS variables). This is the fix for the
# "fonts disappear in light mode" bug — nothing is hardcoded anymore.
try:
    is_dark = st.context.theme.type == "dark"
except Exception:
    # Fallback for older Streamlit versions without st.context.theme
    is_dark = True

if is_dark:
    THEME = {
        "bg":           "#0E1117",
        "surface":      "#161B22",
        "border":       "#2A3038",
        "text":         "#E6E8EB",
        "text_muted":   "#9AA4AF",
        "accent":       "#22C55E",   # green — primary brand / "good"
        "accent_soft":  "rgba(34,197,94,0.15)",
        "accent2":      "#3B82F6",   # blue — secondary series
        "negative":     "#EF4444",   # red — decline / warning
        "grid":         "#262C36",
    }
else:
    THEME = {
        "bg":           "#FFFFFF",
        "surface":      "#F6F7F9",
        "border":       "#E3E6EA",
        "text":         "#1A1D21",
        "text_muted":   "#5B6470",
        "accent":       "#16A34A",   # slightly deeper green for contrast on white
        "accent_soft":  "rgba(22,163,74,0.10)",
        "accent2":      "#2563EB",
        "negative":     "#DC2626",
        "grid":         "#EAECEF",
    }

# Small, restrained categorical palette (was 20 neon colors → now 6,
# all chosen to work on both light and dark backgrounds).
COLOR_PALETTE = [
    THEME["accent"], THEME["accent2"], "#F59E0B",
    "#A855F7", "#06B6D4", "#EC4899",
]

# Single gradient scale, reused everywhere a gradient is genuinely
# needed (choropleth map, heatmaps). One scale = one visual language
# instead of four competing ones.
SCALE_MAIN = [[0, THEME["surface"]], [0.5, THEME["accent2"]], [1, THEME["accent"]]]
SCALE_HEAT = [[0, THEME["surface"]], [0.5, "#F59E0B"], [1, THEME["negative"]]]

VEHICLE_LABELS = {
    "2W": "2-Wheeler",
    "3W": "3-Wheeler",
    "4W": "4-Wheeler (LMV)",
}
FUEL_LABELS = {"PureEV": "Electric Only", "AllFuel": "All Fuel Types"}

STATE_NAME_MAP = {
    "Andaman & Nicobar Island": "Andaman & Nicobar Island",
    "Andhra Pradesh": "Andhra Pradesh",
    "Arunachal Pradesh": "Arunachal Pradesh",
    "Assam": "Assam",
    "Bihar": "Bihar",
    "Chandigarh": "Chandigarh",
    "Chhattisgarh": "Chhattisgarh",
    "UT of DNH and DD": "Dadra and Nagar Haveli and Daman and Diu",
    "Delhi": "NCT of Delhi",
    "Goa": "Goa",
    "Gujarat": "Gujarat",
    "Haryana": "Haryana",
    "Himachal Pradesh": "Himachal Pradesh",
    "Jammu and Kashmir": "Jammu & Kashmir",
    "Jharkhand": "Jharkhand",
    "Karnataka": "Karnataka",
    "Kerala": "Kerala",
    "Ladakh": "Ladakh",
    "Lakshadweep": "Lakshadweep",
    "Madhya Pradesh": "Madhya Pradesh",
    "Maharashtra": "Maharashtra",
    "Manipur": "Manipur",
    "Meghalaya": "Meghalaya",
    "Mizoram": "Mizoram",
    "Nagaland": "Nagaland",
    "Odisha": "Odisha",
    "Puducherry": "Puducherry",
    "Punjab": "Punjab",
    "Rajasthan": "Rajasthan",
    "Sikkim": "Sikkim",
    "Tamil Nadu": "Tamil Nadu",
    "Telangana": "Telangana",
    "Tripura": "Tripura",
    "Uttar Pradesh": "Uttar Pradesh",
    "Uttarakhand": "Uttarakhand",
    "West Bengal": "West Bengal",
}

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
          "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
MONTH_LABELS = {
    "JAN": "Jan", "FEB": "Feb", "MAR": "Mar", "APR": "Apr",
    "MAY": "May", "JUN": "Jun", "JUL": "Jul", "AUG": "Aug",
    "SEP": "Sep", "OCT": "Oct", "NOV": "Nov", "DEC": "Dec",
}


def plotly_layout(**overrides):
    """Every chart pulls from the SAME theme dict, so charts always
    match the surrounding page — in both light and dark mode."""
    base = dict(
        paper_bgcolor=THEME["bg"],
        plot_bgcolor=THEME["bg"],
        font=dict(color=THEME["text"], family="Inter, -apple-system, sans-serif", size=12),
        title_font=dict(size=14, color=THEME["text"]),
        legend=dict(font=dict(color=THEME["text_muted"], size=11)),
        xaxis=dict(gridcolor=THEME["grid"], zerolinecolor=THEME["grid"],
                    color=THEME["text_muted"]),
        yaxis=dict(gridcolor=THEME["grid"], zerolinecolor=THEME["grid"],
                    color=THEME["text_muted"]),
        margin=dict(l=10, r=10, t=45, b=10),
        hoverlabel=dict(bgcolor=THEME["surface"], font_color=THEME["text"],
                         bordercolor=THEME["border"]),
    )
    base.update(overrides)
    return base


# ══════════════════════════════════════════════════════════════
# CSS — uses Streamlit's own theme vars + our THEME dict together,
# so native widgets (which Streamlit already themes correctly) are
# left alone, and we only style the custom elements we add.
# ══════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', -apple-system, sans-serif;
}}
.block-container {{
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}}

/* KPI cards */
[data-testid="stMetric"] {{
    background: {THEME["surface"]};
    border: 1px solid {THEME["border"]};
    border-radius: 12px;
    padding: 14px 18px;
}}
[data-testid="stMetricLabel"] {{
    color: {THEME["text_muted"]} !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
}}
[data-testid="stMetricValue"] {{
    font-size: 1.45rem !important;
    font-weight: 700 !important;
    color: {THEME["text"]} !important;
}}

/* Tabs — quiet until selected */
.stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
}}
.stTabs [data-baseweb="tab"] {{
    color: {THEME["text_muted"]};
    font-weight: 500;
}}
.stTabs [aria-selected="true"] {{
    background: {THEME["accent_soft"]} !important;
    color: {THEME["accent"]} !important;
    font-weight: 600 !important;
    border-radius: 8px;
}}

/* Hero header */
.hero {{
    padding: 22px 28px;
    border-radius: 14px;
    background: {THEME["surface"]};
    border: 1px solid {THEME["border"]};
    margin-bottom: 1.2rem;
}}
.hero h1 {{
    margin: 0 0 6px;
    font-size: 1.5rem;
    font-weight: 700;
    color: {THEME["text"]};
}}
.hero p {{
    margin: 0;
    color: {THEME["text_muted"]};
    font-size: 0.92rem;
}}
.hero a {{ color: {THEME["accent"]}; text-decoration: none; font-weight: 600; }}

/* Sidebar info box */
.sidebar-info {{
    font-size: 0.8rem;
    color: {THEME["text_muted"]};
    line-height: 1.6;
}}
.sidebar-info b {{ color: {THEME["text"]}; }}

/* Footer */
.app-footer {{
    text-align: center;
    color: {THEME["text_muted"]};
    font-size: 0.8rem;
    padding: 10px 0;
}}
.app-footer a {{ color: {THEME["accent"]}; text-decoration: none; }}
.app-footer b {{ color: {THEME["text"]}; }}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# DATA LOADING  — single parquet, cached forever until file changes
# ══════════════════════════════════════════════════════════════
PARQUET_PATH = Path(__file__).parent / "master.parquet"


@st.cache_data(show_spinner="Loading data…", ttl=3600)
def load_data() -> pd.DataFrame:
    if not PARQUET_PATH.exists():
        st.error(
            "`master.parquet` not found. "
            "Run `python compile_parquet.py` first to generate it."
        )
        st.stop()
    df = pd.read_parquet(PARQUET_PATH)
    for m in MONTHS:
        if m not in df.columns:
            df[m] = 0
    return df


@st.cache_data(show_spinner=False, ttl=86400)
def load_geojson():
    url = (
        "https://gist.githubusercontent.com/jbrobst/"
        "56c13bbbf9d97d187fea01ca62ea5112/raw/"
        "e388c4cae20aa53cb5090210a42ebb9b765c0a36/india_states.geojson"
    )
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as r:
            return json.loads(r.read().decode())
    except Exception:
        return None


df_master = load_data()

month_sums = {m: int(df_master[m].sum()) for m in MONTHS}
active_months = [m for m in MONTHS if month_sums[m] > 0]

try:
    mtime = PARQUET_PATH.stat().st_mtime
    last_updated = datetime.fromtimestamp(mtime).strftime("%d %b %Y, %H:%M")
except Exception:
    last_updated = "Unknown"


# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:6px 0 12px;">
        <span style="font-size:1.8rem;">⚡</span>
        <h2 style="margin:4px 0 0;color:{THEME['text']};font-size:1.15rem;font-weight:700;">
            Vahan EV Dashboard
        </h2>
        <p style="color:{THEME['text_muted']};font-size:0.78rem;margin:4px 0 0;">
            India Vehicle Registrations &middot; Source: Vahan Portal (MoRTH)
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    fuel_choice = st.radio(
        "Fuel Category",
        options=sorted(df_master["FuelType"].unique()),
        format_func=lambda x: FUEL_LABELS.get(x, x),
    )

    available_vtypes = sorted(df_master["VehicleType"].unique())
    sel_vtypes = st.multiselect(
        "Vehicle Type",
        options=available_vtypes,
        default=available_vtypes,
        format_func=lambda x: VEHICLE_LABELS.get(x, x),
    )
    if not sel_vtypes:
        st.warning("Select at least one vehicle type.")
        st.stop()

    available_years = sorted(df_master["Year"].unique(), reverse=True)
    sel_year = st.selectbox("Year", options=available_years, index=0)

    individual_states = sorted(
        df_master[df_master["State"] != "All India"]["State"].unique()
    )
    sel_states = st.multiselect("States", options=individual_states, default=individual_states)

    top_n = st.slider("Top N Makers", 5, 30, 12)

    sel_months = st.multiselect(
        "Months",
        options=active_months,
        default=active_months,
        format_func=lambda m: MONTH_LABELS.get(m, m),
    )
    if not sel_months:
        sel_months = active_months

    st.divider()
    if st.button("Reload data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown(f"""
    <div class="sidebar-info">
        <b>Data updated:</b> {last_updated}<br>
        <b>Total rows:</b> {len(df_master):,}<br>
        <b>States:</b> {len(individual_states)} &nbsp;|&nbsp;
        <b>Makers:</b> {df_master['Maker'].nunique():,}
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    st.caption("Built by Harshal Panchal · Data: Vahan Portal")

# ══════════════════════════════════════════════════════════════
# FILTER
# ══════════════════════════════════════════════════════════════
base_filter = (
    (df_master["FuelType"] == fuel_choice)
    & (df_master["VehicleType"].isin(sel_vtypes))
    & (df_master["Year"] == sel_year)
)

df_national = df_master[base_filter & (df_master["State"] == "All India")].copy()
if df_national.empty:
    df_national = df_master[base_filter & (df_master["State"] != "All India")].copy()

df_states = df_master[
    base_filter
    & (df_master["State"] != "All India")
    & (df_master["State"].isin(sel_states))
].copy()


# ══════════════════════════════════════════════════════════════
# HEADER + KPIs
# ══════════════════════════════════════════════════════════════
fuel_label = FUEL_LABELS.get(fuel_choice, fuel_choice)
vtype_label = " + ".join(VEHICLE_LABELS.get(v, v) for v in sel_vtypes)

st.markdown(f"""
<div class="hero">
    <h1>⚡ India Vehicle Registration Dashboard</h1>
    <p>
        Maker-wise registrations &middot; <b>{fuel_label}</b> &middot; <b>{vtype_label}</b> &middot;
        Source: <a href="https://vahan.parivahan.gov.in/vahan4dashboard/" target="_blank">Vahan Portal</a>
    </p>
</div>
""", unsafe_allow_html=True)

total_reg = int(df_national[sel_months].sum().sum()) if not df_national.empty else 0
total_makers = df_national["Maker"].nunique()
total_states_count = df_states["State"].nunique()

active_m_data = {m: int(df_national[m].sum()) for m in sel_months if df_national[m].sum() > 0}
active_m_list = list(active_m_data.keys())
if len(active_m_list) >= 2:
    lm, pm = active_m_list[-1], active_m_list[-2]
    lv, pv = active_m_data[lm], active_m_data[pm]
    mom_pct = ((lv - pv) / max(pv, 1)) * 100
    mom_delta = lv - pv
    mom_label = f"Growth ({MONTH_LABELS[pm]} → {MONTH_LABELS[lm]})"
else:
    mom_pct = mom_delta = 0
    mom_label = "Month-on-Month Growth"

# EV penetration rate (PureEV ÷ AllFuel) — kept for use in tab6, not a top KPI anymore
df_allfuel = df_master[
    (df_master["FuelType"] == "AllFuel")
    & (df_master["VehicleType"].isin(sel_vtypes))
    & (df_master["Year"] == sel_year)
    & (df_master["State"] == "All India")
]
if df_allfuel.empty:
    df_allfuel = df_master[
        (df_master["FuelType"] == "AllFuel")
        & (df_master["VehicleType"].isin(sel_vtypes))
        & (df_master["Year"] == sel_year)
        & (df_master["State"] != "All India")
    ]
allfuel_total = int(df_allfuel[sel_months].sum().sum()) if not df_allfuel.empty else 0
pureev_total = int(df_master[
    (df_master["FuelType"] == "PureEV")
    & (df_master["VehicleType"].isin(sel_vtypes))
    & (df_master["Year"] == sel_year)
][sel_months].sum().sum())
penetration = (pureev_total / max(allfuel_total, 1)) * 100

# 4 KPI cards instead of 6 — kept to what a dealer/OEM person scans in 3 seconds.
k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Registrations", f"{total_reg:,}")
k2.metric("Active Makers", f"{total_makers}")
k3.metric("States Covered", f"{total_states_count}")
k4.metric(mom_label, f"{mom_pct:+.1f}%", delta=f"{mom_delta:+,}")

st.divider()

# ══════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "National Overview",
    "State Analysis",
    "Maker Deep-Dive",
    "Monthly Trends",
    "Heatmap & Rankings",
    "Market Intelligence",
    "Data Explorer",
])


# ─── TAB 1: NATIONAL OVERVIEW ────────────────────────────────
with tab1:
    geojson = load_geojson()
    if geojson and not df_states.empty:
        state_totals = df_states.groupby("State")[sel_months].sum().sum(axis=1).reset_index()
        state_totals.columns = ["State", "TOTAL"]
        state_totals["GeoName"] = state_totals["State"].map(STATE_NAME_MAP)
        state_totals = state_totals.dropna(subset=["GeoName"])

        if not state_totals.empty:
            fig = px.choropleth(
                state_totals, geojson=geojson,
                featureidkey="properties.ST_NM",
                locations="GeoName", color="TOTAL",
                color_continuous_scale=SCALE_MAIN,
                hover_name="State",
                hover_data={"TOTAL": ":,.0f", "GeoName": False},
                title="State-wise Registrations — India Map",
            )
            fig.update_geos(fitbounds="locations", visible=False, bgcolor=THEME["bg"])
            fig.update_layout(**plotly_layout(height=500, margin=dict(l=0, r=0, t=45, b=0)))
            st.plotly_chart(fig, use_container_width=True)

    st.divider()
    c1, c2 = st.columns(2)

    with c1:
        maker_tot = (df_national.groupby("Maker")[sel_months].sum()
                     .sum(axis=1).nlargest(top_n).reset_index())
        maker_tot.columns = ["Maker", "TOTAL"]
        if not maker_tot.empty:
            fig = px.bar(maker_tot, x="TOTAL", y="Maker", orientation="h",
                         title=f"Top {top_n} Makers — {fuel_label} (National)")
            fig.update_traces(marker_color=THEME["accent"])
            fig.update_layout(**plotly_layout(
                height=460, showlegend=False,
                yaxis=dict(autorange="reversed", gridcolor=THEME["grid"], color=THEME["text_muted"]),
            ))
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        state_bar = (df_states.groupby("State")[sel_months].sum()
                     .sum(axis=1).sort_values(ascending=False).reset_index())
        state_bar.columns = ["State", "TOTAL"]
        if not state_bar.empty:
            fig = px.bar(state_bar.head(15), x="TOTAL", y="State", orientation="h",
                         title=f"Top 15 States — {fuel_label}")
            fig.update_traces(marker_color=THEME["accent2"])
            fig.update_layout(**plotly_layout(
                height=460, showlegend=False,
                yaxis=dict(autorange="reversed", gridcolor=THEME["grid"], color=THEME["text_muted"]),
            ))
            st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        top10 = (df_national.groupby("Maker")[sel_months].sum()
                 .sum(axis=1).nlargest(10).reset_index())
        top10.columns = ["Maker", "TOTAL"]
        if not top10.empty:
            others_val = max(int(df_national[sel_months].sum().sum()) - int(top10["TOTAL"].sum()), 0)
            if others_val > 0:
                top10 = pd.concat(
                    [top10, pd.DataFrame([{"Maker": "Others", "TOTAL": others_val}])],
                    ignore_index=True)
            fig = px.pie(top10, values="TOTAL", names="Maker",
                         title="Market Share — Top 10 Makers", hole=0.55,
                         color_discrete_sequence=COLOR_PALETTE)
            fig.update_traces(textposition="inside", textinfo="percent", textfont_size=10)
            fig.update_layout(**plotly_layout(height=440))
            st.plotly_chart(fig, use_container_width=True)

    with c4:
        monthly_nat = pd.DataFrame({
            "Month": sel_months,
            "Registrations": [int(df_national[m].sum()) for m in sel_months],
        })
        if monthly_nat["Registrations"].sum() > 0:
            fig = px.line(monthly_nat, x="Month", y="Registrations",
                          title="Monthly Registration Trend (National)", markers=True)
            fig.update_traces(line=dict(width=3, color=THEME["accent"]),
                              marker=dict(size=8, color=THEME["accent"]))
            fig.update_layout(**plotly_layout(height=440))
            st.plotly_chart(fig, use_container_width=True)


# ─── TAB 2: STATE ANALYSIS ───────────────────────────────────
with tab2:
    avail_states = sorted(df_states["State"].unique())
    if not avail_states:
        st.warning("No state data for current filters.")
    else:
        def_idx = avail_states.index("Gujarat") if "Gujarat" in avail_states else 0
        sel_state = st.selectbox("Select a State / UT", avail_states, index=def_idx)
        dfs = df_states[df_states["State"] == sel_state]

        dfs_total = int(dfs[sel_months].sum().sum())
        top_mk = dfs.groupby("Maker")[sel_months].sum().sum(axis=1).idxmax() \
            if not dfs.empty else "N/A"

        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric("Total Registrations", f"{dfs_total:,}")
        sc2.metric("Active Makers", f"{dfs['Maker'].nunique()}")
        sc3.metric("Top Maker", top_mk)
        sc4.metric("Vehicle Types", f"{dfs['VehicleType'].nunique()}")

        s1, s2 = st.columns(2)
        with s1:
            sm = (dfs.groupby("Maker")[sel_months].sum()
                  .sum(axis=1).nlargest(top_n).reset_index())
            sm.columns = ["Maker", "TOTAL"]
            fig = px.bar(sm, x="TOTAL", y="Maker", orientation="h",
                         title=f"Top {top_n} Makers in {sel_state}")
            fig.update_traces(marker_color=THEME["accent"])
            fig.update_layout(**plotly_layout(
                height=440, showlegend=False,
                yaxis=dict(autorange="reversed", gridcolor=THEME["grid"], color=THEME["text_muted"]),
            ))
            st.plotly_chart(fig, use_container_width=True)

        with s2:
            s_mon = pd.DataFrame({
                "Month": sel_months,
                "Registrations": [int(dfs[m].sum()) for m in sel_months],
            })
            fig = px.area(s_mon, x="Month", y="Registrations",
                          title=f"Monthly Trend — {sel_state}", markers=True)
            fig.update_traces(
                line_color=THEME["accent"], fillcolor=THEME["accent_soft"],
                marker=dict(size=7, color=THEME["accent"]),
            )
            fig.update_layout(**plotly_layout(height=440))
            st.plotly_chart(fig, use_container_width=True)

        allfuel_state = df_master[
            (df_master["FuelType"] == "AllFuel")
            & (df_master["VehicleType"].isin(sel_vtypes))
            & (df_master["Year"] == sel_year)
            & (df_master["State"] == sel_state)
        ]
        pureev_state = df_master[
            (df_master["FuelType"] == "PureEV")
            & (df_master["VehicleType"].isin(sel_vtypes))
            & (df_master["Year"] == sel_year)
            & (df_master["State"] == sel_state)
        ]
        if not allfuel_state.empty and not pureev_state.empty:
            af_tot = int(allfuel_state[sel_months].sum().sum())
            pe_tot = int(pureev_state[sel_months].sum().sum())
            pen = (pe_tot / max(af_tot, 1)) * 100
            st.info(f"**EV Penetration in {sel_state}:** {pen:.1f}%  "
                    f"({pe_tot:,} Electric out of {af_tot:,} total registrations)")


# ─── TAB 3: MAKER DEEP-DIVE ──────────────────────────────────
with tab3:
    all_makers = (df_national.groupby("Maker")[sel_months].sum()
                  .sum(axis=1).sort_values(ascending=False).index.tolist())
    if not all_makers:
        all_makers = (df_states.groupby("Maker")[sel_months].sum()
                      .sum(axis=1).sort_values(ascending=False).index.tolist())

    if not all_makers:
        st.warning("No maker data for current filters.")
    else:
        sel_maker = st.selectbox("Select a Maker", all_makers)
        dfm_nat = df_national[df_national["Maker"] == sel_maker]
        dfm_st = df_states[df_states["Maker"] == sel_maker]

        nat_total = int(dfm_nat[sel_months].sum().sum())
        top_st = dfm_st.groupby("State")[sel_months].sum().sum(axis=1).idxmax() \
            if not dfm_st.empty else "N/A"

        mc1, mc2, mc3, mc4 = st.columns(4)
        mc1.metric("National Total", f"{nat_total:,}")
        mc2.metric("States Present", f"{dfm_st['State'].nunique()}")
        mc3.metric("Vehicle Types", f"{dfm_nat['VehicleType'].nunique()}")
        mc4.metric("Top State", top_st)

        m1, m2 = st.columns(2)
        with m1:
            ms = (dfm_st.groupby("State")[sel_months].sum()
                  .sum(axis=1).sort_values(ascending=False).head(15).reset_index())
            ms.columns = ["State", "TOTAL"]
            if not ms.empty:
                fig = px.bar(ms, x="TOTAL", y="State", orientation="h",
                             title=f"{sel_maker} — State-wise")
                fig.update_traces(marker_color=THEME["accent"])
                fig.update_layout(**plotly_layout(
                    height=440, showlegend=False,
                    yaxis=dict(autorange="reversed", gridcolor=THEME["grid"], color=THEME["text_muted"]),
                ))
                st.plotly_chart(fig, use_container_width=True)

        with m2:
            mm = pd.DataFrame({
                "Month": sel_months,
                "Registrations": [int(dfm_nat[m].sum()) for m in sel_months],
            })
            fig = px.bar(mm, x="Month", y="Registrations",
                         title=f"{sel_maker} — Monthly Trend")
            fig.update_traces(marker_color=THEME["accent2"])
            fig.update_layout(**plotly_layout(height=440, showlegend=False))
            st.plotly_chart(fig, use_container_width=True)

        # MoM growth chart — dual-axis kept (it's information-dense by
        # nature) but now theme-matched and using just 2 colors.
        m_vals = [int(dfm_nat[m].sum()) for m in sel_months]
        growth_data = []
        for i in range(1, len(sel_months)):
            prev, curr = m_vals[i - 1], m_vals[i]
            g = ((curr - prev) / max(prev, 1)) * 100 if prev > 0 else 0
            growth_data.append({
                "Month": sel_months[i],
                "Growth%": round(g, 1),
                "Registrations": curr,
            })
        growth_df = pd.DataFrame(growth_data)
        if not growth_df.empty and growth_df["Registrations"].sum() > 0:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=growth_df["Month"], y=growth_df["Registrations"],
                name="Registrations", marker_color=THEME["accent2"], yaxis="y",
                opacity=0.85,
            ))
            fig.add_trace(go.Scatter(
                x=growth_df["Month"], y=growth_df["Growth%"],
                name="MoM %", mode="lines+markers",
                line=dict(color=THEME["accent"], width=2.5),
                marker=dict(size=8, color=THEME["accent"]), yaxis="y2",
            ))
            fig.update_layout(**plotly_layout(
                title=f"{sel_maker} — Month-over-Month Growth",
                height=380,
                yaxis=dict(title="Registrations", side="left", gridcolor=THEME["grid"], color=THEME["text_muted"]),
                yaxis2=dict(title="MoM %", side="right", overlaying="y", showgrid=False, color=THEME["text_muted"]),
                legend=dict(orientation="h", yanchor="bottom", y=1.05),
            ))
            st.plotly_chart(fig, use_container_width=True)


# ─── TAB 4: MONTHLY TRENDS ───────────────────────────────────
with tab4:
    t10_mkrs = (df_national.groupby("Maker")[sel_months].sum()
                .sum(axis=1).nlargest(10).index.tolist())
    df_t10 = df_national[df_national["Maker"].isin(t10_mkrs)]
    mon_mk = df_t10.groupby("Maker")[sel_months].sum().reset_index()
    mon_mk_m = mon_mk.melt(id_vars="Maker", var_name="Month", value_name="Registrations")

    fig = px.area(mon_mk_m, x="Month", y="Registrations", color="Maker",
                  title="Top 10 Makers — Monthly Trend (Stacked)",
                  color_discrete_sequence=COLOR_PALETTE)
    fig.update_layout(**plotly_layout(
        height=440, legend=dict(orientation="h", yanchor="bottom", y=-0.3),
    ))
    st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        t8_st = (df_states.groupby("State")[sel_months].sum()
                 .sum(axis=1).nlargest(8).index.tolist())
        sm2 = df_states[df_states["State"].isin(t8_st)].groupby("State")[sel_months].sum().reset_index()
        sm2m = sm2.melt(id_vars="State", var_name="Month", value_name="Registrations")
        fig = px.line(sm2m, x="Month", y="Registrations", color="State",
                      title="Top 8 States — Monthly Trends", markers=True,
                      color_discrete_sequence=COLOR_PALETTE)
        fig.update_layout(**plotly_layout(height=440, legend=dict(font_size=9)))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        if len(active_m_list) >= 2:
            lm_t, pm_t = active_m_list[-1], active_m_list[-2]
            st_mt = df_states.groupby("State")[MONTHS].sum()
            smom = pd.DataFrame({
                "State": st_mt.index,
                "Growth%": ((st_mt.get(lm_t, 0).values - st_mt.get(pm_t, 0).values)
                            / np.maximum(st_mt.get(pm_t, 0).values, 1) * 100).round(1),
                "Total": st_mt.get(lm_t, pd.Series([0] * len(st_mt))).values,
            })
            smom = smom[smom["Total"] > 0].sort_values("Growth%").tail(15)
            if not smom.empty:
                # Two flat colors (accent = growth, negative = decline)
                # instead of a continuous RdYlGn gradient — easier to
                # read at a glance, and theme-safe.
                bar_colors = [THEME["accent"] if v >= 0 else THEME["negative"] for v in smom["Growth%"]]
                fig = px.bar(smom, x="Growth%", y="State", orientation="h",
                             title=f"States — Growth ({MONTH_LABELS.get(pm_t,pm_t)} → {MONTH_LABELS.get(lm_t,lm_t)})")
                fig.update_traces(marker_color=bar_colors)
                fig.update_layout(**plotly_layout(height=440, showlegend=False))
                st.plotly_chart(fig, use_container_width=True)


# ─── TAB 5: HEATMAP & RANKINGS ───────────────────────────────
with tab5:
    smp = df_states.groupby("State")[sel_months].sum()
    smp = smp.loc[smp.sum(axis=1).sort_values(ascending=False).index]

    if not smp.empty:
        fig = px.imshow(
            smp.values,
            labels=dict(x="Month", y="State", color="Registrations"),
            x=sel_months, y=smp.index.tolist(),
            color_continuous_scale=SCALE_MAIN,
            title=f"State × Month Heatmap — {fuel_label}",
            aspect="auto",
        )
        # Numbers shown on hover, not stamped on every cell — far
        # less visual noise on a 37-state x 12-month grid.
        fig.update_layout(**plotly_layout(height=max(560, len(smp) * 20)))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    t15m = (df_states.groupby("Maker")[sel_months].sum()
            .sum(axis=1).nlargest(15).index.tolist())
    t10s = (df_states.groupby("State")[sel_months].sum()
            .sum(axis=1).nlargest(10).index.tolist())
    cross = df_states[df_states["Maker"].isin(t15m) & df_states["State"].isin(t10s)]
    if not cross.empty:
        piv = cross.pivot_table(index="Maker", columns="State",
                                values=sel_months[0] if len(sel_months) == 1
                                else sel_months,
                                aggfunc="sum", fill_value=0)
        if isinstance(piv.columns, pd.MultiIndex):
            piv = piv.T.groupby(level=1).sum().T
        piv = piv.loc[piv.sum(axis=1).sort_values(ascending=False).index]
        fig = px.imshow(
            piv.values,
            labels=dict(x="State", y="Maker", color="Registrations"),
            x=piv.columns.tolist(), y=piv.index.tolist(),
            color_continuous_scale=SCALE_MAIN,
            title="Top 15 Makers × Top 10 States",
            aspect="auto",
        )
        fig.update_layout(**plotly_layout(height=560))
        st.plotly_chart(fig, use_container_width=True)


# ─── TAB 6: MARKET INTELLIGENCE ───────────────────────────────
with tab6:
    st.subheader("Market Intelligence")

    st.markdown("#### EV Penetration Rate by State")
    allfuel_all = df_master[
        (df_master["FuelType"] == "AllFuel")
        & (df_master["VehicleType"].isin(sel_vtypes))
        & (df_master["Year"] == sel_year)
        & (df_master["State"].isin(sel_states))
    ]
    pureev_all = df_master[
        (df_master["FuelType"] == "PureEV")
        & (df_master["VehicleType"].isin(sel_vtypes))
        & (df_master["Year"] == sel_year)
        & (df_master["State"].isin(sel_states))
    ]
    if not allfuel_all.empty and not pureev_all.empty:
        af_by_state = allfuel_all.groupby("State")[sel_months].sum().sum(axis=1)
        pe_by_state = pureev_all.groupby("State")[sel_months].sum().sum(axis=1)
        pen_df = pd.DataFrame({
            "State": af_by_state.index,
            "AllFuel": af_by_state.values,
            "PureEV": pe_by_state.reindex(af_by_state.index, fill_value=0).values,
        })
        pen_df["Penetration%"] = (pen_df["PureEV"] / pen_df["AllFuel"].clip(1) * 100).round(1)
        pen_df = pen_df.sort_values("Penetration%", ascending=False)
        fig = px.bar(pen_df, x="Penetration%", y="State", orientation="h",
                     title="EV Penetration Rate by State (Electric ÷ All Fuel × 100)")
        fig.update_traces(marker_color=THEME["accent"])
        fig.update_layout(**plotly_layout(
            height=max(460, len(pen_df) * 20), showlegend=False,
            yaxis=dict(autorange="reversed", gridcolor=THEME["grid"], color=THEME["text_muted"]),
        ))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    st.markdown("#### Maker Rank Tracker (month-by-month)")
    rank_data = []
    for m in sel_months:
        m_totals = df_national.groupby("Maker")[m].sum().sort_values(ascending=False)
        for rank, (maker, val) in enumerate(m_totals.head(10).items(), 1):
            rank_data.append({"Month": m, "Maker": maker, "Rank": rank, "Units": int(val)})
    rank_df = pd.DataFrame(rank_data)
    if not rank_df.empty:
        fig = px.line(rank_df, x="Month", y="Rank", color="Maker",
                      title="Top 10 Maker Rank Over Months (lower = better)",
                      markers=True, color_discrete_sequence=COLOR_PALETTE)
        fig.update_layout(**plotly_layout(
            height=400,
            yaxis=dict(autorange="reversed", dtick=1, gridcolor=THEME["grid"], color=THEME["text_muted"]),
            legend=dict(orientation="h", yanchor="bottom", y=-0.35),
        ))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    st.markdown("#### New Entrants (0 units in first month → units in latest month)")
    if len(sel_months) >= 2:
        first_m, last_m = sel_months[0], sel_months[-1]
        maker_first = df_national.groupby("Maker")[first_m].sum()
        maker_last = df_national.groupby("Maker")[last_m].sum()
        new_entrants = maker_last[(maker_first == 0) & (maker_last > 0)].sort_values(ascending=False)
        if not new_entrants.empty:
            st.dataframe(
                new_entrants.reset_index().rename(columns={last_m: f"Units ({last_m})", "Maker": "Maker"}),
                use_container_width=True, height=220,
            )
        else:
            st.info(f"No new entrants detected between {first_m} and {last_m}.")


# ─── TAB 7: DATA EXPLORER ────────────────────────────────────
with tab7:
    st.subheader("Data Explorer & Download")

    data_scope = st.radio("Scope", ["State-wise", "All India"], horizontal=True)
    df_view = df_states.copy() if data_scope == "State-wise" else df_national.copy()

    search = st.text_input("Search Maker", "")
    if search:
        df_view = df_view[df_view["Maker"].str.contains(search, case=False, na=False)]

    df_view = df_view.copy()
    df_view["TOTAL"] = df_view[sel_months].sum(axis=1)
    display_cols = list(dict.fromkeys(
        (["State"] if data_scope == "State-wise" else [])
        + ["Maker", "VehicleType", "FuelType"]
        + sel_months + ["TOTAL"]
    ))
    display_cols = [c for c in display_cols if c in df_view.columns]

    st.dataframe(
        df_view[display_cols].sort_values("TOTAL", ascending=False).reset_index(drop=True),
        use_container_width=True, height=520,
    )
    st.metric("Rows shown", f"{len(df_view):,}")

    csv = df_view[display_cols].to_csv(index=False).encode("utf-8")
    st.download_button("Download as CSV", data=csv,
                       file_name="vahan_filtered.csv", mime="text/csv")


# ── Footer ───────────────────────────────────────────────────
st.divider()
st.markdown(f"""
<div class="app-footer">
    ⚡ India Vehicle Registration Dashboard &middot;
    Data: <a href="https://vahan.parivahan.gov.in/vahan4dashboard/" target="_blank">Vahan Portal</a>
    &middot; MoRTH, Govt of India &middot;
    Built by <b>Harshal Panchal</b>
</div>
""", unsafe_allow_html=True)
