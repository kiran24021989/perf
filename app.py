from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Historical Analysis Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ========== CUSTOM CSS ==========
st.markdown(
    """
<style>
    .block-container { padding-top: 0.5rem; padding-bottom: 0.5rem; max-width: 100%; }

    /* 1. Eliminate internal vertical gaps between stacked Streamlit blocks */
    div[data-testid="stVerticalBlock"] {
        gap: 0rem !important;
    }

    div[data-testid="stVerticalBlock"] > div {
        padding-top: 2px !important;
        padding-bottom: 2px !important;
        margin-top: 5px !important;
        margin-bottom: 2px !important;
    }

    /* 2. Filter Panel Styling - Zero bottom margin */
    div[data-testid="stHorizontalBlock"]:has(div[data-testid="stSelectbox"]) {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 8px 12px !important;
        border-radius: 8px;
        box-shadow: 0 4px 18px rgba(0, 0, 0, 0.2);
        border: 1px solid #334155;
        margin-bottom: -16px !important; /* Pulls tabs directly against the bottom edge */
        gap: 6px !important;
    }

    /* 3. Filter Labels - Compact & Single-line */
    div[data-testid="stSelectbox"] label p {
        color: #FDFBF7 !important;
        font-family: 'Segoe UI', Roboto, sans-serif !important;
        font-weight: 700 !important;
        font-size: 11px !important;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-bottom: 2px !important;
        line-height: 1.1 !important;
        white-space: nowrap !important;
    }

    /* 4. Filter Selectbox Input Box - Minimum Height */
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 11px !important;
        font-family: 'Segoe UI', Roboto, sans-serif !important;
        min-height: 28px !important;
        height: 28px !important;
        padding-top: 0px !important;
        padding-bottom: 0px !important;
        padding-left: 8px !important;
        line-height: 28px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }

    /* Dropdown Arrow Icon Size Fix */
    div[data-testid="stSelectbox"] div[data-baseweb="select"] svg {
        width: 14px !important;
        height: 14px !important;
    }

    /* 5. Tab Container Layout - Collapsed Spacing */
    div[data-testid="stTabs"] {
        margin-top: 0px !important;
        padding-top: 0px !important;
    }

    div[data-testid="stTabs"] div[role="tablist"] {
        gap: 8px;
        border-bottom: none !important;
        padding-top: 0px !important;
        margin-top: 0px !important;
    }

    /* Inactive Tab Block Styling */
    div[data-testid="stTabs"] button {
        font-family: 'Segoe UI', Roboto, sans-serif !important;
        font-weight: 700 !important;
        font-size: 13px !important;
        padding: 6px 16px !important;
        color: #1e293b !important;
        background-color: #e2e8f0 !important;
        border-radius: 6px !important;
        border: 1px solid #cbd5e1 !important;
        transition: all 0.2s ease-in-out !important;
    }

    /* Hover State */
    div[data-testid="stTabs"] button:hover {
        background-color: #cbd5e1 !important;
        color: #0f172a !important;
    }

    /* Active Selected Tab Block */
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: #ffffff !important;
        background-color: #1d4ed8 !important;
        border-color: #1e40af !important;
        box-shadow: 0 3px 8px rgba(29, 78, 216, 0.3) !important;
    }

    /* Force text color on inner elements */
    div[data-testid="stTabs"] button[aria-selected="true"] p {
        color: #ffffff !important;
    }

    /* Hide the default Streamlit red underline */
    div[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
        display: none !important;
    }
    /* Excel Table Styling */
    .excel-table {
        border-collapse: collapse;
        width: 100%;
        table-layout: auto !important;
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 12px;
        margin-top: 6px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .excel-table th, .excel-table td {
        border: 1px solid #cbd5e1;
        padding: 3px 4px;
        text-align: center;
        white-space: nowrap;
        font-size: 11px;
    }
    .excel-table th { font-weight: 700; font-size: 10px; text-transform: uppercase; }
    .header-tot { background-color: #6b21a8; color: white; }
    .header-fpd { background-color: #15803d; color: white; }
    .header-mhl { background-color: #1d4ed8; color: white; }
    .header-sub { background-color: #f1f5f9; color: #334155; font-size: 10px; }
    .header-left { background-color: #0369a1; color: white; }
    .header-km { background-color: #c2410c; color: white; }
    .header-earn { background-color: #047857; color: white; }
    .pos { background-color: #dcfce7; color: #15803d; font-weight: 600; }
    .neg { background-color: #fee2e2; color: #b91c1c; font-weight: 600; }
   
    .title-bar {
        background: linear-gradient(90deg, #fef08a 0%, #fef9c3 100%);
        color: #854d0e;
        text-align: center;
        font-weight: 700;
        font-size: 14px;
        padding: 2px;
        border-radius: 2px;
        border: 1px solid #fef08a;
        margin-bottom: 2px;
        font-family: 'Segoe UI', Roboto, sans-serif !important;
    }
</style>
""",
    unsafe_allow_html=True,
)
# ========== PATH ==========
PARQUET_FILE = r"D:\dashboard\ser_wise.parquet"
@st.cache_data(ttl=300)
def load_data():
    path = Path(PARQUET_FILE)
    if not path.exists():
        st.error(f"File not found: {PARQUET_FILE}")
        st.stop()
    df = pd.read_parquet(path)
    if not pd.api.types.is_datetime64_any_dtype(df["Date"]):
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    return df
df = load_data()


# ========== CASCADING FILTERS ==========
# HEADING ABOVE FILTERS
st.markdown(
    """
    <div style="
        background-color: #0f172a; 
        color: #ffffff; 
        text-align: center; 
        padding: 8px 12px; 
        border-radius: 6px; 
        font-weight: 700; 
        font-size: 20px; 
        letter-spacing: 0.5px;
        margin-bottom: -6px;
        margin-top: 28px;
        text-transform: uppercase;">
        HISTORICAL ANALYSIS OF RANGAREDDY REGION
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div style='font-size:13px;font-weight:600;color:#334155;margin-bottom:2px;'>Filters</div>", unsafe_allow_html=True)
temp = df.copy()
c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 = st.columns(10)
with c1:
    depot_opts = ["ALL"] + sorted([x for x in temp["DEPOT"].dropna().unique() if str(x).strip()])
    depot = st.selectbox("DEPOT", depot_opts, index=0)
if depot != "ALL":
    temp = temp[temp["DEPOT"] == depot]
with c2:
    mhl_opts = ["ALL"] + sorted([x for x in temp["MHL_NMHL"].dropna().unique() if str(x).strip()])
    mhl = st.selectbox("MHL / NMHL", mhl_opts, index=0)
if mhl != "ALL":
    temp = temp[temp["MHL_NMHL"] == mhl]
with c3:
    rtc_opts = ["ALL"] + sorted([x for x in temp["RTC_HIRE"].dropna().unique() if str(x).strip()])
    rtc = st.selectbox("RTC / HIRE", rtc_opts, index=0)
if rtc != "ALL":
    temp = temp[temp["RTC_HIRE"] == rtc]
with c4:
    pax_opts = [ "TOT", "FPD", "MHL"]
    passengers = st.selectbox("PASSENGERS", pax_opts, index=0)
with c5:
    product_opts = ["ALL"] + sorted([x for x in temp["PRODUCT"].dropna().unique() if str(x).strip()])
    product = st.selectbox("PRODUCT", product_opts, index=0)
if product != "ALL":
    temp = temp[temp["PRODUCT"] == product]
with c6:
    route_opts = ["ALL"] + sorted([x for x in temp["ROUTEE"].dropna().unique() if str(x).strip()])
    route = st.selectbox("ROUTE", route_opts, index=0)
if route != "ALL":
    temp = temp[temp["ROUTEE"] == route]
with c7:
    service_col = next((col for col in ["SERVICE_NO", "SERVICE", "SER_NO", "SERVICE_NUMBER"] if col in temp.columns), None)
    if service_col:
        service_opts = ["ALL"] + sorted([x for x in temp[service_col].dropna().unique() if str(x).strip()])
    else:
        service_opts = ["ALL"]
    service_no = st.selectbox("SERVICE NO", service_opts, index=0)
if service_col and service_no != "ALL":
    temp = temp[temp[service_col] == service_no]
with c8:
    raw_months = [x for x in temp["Month_Name"].dropna().unique() if str(x).strip()]
    if not raw_months:
        raw_months = [x for x in df["Month_Name"].dropna().unique() if str(x).strip()]
    def parse_month_key(m_str):
        try:
            return pd.to_datetime(m_str, format="%b-%y")
        except Exception:
            try:
                return pd.to_datetime(m_str, format="%b-%Y")
            except Exception:
                return pd.to_datetime(m_str, errors="coerce")
    month_opts = sorted(raw_months, key=parse_month_key, reverse=True)
    month = st.selectbox("MONTH", month_opts, index=0)
with c9:
    for_upto = st.selectbox("For / Upto", ["UPTO", "FOR"], index=0)
with c10:
    net_gross = st.selectbox("NET / GROSS", ["Gross", "Net"], index=0)
st.markdown("<div style='margin-top:2px;margin-bottom:2px;'></div>", unsafe_allow_html=True)
# ========== GLOBAL FILTERING LOGIC ==========
base_mask = pd.Series(True, index=df.index)
if depot != "ALL":
    base_mask &= df["DEPOT"] == depot
if mhl != "ALL":
    base_mask &= df["MHL_NMHL"] == mhl
if route != "ALL":
    base_mask &= df["ROUTEE"] == route
if product != "ALL":
    base_mask &= df["PRODUCT"] == product
if "rtc" in dir() and rtc != "ALL":
    base_mask &= df["RTC_HIRE"] == rtc
if service_col and service_no != "ALL":
    base_mask &= df[service_col] == service_no
selected_max_date = df[df["Month_Name"] == month]["Date"].max()
if pd.isna(selected_max_date):
    cy_mask = base_mask & (df["Month_Name"] == month)
    ly_mask = pd.Series(False, index=df.index)
else:
    cy_year = selected_max_date.year
    cy_month_num = selected_max_date.month
    fy_start_year = cy_year if cy_month_num >= 4 else cy_year - 1
    fy_start_date = pd.Timestamp(year=fy_start_year, month=4, day=1)
    if for_upto == "FOR":
        cy_mask = base_mask & (df["Month_Name"] == month)
    else:
        cy_mask = (base_mask & (df["Date"] >= fy_start_date) & (df["Date"] <= selected_max_date))
    ly_max_date = selected_max_date - pd.DateOffset(years=1)
    ly_fy_start_date = fy_start_date - pd.DateOffset(years=1)
    if for_upto == "FOR":
        try:
            mon_name, yr = month.split("-")[0], int(month.split("-")[1])
            ly_month_str = f"{mon_name}-{yr-1}"
            ly_mask = base_mask & (df["Month_Name"] == ly_month_str)
        except Exception:
            ly_mask = pd.Series(False, index=df.index)
    else:
        ly_mask = (base_mask & (df["Date"] >= ly_fy_start_date) & (df["Date"] <= ly_max_date))
cy_data = df[cy_mask].copy()
ly_data = df[ly_mask].copy()
if net_gross == "Gross":
    earn_tot, earn_fpd, earn_mhl = "GE_TOT", "GE_FPD", "GE_MHL"
    prefix = "Gross"
else:
    earn_tot, earn_fpd, earn_mhl = "NE_TOT", "NE_FPD", "NE_MHL"
    prefix = "Net"

# Passenger column based on PASSENGERS filter
if passengers == "FPD":
    pax_col = "PSNGR_FPD"
elif passengers == "MHL":
    pax_col = "PSNGR_MHL"
else:
    pax_col = "PSNGR_TOT"  # BOTH or TOT
# ========== HELPER FUNCTIONS ==========
day_order = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
day_short = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
def weighted_epk(data, earn_col, group_cols=["DEPOT", "ROUTEE", "PRODUCT"]):
    if len(data) == 0:
        return pd.DataFrame()
    g = (
        data.groupby(group_cols + ["Weekday"])
        .agg(earnings=(earn_col, "sum"), kms=("Optd_KMs", "sum"))
        .reset_index()
    )
    g["epk"] = np.where(g["kms"] > 0, g["earnings"] / g["kms"], np.nan)
    pivot = g.pivot_table(
        index=group_cols,
        columns="Weekday",
        values="epk",
        aggfunc="first",
    ).reindex(columns=day_order)
    overall = data.groupby(group_cols).agg(
        earnings=(earn_col, "sum"), kms=("Optd_KMs", "sum")
    )
    pivot["UPTO"] = np.where(overall["kms"] > 0, overall["earnings"] / overall["kms"], np.nan)
    return pivot.round(2)
def fmt(v):
    if pd.isna(v) or v is None:
        return ""
    try:
        if float(v) == 0:
            return ""
    except Exception:
        pass
    return f"{v:,.2f}"
def fmt_pax(v):
    """Passengers whole number, no decimals; blank if 0"""
    if pd.isna(v) or v is None:
        return ""
    try:
        iv = int(round(float(v)))
        if iv == 0:
            return ""
        return f"{iv:,}"
    except Exception:
        return ""
def fmt_growth(v):
    """% Growth number only; blank if 0"""
    if pd.isna(v) or v is None:
        return ""
    try:
        if float(v) == 0:
            return ""
    except Exception:
        pass
    return f"{v:.2f}"
def var_class(v):
    if pd.isna(v) or v is None:
        return ""
    if v > 0:
        return "pos"
    if v < 0:
        return "neg"
    return ""
def build_act_vs_act_table(group_col, data_cy=None, data_ly=None, cy_label="CY", ly_label="LY"):
    if data_cy is None:
        data_cy = cy_data
    if data_ly is None:
        data_ly = ly_data

    if len(data_cy) == 0 and len(data_ly) == 0:
        return None, pd.DataFrame()
    def agg_summary(data):
        if len(data) == 0:
            return pd.DataFrame(columns=[group_col, "kms", "earn_tot", "earn_fpd", "earn_mhl", "pax"])
        g = (
            data.groupby(group_col)
            .agg(
                kms=("Optd_KMs", "sum"),
                earn_tot=(earn_tot, "sum"),
                earn_fpd=(earn_fpd, "sum"),
                earn_mhl=(earn_mhl, "sum"),
                pax=(pax_col, "sum"),
            )
            .reset_index()
        )
        return g
    cy_sum = agg_summary(data_cy)
    ly_sum = agg_summary(data_ly)
    merged = cy_sum.merge(ly_sum, on=group_col, how="outer", suffixes=("_CY", "_LY"))
    for col in ["kms_CY", "kms_LY", "earn_tot_CY", "earn_tot_LY", "earn_fpd_CY", "earn_fpd_LY", "earn_mhl_CY", "earn_mhl_LY", "pax_CY", "pax_LY"]:
        merged[col] = merged.get(col, pd.Series(0, index=merged.index)).fillna(0)
    merged["pax_VAR"] = merged["pax_CY"] - merged["pax_LY"]
    merged["pax_PCT"] = np.where(merged["pax_LY"] != 0, merged["pax_VAR"] * 100 / merged["pax_LY"], np.nan)
    # Passengers as whole numbers (no lakhs, no decimals)
    for col in ["pax_CY", "pax_LY", "pax_VAR"]:
        merged[col] = merged[col].round(0)
    merged["epk_tot_CY"] = np.where(merged["kms_CY"] > 0, merged["earn_tot_CY"] / merged["kms_CY"], np.nan)
    merged["epk_tot_LY"] = np.where(merged["kms_LY"] > 0, merged["earn_tot_LY"] / merged["kms_LY"], np.nan)
    merged["epk_fpd_CY"] = np.where(merged["kms_CY"] > 0, merged["earn_fpd_CY"] / merged["kms_CY"], np.nan)
    merged["epk_fpd_LY"] = np.where(merged["kms_LY"] > 0, merged["earn_fpd_LY"] / merged["kms_LY"], np.nan)
    merged["epk_mhl_CY"] = np.where(merged["kms_CY"] > 0, merged["earn_mhl_CY"] / merged["kms_CY"], np.nan)
    merged["epk_mhl_LY"] = np.where(merged["kms_LY"] > 0, merged["earn_mhl_LY"] / merged["kms_LY"], np.nan)
    merged["kms_VAR"] = merged["kms_CY"] - merged["kms_LY"]
    merged["kms_PCT"] = np.where(merged["kms_LY"] != 0, merged["kms_VAR"] * 100 / merged["kms_LY"], np.nan)
    merged["earn_tot_VAR"] = merged["earn_tot_CY"] - merged["earn_tot_LY"]
    merged["earn_tot_PCT"] = np.where(merged["earn_tot_LY"] != 0, merged["earn_tot_VAR"] * 100 / merged["earn_tot_LY"], np.nan)
    merged["epk_tot_VAR"] = merged["epk_tot_CY"] - merged["epk_tot_LY"]
    merged["epk_tot_PCT"] = np.where((merged["epk_tot_LY"].notna()) & (merged["epk_tot_LY"] != 0), merged["epk_tot_VAR"] * 100 / merged["epk_tot_LY"], np.nan)
    merged["epk_fpd_VAR"] = merged["epk_fpd_CY"] - merged["epk_fpd_LY"]
    merged["epk_fpd_PCT"] = np.where((merged["epk_fpd_LY"].notna()) & (merged["epk_fpd_LY"] != 0), merged["epk_fpd_VAR"] * 100 / merged["epk_fpd_LY"], np.nan)
    merged["epk_mhl_VAR"] = merged["epk_mhl_CY"] - merged["epk_mhl_LY"]
    merged["epk_mhl_PCT"] = np.where((merged["epk_mhl_LY"].notna()) & (merged["epk_mhl_LY"] != 0), merged["epk_mhl_VAR"] * 100 / merged["epk_mhl_LY"], np.nan)
    for col in ["kms_CY", "kms_LY", "kms_VAR", "earn_tot_CY", "earn_tot_LY", "earn_tot_VAR"]:
        merged[col] = merged[col] / 100000
    merged = merged.sort_values(group_col).reset_index(drop=True)
    t_kms_cy = data_cy["Optd_KMs"].sum() / 100000 if len(data_cy) else 0
    t_kms_ly = data_ly["Optd_KMs"].sum() / 100000 if len(data_ly) else 0
    t_earn_cy = data_cy[earn_tot].sum() / 100000 if len(data_cy) else 0
    t_earn_ly = data_ly[earn_tot].sum() / 100000 if len(data_ly) else 0
    t_epk_tot_cy = (data_cy[earn_tot].sum() / data_cy["Optd_KMs"].sum()) if len(data_cy) and data_cy["Optd_KMs"].sum() > 0 else np.nan
    t_epk_tot_ly = (data_ly[earn_tot].sum() / data_ly["Optd_KMs"].sum()) if len(data_ly) and data_ly["Optd_KMs"].sum() > 0 else np.nan
    t_epk_fpd_cy = (data_cy[earn_fpd].sum() / data_cy["Optd_KMs"].sum()) if len(data_cy) and data_cy["Optd_KMs"].sum() > 0 else np.nan
    t_epk_fpd_ly = (data_ly[earn_fpd].sum() / data_ly["Optd_KMs"].sum()) if len(data_ly) and data_ly["Optd_KMs"].sum() > 0 else np.nan
    t_epk_mhl_cy = (data_cy[earn_mhl].sum() / data_cy["Optd_KMs"].sum()) if len(data_cy) and data_cy["Optd_KMs"].sum() > 0 else np.nan
    t_epk_mhl_ly = (data_ly[earn_mhl].sum() / data_ly["Optd_KMs"].sum()) if len(data_ly) and data_ly["Optd_KMs"].sum() > 0 else np.nan
    kms_v = t_kms_cy - t_kms_ly
    earn_v = t_earn_cy - t_earn_ly
    epk_tot_v = (t_epk_tot_cy - t_epk_tot_ly) if pd.notna(t_epk_tot_cy) and pd.notna(t_epk_tot_ly) else np.nan
    epk_fpd_v = (t_epk_fpd_cy - t_epk_fpd_ly) if pd.notna(t_epk_fpd_cy) and pd.notna(t_epk_fpd_ly) else np.nan
    epk_mhl_v = (t_epk_mhl_cy - t_epk_mhl_ly) if pd.notna(t_epk_mhl_cy) and pd.notna(t_epk_mhl_ly) else np.nan
    total = {
        group_col: "TOTAL",
        "kms_CY": t_kms_cy,
        "kms_LY": t_kms_ly,
        "kms_VAR": kms_v,
        "kms_PCT": (kms_v * 100 / t_kms_ly) if t_kms_ly else np.nan,
        "earn_tot_CY": t_earn_cy,
        "earn_tot_LY": t_earn_ly,
        "earn_tot_VAR": earn_v,
        "earn_tot_PCT": (earn_v * 100 / t_earn_ly) if t_earn_ly else np.nan,
        "epk_tot_CY": t_epk_tot_cy,
        "epk_tot_LY": t_epk_tot_ly,
        "epk_tot_VAR": epk_tot_v,
        "epk_tot_PCT": (epk_tot_v * 100 / t_epk_tot_ly) if (pd.notna(t_epk_tot_ly) and t_epk_tot_ly != 0) else np.nan,
        "epk_fpd_CY": t_epk_fpd_cy,
        "epk_fpd_LY": t_epk_fpd_ly,
        "epk_fpd_VAR": epk_fpd_v,
        "epk_fpd_PCT": (epk_fpd_v * 100 / t_epk_fpd_ly) if (pd.notna(t_epk_fpd_ly) and t_epk_fpd_ly != 0) else np.nan,
        "epk_mhl_CY": t_epk_mhl_cy,
        "epk_mhl_LY": t_epk_mhl_ly,
        "epk_mhl_VAR": epk_mhl_v,
        "epk_mhl_PCT": (epk_mhl_v * 100 / t_epk_mhl_ly) if (pd.notna(t_epk_mhl_ly) and t_epk_mhl_ly != 0) else np.nan,
        "pax_CY": round(data_cy[pax_col].sum() if len(data_cy) else 0),
        "pax_LY": round(data_ly[pax_col].sum() if len(data_ly) else 0),
        "pax_VAR": 0,
        "pax_PCT": np.nan,
    }
    total["pax_VAR"] = total["pax_CY"] - total["pax_LY"]
    if total["pax_LY"]:
        total["pax_PCT"] = total["pax_VAR"] * 100 / total["pax_LY"]
    merged = pd.concat([merged, pd.DataFrame([total])], ignore_index=True)
    html = ['<div style="overflow-x:auto;"><table class="excel-table">']
    html.append("<tr>")
    html.append(f'<th class="header-left" rowspan="2">{group_col}</th>')
    html.append('<th class="header-km" colspan="4">KILOMETERS (in lks.)</th>')
    html.append(f'<th class="header-earn" colspan="4">{prefix} EARNINGS (in lks.)</th>')
    html.append(f'<th class="header-tot" colspan="4">{prefix} TOT EPK</th>')
    html.append(f'<th class="header-fpd" colspan="4">{prefix} FPD EPK</th>')
    html.append(f'<th class="header-mhl" colspan="4">{prefix} MHL EPK</th>')
    pax_heading = {"FPD": "FPD PASSENGERS", "MHL": "MHL PASSENGERS"}.get(passengers, "TOTAL PASSENGERS")
    html.append(f'<th class="header-left" colspan="4">{pax_heading}</th>')
    html.append("</tr><tr>")
    for _ in range(6):
        html.append(f'<th class="header-sub">{cy_label}</th>')
        html.append(f'<th class="header-sub">{ly_label}</th>')
        html.append('<th class="header-sub">VAR</th>')
        html.append('<th class="header-sub">% ▲/▼</th>')
    html.append("</tr>")
    for _, row in merged.iterrows():
        is_total = row[group_col] == "TOTAL"
        style = "font-weight:bold; background:#e2efda;" if is_total else ""
        html.append(f'<tr style="{style}">')
        html.append(f'<td>{row[group_col]}</td>')
       
        for metric in ["kms", "earn_tot"]:
            html.append(f'<td>{fmt(row[f"{metric}_CY"])}</td>')
            html.append(f'<td>{fmt(row[f"{metric}_LY"])}</td>')
            html.append(f'<td class="{var_class(row[f"{metric}_VAR"])}">{fmt(row[f"{metric}_VAR"])}</td>')
            html.append(f'<td class="{var_class(row[f"{metric}_PCT"])}">{fmt_growth(row[f"{metric}_PCT"])}</td>')
        for epk_type in ["tot", "fpd", "mhl"]:
            html.append(f'<td>{fmt(row[f"epk_{epk_type}_CY"])}</td>')
            html.append(f'<td>{fmt(row[f"epk_{epk_type}_LY"])}</td>')
            html.append(f'<td class="{var_class(row[f"epk_{epk_type}_VAR"])}">{fmt(row[f"epk_{epk_type}_VAR"])}</td>')
            html.append(f'<td class="{var_class(row[f"epk_{epk_type}_PCT"])}">{fmt_growth(row[f"epk_{epk_type}_PCT"])}</td>')
        # Passengers (already in lakhs, rounded)
        html.append(f'<td>{fmt_pax(row.get("pax_CY", 0))}</td>')
        html.append(f'<td>{fmt_pax(row.get("pax_LY", 0))}</td>')
        html.append(f'<td class="{var_class(row.get("pax_VAR", 0))}">{fmt_pax(row.get("pax_VAR", 0))}</td>')
        html.append(f'<td class="{var_class(row.get("pax_PCT", 0))}">{fmt_growth(row.get("pax_PCT", 0))}</td>')
        html.append("</tr>")
    html.append("</table></div>")
    return "".join(html), merged

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(["Route Day-wise", "ACT VS ACT", "Product wise", "Day wise","trends","cy trends", "Service performance", "Period Comparison"])
# ==================== TAB 1 ====================
with tab1:
    title = f"Route WISE, Day Wise {prefix} EPK ({for_upto}) - {month}"
    if depot != "ALL":
        title += f" of {depot} Depot"
    st.markdown(f'<div class="title-bar">{title}</div>', unsafe_allow_html=True)
    st.caption(f"CY rows: {len(cy_data):,} | LY rows: {len(ly_data):,}")
    if len(cy_data) == 0:
        st.warning("No data for selected filters.")
    else:
        all_parts = []
        for label, earn_col in [("TOT", earn_tot), ("FPD", earn_fpd), ("MHL", earn_mhl)]:
            cy = weighted_epk(cy_data, earn_col)
            ly = weighted_epk(ly_data, earn_col)
            all_idx = cy.index.union(ly.index) if len(ly) > 0 else cy.index
            cy = cy.reindex(all_idx)
            ly = ly.reindex(all_idx) if len(ly) > 0 else pd.DataFrame(np.nan, index=all_idx, columns=day_order + ["UPTO"], dtype=float)
            block = pd.DataFrame(index=all_idx)
            for d in day_order:
                block[f"{label}_{d}"] = pd.to_numeric(cy[d] if d in cy.columns else np.nan, errors="coerce")
            block[f"{label}_CY"] = pd.to_numeric(cy["UPTO"] if "UPTO" in cy.columns else np.nan, errors="coerce")
            block[f"{label}_LY"] = pd.to_numeric(ly["UPTO"] if "UPTO" in ly.columns else np.nan, errors="coerce")
            block[f"{label}_Var"] = (block[f"{label}_CY"] - block[f"{label}_LY"]).astype(float).round(2)
            all_parts.append(block)
        result = pd.concat(all_parts, axis=1).reset_index()
        result = result.rename(columns={"ROUTEE": "ROUTE"})
        result = result.sort_values(["DEPOT", "ROUTE", "PRODUCT"]).reset_index(drop=True)
        html = ['<div style="overflow-x:auto;"><table class="excel-table">']
        html.append("<tr>")
        html.append('<th class="header-left" rowspan="2">DEPOT</th>')
        html.append('<th class="header-left" rowspan="2">ROUTE</th>')
        html.append('<th class="header-left" rowspan="2">PRODUCT</th>')
        html.append(f'<th class="header-tot" colspan="10">{prefix} TOT. E.P.K (in Ps/kms.)</th>')
        html.append(f'<th class="header-fpd" colspan="10">{prefix} FPD. E.P.K (in Ps/kms.)</th>')
        html.append(f'<th class="header-mhl" colspan="10">{prefix} MHL. E.P.K (in Ps/kms.)</th>')
        html.append("</tr><tr>")
        for _ in range(3):
            for d in day_short:
                html.append(f'<th class="header-sub">{d}</th>')
            html.append(f'<th class="header-sub">{for_upto} CY</th>')
            html.append(f'<th class="header-sub">{for_upto} LY</th>')
            html.append('<th class="header-sub">Var</th>')
        html.append("</tr>")
        for _, row in result.iterrows():
            html.append("<tr>")
            html.append(f'<td>{row["DEPOT"]}</td>')
            html.append(f'<td>{row["ROUTE"]}</td>')
            html.append(f'<td>{row["PRODUCT"]}</td>')
            for label in ["TOT", "FPD", "MHL"]:
                # Peak / Slack among Mon-Sun for this metric on this row
                day_vals = {}
                for d in day_order:
                    v = row.get(f"{label}_{d}")
                    try:
                        fv = float(v) if pd.notna(v) else None
                    except Exception:
                        fv = None
                    if fv is not None and fv != 0:
                        day_vals[d] = fv
                peak_d = max(day_vals, key=day_vals.get) if day_vals else None
                slack_d = min(day_vals, key=day_vals.get) if day_vals else None
                for d in day_order:
                    val = row.get(f"{label}_{d}")
                    style = ""
                    if peak_d and d == peak_d:
                        style = 'background-color:#c6efce; color:#006100; font-weight:600;'
                    elif slack_d and d == slack_d:
                        style = 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
                    html.append(f'<td style="{style}">{fmt(val)}</td>')
                html.append(f'<td>{fmt(row.get(f"{label}_CY"))}</td>')
                html.append(f'<td>{fmt(row.get(f"{label}_LY"))}</td>')
                var_val = row.get(f"{label}_Var")
                html.append(f'<td class="{var_class(var_val)}">{fmt(var_val)}</td>')
            html.append("</tr>")
        html.append("</table></div>")
        st.markdown("".join(html), unsafe_allow_html=True)
        st.markdown("""
<span style="background:#c6efce; padding:2px 8px;">Peak day</span> = Highest EPK among Mon–Sun &nbsp;
<span style="background:#ffc7ce; padding:2px 8px;">Slack day</span> = Lowest EPK among Mon–Sun
&nbsp;(per row, for each of TOT / FPD / MHL)
""", unsafe_allow_html=True)
        st.download_button("Download CSV", result.to_csv(index=False).encode("utf-8"), f"Route_Daywise_{month}_{depot}.csv", "text/csv", key="dl1")
# ==================== TAB 2 ====================
with tab2:
    title2 = f"Actual vs Actual Performance ({net_gross}) {for_upto} the - Month of : {month}"
    st.markdown(f'<div class="title-bar">{title2}</div>', unsafe_allow_html=True)
    html_str, merged_df = build_act_vs_act_table(group_col="DEPOT")
    if html_str is None:
        st.warning("No data for selected filters.")
    else:
        st.markdown(html_str, unsafe_allow_html=True)
        st.download_button("Download CSV", merged_df.to_csv(index=False).encode("utf-8"), f"ACT_VS_ACT_{month}.csv", "text/csv", key="dl2")
# ==================== TAB 3 ====================
with tab3:
    title3 = f"Product Wise Performance ({net_gross}) ({for_upto}) - Month: {month}"
    st.markdown(f'<div class="title-bar">{title3}</div>', unsafe_allow_html=True)
    html_str, merged_df = build_act_vs_act_table(group_col="PRODUCT")
    if html_str is None:
        st.warning("No data for selected filters.")
    else:
        st.markdown(html_str, unsafe_allow_html=True)
        st.download_button("Download CSV", merged_df.to_csv(index=False).encode("utf-8"), f"Product_Summary_ACT_VS_ACT_{month}.csv", "text/csv", key="dl3")
# ==================== TAB 4 ====================
with tab4:
    prep = "upto" if for_upto == "UPTO" else "for"

    # Construct the dynamic title
    title5 = f"'AVERAGE' Weekday Wise Actual vs Actuals"
    if depot != "ALL":
        title5 += f" of {depot} Depot"
    else:
        title5 += " of ALL Depots"

    title5 += f" {prep} the month of {month}"

    st.markdown(f'<div class="title-bar">{title5}</div>', unsafe_allow_html=True)
    if len(cy_data) == 0 and for_upto != "UPTO":
        st.warning("No data for selected filters.")
    else:
        from calendar import monthcalendar
        day_order = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        mon_map = {
            "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
            "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
        }
        def parse_month(mstr):
            try:
                parts = mstr.replace(" ", "").split("-")
                return mon_map.get(parts[0][:3], 1), int(parts[1])
            except:
                return 1, 2026
        def count_weekdays_in_month(month_str):
            try:
                m, y = parse_month(month_str)
                cal = monthcalendar(y, m)
                counts = {d: 0 for d in day_order}
                for week in cal:
                    for idx, day in enumerate(week):
                        if day != 0:
                            counts[day_order[idx]] += 1
                return counts
            except:
                return {d: 4 for d in day_order}
        def count_weekdays_upto(end_month_str, start_month=4):
            try:
                em, ey = parse_month(end_month_str)
                start_y = ey if em >= start_month else ey - 1
                counts = {d: 0 for d in day_order}
                y, m = start_y, start_month
                while True:
                    cal = monthcalendar(y, m)
                    for week in cal:
                        for idx, day in enumerate(week):
                            if day != 0:
                                counts[day_order[idx]] += 1
                    if y == ey and m == em:
                        break
                    m += 1
                    if m > 12:
                        m = 1
                        y += 1
                    if y > ey + 1:
                        break
                return counts
            except:
                return {d: 4 for d in day_order}
        def get_months_upto(end_month_str, start_month=4):
            try:
                em, ey = parse_month(end_month_str)
                start_y = ey if em >= start_month else ey - 1
                result = []
                y, m = start_y, start_month
                abbr = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                        7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
                while True:
                    result.append(f"{abbr[m]}-{y}")
                    if y == ey and m == em:
                        break
                    m += 1
                    if m > 12:
                        m = 1
                        y += 1
                    if y > ey + 1:
                        break
                return result
            except:
                return [end_month_str]
        try:
            parts = month.split("-")
            _ly_month = f"{parts[0]}-{int(parts[1]) - 1}"
        except:
            _ly_month = None
        if for_upto == "UPTO":
            cy_months = get_months_upto(month)
            ly_months = get_months_upto(_ly_month) if _ly_month else []
            cy_use = df[base_mask & df["Month_Name"].isin(cy_months)].copy()
            ly_use = df[base_mask & df["Month_Name"].isin(ly_months)].copy() if ly_months else pd.DataFrame()
            cy_counts = count_weekdays_upto(month)
            ly_counts = count_weekdays_upto(_ly_month) if _ly_month else {d: 4 for d in day_order}
        else:
            cy_use = cy_data
            ly_use = ly_data
            cy_counts = count_weekdays_in_month(month)
            ly_counts = count_weekdays_in_month(_ly_month) if _ly_month else {d: 4 for d in day_order}
        def weekday_excel_style(data, earn_tot_col, earn_fpd_col, earn_mhl_col, wd_counts):
            if len(data) == 0:
                return pd.DataFrame()
            g = data.groupby("Weekday").agg(
                Total_KMs=("Optd_KMs", "sum"),
                Total_Earn_TOT=(earn_tot_col, "sum"),
                Total_Earn_FPD=(earn_fpd_col, "sum"),
                Total_Earn_MHL=(earn_mhl_col, "sum"),
                Total_Pax=(pax_col, "sum"),
            ).reset_index()
            g["Days_In_Month"] = g["Weekday"].map(wd_counts).fillna(4)
            g["Avg_KMs"] = (g["Total_KMs"] / g["Days_In_Month"]) / 100000
            g["Avg_Earn_TOT"] = (g["Total_Earn_TOT"] / g["Days_In_Month"]) / 100000
            g["Avg_Pax"] = (g["Total_Pax"] / g["Days_In_Month"]).round(0)
            g["EPK_TOT"] = np.where(g["Total_KMs"] > 0, g["Total_Earn_TOT"] / g["Total_KMs"], np.nan)
            g["EPK_FPD"] = np.where(g["Total_KMs"] > 0, g["Total_Earn_FPD"] / g["Total_KMs"], np.nan)
            g["EPK_MHL"] = np.where(g["Total_KMs"] > 0, g["Total_Earn_MHL"] / g["Total_KMs"], np.nan)
            g["Weekday"] = pd.Categorical(g["Weekday"], categories=day_order, ordered=True)
            return g.sort_values("Weekday").reset_index(drop=True)
        cy_wd = weekday_excel_style(cy_use, earn_tot, earn_fpd, earn_mhl, cy_counts)
        ly_wd = weekday_excel_style(ly_use, earn_tot, earn_fpd, earn_mhl, ly_counts)
        if len(cy_wd) == 0:
            st.warning("No data for selected filters.")
        else:
            merged = cy_wd.merge(ly_wd, on="Weekday", how="outer", suffixes=("_CY", "_LY"))
            merged = merged.sort_values("Weekday").reset_index(drop=True)
            for metric in ["Avg_KMs", "Avg_Earn_TOT", "EPK_TOT", "EPK_FPD", "EPK_MHL", "Avg_Pax"]:
                merged[f"{metric}_VAR"] = merged.get(f"{metric}_CY", 0).fillna(0) - merged.get(f"{metric}_LY", 0).fillna(0)
            # TOTAL = average of Mon-Sun
            total_row = {"Weekday": "TOTAL"}
            for metric in ["Avg_KMs", "Avg_Earn_TOT", "EPK_TOT", "EPK_FPD", "EPK_MHL", "Avg_Pax"]:
                total_row[f"{metric}_CY"] = merged[f"{metric}_CY"].mean()
                total_row[f"{metric}_LY"] = merged[f"{metric}_LY"].mean()
                total_row[f"{metric}_VAR"] = (total_row[f"{metric}_CY"] or 0) - (total_row[f"{metric}_LY"] or 0)
            if len(cy_use) > 0 and cy_use["Optd_KMs"].sum() > 0:
                total_row["EPK_TOT_CY"] = cy_use[earn_tot].sum() / cy_use["Optd_KMs"].sum()
                total_row["EPK_FPD_CY"] = cy_use[earn_fpd].sum() / cy_use["Optd_KMs"].sum()
                total_row["EPK_MHL_CY"] = cy_use[earn_mhl].sum() / cy_use["Optd_KMs"].sum()
            if len(ly_use) > 0 and ly_use["Optd_KMs"].sum() > 0:
                total_row["EPK_TOT_LY"] = ly_use[earn_tot].sum() / ly_use["Optd_KMs"].sum()
                total_row["EPK_FPD_LY"] = ly_use[earn_fpd].sum() / ly_use["Optd_KMs"].sum()
                total_row["EPK_MHL_LY"] = ly_use[earn_mhl].sum() / ly_use["Optd_KMs"].sum()
            for m in ["EPK_TOT", "EPK_FPD", "EPK_MHL"]:
                total_row[f"{m}_VAR"] = (total_row.get(f"{m}_CY") or 0) - (total_row.get(f"{m}_LY") or 0)
            merged = pd.concat([merged, pd.DataFrame([total_row])], ignore_index=True)
            # Colour scale helper for CY values (High=Green → Low=Red)
            def cy_color(val, series):
                if pd.isna(val):
                    return ""
                s = series.dropna()
                if len(s) < 2:
                    return ""
                mn, mx = s.min(), s.max()
                if mx == mn:
                    return "background-color:#ffff99;"
                # 0 = low (red), 1 = high (green)
                ratio = (float(val) - mn) / (mx - mn)
                # interpolate red -> yellow -> green
                if ratio >= 0.6:
                    return "background-color:#c6efce; color:#006100;" # high green
                elif ratio >= 0.3:
                    return "background-color:#ffeb9c; color:#9c5700;" # mid yellow
                else:
                    return "background-color:#ffc7ce; color:#9c0006;" # low red
            # Precompute series for each CY metric (exclude TOTAL)
            non_tot = merged[merged["Weekday"] != "TOTAL"]
            cy_series = {
                "Avg_KMs": non_tot["Avg_KMs_CY"] if "Avg_KMs_CY" in non_tot else pd.Series(),
                "Avg_Earn_TOT": non_tot["Avg_Earn_TOT_CY"] if "Avg_Earn_TOT_CY" in non_tot else pd.Series(),
                "EPK_TOT": non_tot["EPK_TOT_CY"] if "EPK_TOT_CY" in non_tot else pd.Series(),
                "EPK_FPD": non_tot["EPK_FPD_CY"] if "EPK_FPD_CY" in non_tot else pd.Series(),
                "EPK_MHL": non_tot["EPK_MHL_CY"] if "EPK_MHL_CY" in non_tot else pd.Series(),
                "Avg_Pax": non_tot["Avg_Pax_CY"] if "Avg_Pax_CY" in non_tot else pd.Series(),
            }
            # Charts
            import plotly.graph_objects as go
            from plotly.subplots import make_subplots
            chart_data = merged[merged["Weekday"] != "TOTAL"].copy()
            fig = make_subplots(rows=1, cols=5,
                subplot_titles=("KILOMETERS (in lks.)", f"{prefix} TOT. E.P.K",
                                f"{prefix} FPD. E.P.K", f"{prefix} MHL. E.P.K", {"FPD": "FPD PASSENGERS", "MHL": "MHL PASSENGERS"}.get(passengers, "TOTAL PASSENGERS")),
                horizontal_spacing=0.04)
            for metric, col in [("Avg_KMs",1),("EPK_TOT",2),("EPK_FPD",3),("EPK_MHL",4),("Avg_Pax",5)]:
                fig.add_trace(go.Bar(name="CY", x=chart_data["Weekday"], y=chart_data[f"{metric}_CY"],
                    marker_color="#4472C4", showlegend=(col==1),
                    text=[f"{v:.2f}" if pd.notna(v) else "" for v in chart_data[f"{metric}_CY"]],
                    textposition="outside", textfont_size=9), row=1, col=col)
                fig.add_trace(go.Bar(name="LY", x=chart_data["Weekday"], y=chart_data[f"{metric}_LY"],
                    marker_color="#ED7D31", showlegend=(col==1),
                    text=[f"{v:.2f}" if pd.notna(v) else "" for v in chart_data[f"{metric}_LY"]],
                    textposition="outside", textfont_size=9), row=1, col=col)
            fig.update_layout(barmode="group", height=340, margin=dict(l=20,r=20,t=80,b=20),
                legend=dict(orientation="h", y=1.15, font=dict(color="white")),
                plot_bgcolor="white", paper_bgcolor="#1e3a5f",
                font=dict(color="white", size=10),
                title_font=dict(color="white"))
            fig.update_xaxes(tickfont=dict(size=9, color="white"))
            fig.update_yaxes(tickfont=dict(size=9, color="white"), gridcolor="#555")
            for ann in fig["layout"]["annotations"]:
                ann["font"] = dict(size=11, color="#FFD700", family="Segoe UI")
                try:
                    ann["y"] = float(ann["y"]) + 0.08 if ann["y"] is not None else 1.08
                except Exception:
                    pass
            st.plotly_chart(fig, use_container_width=True)
            # Table - only CY, LY, VAR (no % Growth)
            html = ['<div style="overflow-x:auto;"><table class="excel-table">']
            html.append('<tr>')
            html.append('<th rowspan="2" style="background:#c000c0; color:white;">DAY</th>')
            html.append('<th colspan="3" style="background:#c000c0; color:white;">KILOMETERS (in lks.)</th>')
            html.append(f'<th colspan="3" style="background:#c000c0; color:white;">{prefix} EARNINGS (Rs.in lks.)</th>')
            html.append(f'<th colspan="3" style="background:#c000c0; color:white;">{prefix} TOT. E.P.K</th>')
            html.append(f'<th colspan="3" style="background:#c000c0; color:white;">{prefix} FPD. E.P.K</th>')
            html.append(f'<th colspan="3" style="background:#c000c0; color:white;">{prefix} MHL. E.P.K</th>')
            pax_heading = {"FPD": "FPD PASSENGERS", "MHL": "MHL PASSENGERS"}.get(passengers, "TOTAL PASSENGERS")
            html.append(f'<th colspan="3" style="background:#c000c0; color:white;">{pax_heading}</th>')
            html.append('</tr><tr>')
            for _ in range(6):
                for h in ["CY", "LY", "VAR"]:
                    html.append(f'<th style="background:#ff66ff; color:black; font-size:10px;">{h}</th>')
            html.append('</tr>')
            for _, row in merged.iterrows():
                is_total = str(row["Weekday"]) == "TOTAL"
                style = 'font-weight:bold; background:#e2efda;' if is_total else ''
                html.append(f'<tr style="{style}">')
                html.append(f'<td><b>{row["Weekday"]}</b></td>')
                for metric in ["Avg_KMs", "Avg_Earn_TOT", "EPK_TOT", "EPK_FPD", "EPK_MHL", "Avg_Pax"]:
                    cy_val = row.get(f"{metric}_CY")
                    ly_val = row.get(f"{metric}_LY")
                    var_val = row.get(f"{metric}_VAR")
                    show = fmt_pax if metric == "Avg_Pax" else fmt
                    if is_total:
                        html.append(f'<td>{show(cy_val)}</td>')
                    else:
                        html.append(f'<td style="{cy_color(cy_val, cy_series[metric])}">{show(cy_val)}</td>')
                    html.append(f'<td>{show(ly_val)}</td>')
                    html.append(f'<td>{show(var_val)}</td>')
                html.append('</tr>')
            html.append('</table></div>')
            st.markdown("".join(html), unsafe_allow_html=True)
            st.markdown("""
**designed by:**
&nbsp;&nbsp;|&nbsp;&nbsp; kiran kumar
""", unsafe_allow_html=True)
            st.caption(f"Mode: {for_upto} | Weekday counts CY: {cy_counts}")
            csv5 = merged.to_csv(index=False).encode("utf-8")
            st.download_button("Download CSV", csv5, f"WeekDay_{month}_{depot}.csv", "text/csv", key="dl5")

# ==================== TAB 5 ====================
with tab5:
    prep_text = "upto" if for_upto == "UPTO" else "for"
    loc_text = f"{depot} Depot" if depot != "ALL" else "ALL Depots"
    title_tab5 = f"Month Wise Actual vs Actuals of {loc_text} {prep_text} the month of {month}"
    st.markdown(f'<div class="title-bar">{title_tab5}</div>', unsafe_allow_html=True)

    # 1. Base dataset applying selected filters
    m_base = df[base_mask].copy() if ('base_mask' in locals() and base_mask.any()) else df.copy()

    if len(m_base) == 0:
        st.warning("No data found for selected filters.")
    else:
        date_col = next((col for col in ["Date", "Month_Name", "Month", "MONTH"] if col in m_base.columns), None)

        if date_col is None:
            st.error("Could not find a valid Date or Month column in dataset.")
        else:
            # Parse Dates
            m_base["_parsed_date"] = pd.to_datetime(m_base[date_col], errors="coerce")
            m_base = m_base.dropna(subset=["_parsed_date"]).copy()

            m_base["_month_short"] = m_base["_parsed_date"].dt.strftime("%b")

            # Calculate Financial Year starting year (April to March)
            m_base["_fy"] = np.where(
                m_base["_parsed_date"].dt.month >= 4,
                m_base["_parsed_date"].dt.year,
                m_base["_parsed_date"].dt.year - 1
            )

            # Available FYs sorted descending
            available_fys = sorted(m_base["_fy"].unique(), reverse=True)
            fy_options = [f"{fy}-{str(fy + 1)[-2:]}" for fy in available_fys]

            # SHORT FY SELECTBOX: Wrapped in small column layout
            col_fy, col_spacer = st.columns([1, 4])
            with col_fy:
                if available_fys:
                    selected_fy_str = st.selectbox(
                        "Select Financial Year",
                        options=fy_options,
                        index=0,
                        key=f"fy_select_{month}_{depot}"
                    )
                    selected_fy = int(selected_fy_str.split("-")[0])
                else:
                    selected_fy = 2026
                    selected_fy_str = "2026-27"

            # Filter dataset by Selected FY (CY) and Previous FY (LY)
            cy_data = m_base[m_base["_fy"] == selected_fy].copy()
            ly_data = m_base[m_base["_fy"] == (selected_fy - 1)].copy()

            cy_yr = str(selected_fy)[-2:]
            ly_yr = str(selected_fy - 1)[-2:]
            next_yr = str(selected_fy + 1)[-2:]

            fy_months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
            month_label_map = {
                'Apr': f'Apr-{cy_yr}', 'May': f'May-{cy_yr}', 'Jun': f'Jun-{cy_yr}', 'Jul': f'Jul-{cy_yr}', 
                'Aug': f'Aug-{cy_yr}', 'Sep': f'Sep-{cy_yr}', 'Oct': f'Oct-{cy_yr}', 'Nov': f'Nov-{cy_yr}', 
                'Dec': f'Dec-{cy_yr}', 'Jan': f'Jan-{next_yr}', 'Feb': f'Feb-{next_yr}', 'Mar': f'Mar-{next_yr}'
            }

            # Aggregation for CY
            if not cy_data.empty:
                cy_agg = cy_data.groupby("_month_short").agg(
                    kms=("Optd_KMs", "sum"),
                    earn_tot=(earn_tot, "sum"),
                    earn_fpd=(earn_fpd, "sum"),
                    earn_mhl=(earn_mhl, "sum"),
                    pax=(pax_col, "sum")
                ).reindex(fy_months)
            else:
                cy_agg = pd.DataFrame(index=fy_months, columns=["kms", "earn_tot", "earn_fpd", "earn_mhl", "pax"])

            # Aggregation for LY
            if not ly_data.empty:
                ly_agg = ly_data.groupby("_month_short").agg(
                    kms=("Optd_KMs", "sum"),
                    earn_tot=(earn_tot, "sum"),
                    earn_fpd=(earn_fpd, "sum"),
                    earn_mhl=(earn_mhl, "sum"),
                    pax=(pax_col, "sum")
                ).reindex(fy_months)
            else:
                ly_agg = pd.DataFrame(index=fy_months, columns=["kms", "earn_tot", "earn_fpd", "earn_mhl", "pax"])

            # Dynamic slicing
            valid_cy_months = cy_agg[cy_agg["kms"].notna() & (cy_agg["kms"] > 0)].index.tolist()
            if valid_cy_months:
                last_month_idx = max(fy_months.index(m) for m in valid_cy_months)
                active_fy_months = fy_months[:last_month_idx + 1]
            else:
                active_fy_months = fy_months

            # Chart Dataframe
            chart_labels = [month_label_map[m] for m in active_fy_months]
            m_chart = pd.DataFrame(index=chart_labels)
            m_chart["Month_Group"] = chart_labels
            
            cy_kms_val = cy_agg.loc[active_fy_months, "kms"].fillna(0).values
            m_chart["Kms_Lakhs"] = (cy_kms_val / 100000).round(2)
            m_chart["Earn_Lakhs"] = (cy_agg.loc[active_fy_months, "earn_tot"].fillna(0).values / 100000).round(2)
            m_chart["Pax_Raw"] = cy_agg.loc[active_fy_months, "pax"].fillna(0).astype(int).values
            
            m_chart["Tot_EPK"] = np.where(cy_kms_val > 0, (cy_agg.loc[active_fy_months, "earn_tot"].fillna(0).values / cy_kms_val).round(2), 0.0)
            m_chart["FPD_EPK"] = np.where(cy_kms_val > 0, (cy_agg.loc[active_fy_months, "earn_fpd"].fillna(0).values / cy_kms_val).round(2), 0.0)
            m_chart["MHL_EPK"] = np.where(cy_kms_val > 0, (cy_agg.loc[active_fy_months, "earn_mhl"].fillna(0).values / cy_kms_val).round(2), 0.0)

            # TOP BAR CHARTS WITH EXPANDED RANGE AND WRAPPED TITLES
            def create_card_chart(df_chart, y_col, chart_title, bar_color="#1d4ed8", is_pax=False):
                max_val = df_chart[y_col].max() if not df_chart[y_col].empty else 100
                y_upper = (max_val * 1.25) if max_val > 0 else 10

                fig = px.bar(
                    df_chart,
                    x="Month_Group",
                    y=y_col,
                    title=f"<b>{chart_title}</b>",
                    text=y_col,
                    color_discrete_sequence=[bar_color]
                )
                
                text_fmt = '%{text:,d}' if is_pax else '%{text:.2f}'
                fig.update_traces(
                    texttemplate=text_fmt,
                    textposition='outside',
                    textfont_size=9,
                    textangle=0  # Upright text prevents horizontal clipping
                )
                
                fig.update_layout(
                    height=260,
                    margin=dict(l=2, r=2, t=55, b=5),  # Increased top margin for headings
                    xaxis_title="",
                    yaxis_title="",
                    xaxis=dict(tickangle=-90, showgrid=False, tickfont=dict(size=10)),
                    yaxis=dict(showgrid=True, showticklabels=False, range=[0, y_upper]),
                    template="plotly_white",
                    title=dict(
                        font=dict(size=10),
                        x=0.5,
                        xanchor='center',
                        yanchor='top'
                    )
                )
                return fig

            ch1, ch2, ch3, ch4, ch5 = st.columns(5)
            with ch1:
                st.plotly_chart(create_card_chart(m_chart, "Kms_Lakhs", "KILOMETERS<br>(in lks.)", "#0284c7"), use_container_width=True, key=f"ch_t5_1_{selected_fy}")
            with ch2:
                st.plotly_chart(create_card_chart(m_chart, "Tot_EPK", f"{prefix} TOT. E.P.K<br>(in Ps/kms.)", "#2563eb"), use_container_width=True, key=f"ch_t5_2_{selected_fy}")
            with ch3:
                st.plotly_chart(create_card_chart(m_chart, "FPD_EPK", f"{prefix} FPD. E.P.K<br>(in Ps/kms.)", "#16a34a"), use_container_width=True, key=f"ch_t5_3_{selected_fy}")
            with ch4:
                st.plotly_chart(create_card_chart(m_chart, "MHL_EPK", f"{prefix} MHL. E.P.K<br>(in Ps/kms.)", "#9333ea"), use_container_width=True, key=f"ch_t5_4_{selected_fy}")
            with ch5:
                pax_title = {"FPD": "FPD PASSENGERS", "MHL": "MHL PASSENGERS"}.get(passengers, "TOTAL PASSENGERS")
                st.plotly_chart(create_card_chart(m_chart, "Pax_Raw", f"{pax_title}<br>(IN NUMBERS)", "#db2777", is_pax=True), use_container_width=True, key=f"ch_t5_5_{selected_fy}")

            # Table Output
            html = ['<div style="overflow-x:auto;"><table class="excel-table">']
            html.append("<tr>")
            html.append('<th rowspan="2" class="header-left">DAY / MONTH</th>')
            html.append('<th colspan="4" class="header-sub">KILOMETERS (IN LKS.)</th>')
            html.append(f'<th colspan="4" class="header-sub">{prefix.upper()} EARNINGS (RS.IN LKS.)</th>')
            html.append(f'<th colspan="4" class="header-sub">{prefix.upper()} TOT. E.P.K (IN PS/KMS.)</th>')
            html.append(f'<th colspan="4" class="header-sub">{prefix.upper()} FPD. E.P.K (IN PS/KMS.)</th>')
            html.append(f'<th colspan="4" class="header-sub">{prefix.upper()} MHL. E.P.K (IN PS/KMS.)</th>')
            html.append(f'<th colspan="4" class="header-sub">{pax_title.upper()} </th>')
            html.append("</tr><tr>")
            for _ in range(6):
                html.append('<th class="header-sub">CY</th><th class="header-sub">LY</th><th class="header-sub">VAR</th><th class="header-sub">% ▲/▼</th>')
            html.append("</tr>")

            export_rows = []

            # Variables for UPTO totals
            tot_c_kms = tot_l_kms = 0.0
            tot_c_earn = tot_l_earn = 0.0
            tot_c_fpd_earn = tot_l_fpd_earn = 0.0
            tot_c_mhl_earn = tot_l_mhl_earn = 0.0
            tot_c_pax = tot_l_pax = 0.0

            for m_short in active_fy_months:
                m_label = month_label_map[m_short]

                c_kms_raw = cy_agg.loc[m_short, "kms"] if m_short in cy_agg.index else 0
                l_kms_raw = ly_agg.loc[m_short, "kms"] if m_short in ly_agg.index else 0

                c_earn_raw = cy_agg.loc[m_short, "earn_tot"] if m_short in cy_agg.index else 0
                l_earn_raw = ly_agg.loc[m_short, "earn_tot"] if m_short in ly_agg.index else 0

                c_fpd_raw = cy_agg.loc[m_short, "earn_fpd"] if m_short in cy_agg.index else 0
                l_fpd_raw = ly_agg.loc[m_short, "earn_fpd"] if m_short in ly_agg.index else 0

                c_mhl_raw = cy_agg.loc[m_short, "earn_mhl"] if m_short in cy_agg.index else 0
                l_mhl_raw = ly_agg.loc[m_short, "earn_mhl"] if m_short in ly_agg.index else 0

                c_pax_raw = cy_agg.loc[m_short, "pax"] if m_short in cy_agg.index else 0
                l_pax_raw = ly_agg.loc[m_short, "pax"] if m_short in ly_agg.index else 0

                # Accumulate UPTO Totals
                tot_c_kms += c_kms_raw
                tot_l_kms += l_kms_raw
                tot_c_earn += c_earn_raw
                tot_l_earn += l_earn_raw
                tot_c_fpd_earn += c_fpd_raw
                tot_l_fpd_earn += l_fpd_raw
                tot_c_mhl_earn += c_mhl_raw
                tot_l_mhl_earn += l_mhl_raw
                tot_c_pax += c_pax_raw
                tot_l_pax += l_pax_raw

                # Display values
                c_kms = c_kms_raw / 100000
                l_kms = l_kms_raw / 100000

                c_earn = c_earn_raw / 100000
                l_earn = l_earn_raw / 100000

                c_tot_epk = (c_earn_raw / c_kms_raw) if c_kms_raw > 0 else np.nan
                l_tot_epk = (l_earn_raw / l_kms_raw) if l_kms_raw > 0 else np.nan

                c_fpd_epk = (c_fpd_raw / c_kms_raw) if c_kms_raw > 0 else np.nan
                l_fpd_epk = (l_fpd_raw / l_kms_raw) if l_kms_raw > 0 else np.nan

                c_mhl_epk = (c_mhl_raw / c_kms_raw) if c_kms_raw > 0 else np.nan
                l_mhl_epk = (l_mhl_raw / l_kms_raw) if l_kms_raw > 0 else np.nan

                c_pax = c_pax_raw
                l_pax = l_pax_raw

                metrics = [
                    (c_kms, l_kms, False),
                    (c_earn, l_earn, False),
                    (c_tot_epk, l_tot_epk, False),
                    (c_fpd_epk, l_fpd_epk, False),
                    (c_mhl_epk, l_mhl_epk, False),
                    (c_pax, l_pax, True)
                ]

                html.append("<tr>")
                html.append(f'<td style="font-weight:700; text-align:left;">{m_label}</td>')

                row_export = {"Month": m_label}
                kpi_names = ["KMS", "EARNINGS", "TOT_EPK", "FPD_EPK", "MHL_EPK", "PAX"]

                for idx, (cy_val, ly_val, is_pax) in enumerate(metrics):
                    var_val = cy_val - ly_val if pd.notna(cy_val) and pd.notna(ly_val) else np.nan
                    pct_val = (var_val * 100 / ly_val) if pd.notna(ly_val) and ly_val > 0 else np.nan

                    if is_pax:
                        c_str = f"{int(cy_val):,}" if pd.notna(cy_val) else ""
                        l_str = f"{int(ly_val):,}" if pd.notna(ly_val) else ""
                        v_str = f"{int(var_val):,}" if pd.notna(var_val) else ""
                    else:
                        c_str = fmt(cy_val) if pd.notna(cy_val) else ""
                        l_str = fmt(ly_val) if pd.notna(ly_val) else ""
                        v_str = fmt(var_val) if pd.notna(var_val) else ""

                    p_str = fmt_growth(pct_val) if pd.notna(pct_val) else ""

                    html.append(f'<td>{c_str}</td>')
                    html.append(f'<td>{l_str}</td>')
                    html.append(f'<td class="{var_class(var_val)}">{v_str}</td>')
                    html.append(f'<td class="{var_class(pct_val)}">{p_str}</td>')

                    prefix_kpi = kpi_names[idx]
                    row_export[f"{prefix_kpi}_CY"] = cy_val
                    row_export[f"{prefix_kpi}_LY"] = ly_val
                    row_export[f"{prefix_kpi}_VAR"] = var_val
                    row_export[f"{prefix_kpi}_PCT"] = pct_val

                html.append("</tr>")
                export_rows.append(row_export)

            # Add UPTO Total Row
            upto_label = f"UPTO {month_label_map[active_fy_months[-1]]}"
            
            u_c_kms = tot_c_kms / 100000
            u_l_kms = tot_l_kms / 100000
            u_c_earn = tot_c_earn / 100000
            u_l_earn = tot_l_earn / 100000
            u_c_tot_epk = (tot_c_earn / tot_c_kms) if tot_c_kms > 0 else np.nan
            u_l_tot_epk = (tot_l_earn / tot_l_kms) if tot_l_kms > 0 else np.nan
            u_c_fpd_epk = (tot_c_fpd_earn / tot_c_kms) if tot_c_kms > 0 else np.nan
            u_l_fpd_epk = (tot_l_fpd_earn / tot_l_kms) if tot_l_kms > 0 else np.nan
            u_c_mhl_epk = (tot_c_mhl_earn / tot_c_kms) if tot_c_kms > 0 else np.nan
            u_l_mhl_epk = (tot_l_mhl_earn / tot_l_kms) if tot_l_kms > 0 else np.nan
            
            u_c_pax = tot_c_pax
            u_l_pax = tot_l_pax

            upto_metrics = [
                (u_c_kms, u_l_kms, False),
                (u_c_earn, u_l_earn, False),
                (u_c_tot_epk, u_l_tot_epk, False),
                (u_c_fpd_epk, u_l_fpd_epk, False),
                (u_c_mhl_epk, u_l_mhl_epk, False),
                (u_c_pax, u_l_pax, True)
            ]

            html.append('<tr style="font-weight:700; background-color:#f1f5f9;">')
            html.append(f'<td style="text-align:left;">{upto_label}</td>')

            upto_export = {"Month": upto_label}
            for idx, (cy_val, ly_val, is_pax) in enumerate(upto_metrics):
                var_val = cy_val - ly_val if pd.notna(cy_val) and pd.notna(ly_val) else np.nan
                pct_val = (var_val * 100 / ly_val) if pd.notna(ly_val) and ly_val > 0 else np.nan

                if is_pax:
                    c_str = f"{int(cy_val):,}" if pd.notna(cy_val) else ""
                    l_str = f"{int(ly_val):,}" if pd.notna(ly_val) else ""
                    v_str = f"{int(var_val):,}" if pd.notna(var_val) else ""
                else:
                    c_str = fmt(cy_val) if pd.notna(cy_val) else ""
                    l_str = fmt(ly_val) if pd.notna(ly_val) else ""
                    v_str = fmt(var_val) if pd.notna(var_val) else ""

                p_str = fmt_growth(pct_val) if pd.notna(pct_val) else ""

                html.append(f'<td>{c_str}</td>')
                html.append(f'<td>{l_str}</td>')
                html.append(f'<td class="{var_class(var_val)}">{v_str}</td>')
                html.append(f'<td class="{var_class(pct_val)}">{p_str}</td>')

                prefix_kpi = kpi_names[idx]
                upto_export[f"{prefix_kpi}_CY"] = cy_val
                upto_export[f"{prefix_kpi}_LY"] = ly_val
                upto_export[f"{prefix_kpi}_VAR"] = var_val
                upto_export[f"{prefix_kpi}_PCT"] = pct_val

            html.append("</tr>")
            export_rows.append(upto_export)

            html.append("</table></div>")
            st.markdown("".join(html), unsafe_allow_html=True)

            # CSV Download Button
            export_df = pd.DataFrame(export_rows)
            st.download_button(
                "Download CSV",
                export_df.to_csv(index=False).encode("utf-8"),
                f"Monthwise_Vertical_Actuals_{selected_fy_str}_{month}_{depot}.csv",
                "text/csv",
                key=f"dl_vertical_fy_filter_{selected_fy}_{month}_{depot}"
            )

# ==================== TAB 6 ====================
with tab6:
    prep_text = "upto" if for_upto == "UPTO" else "for"
    loc_text = f"{depot} Depot" if depot != "ALL" else "ALL Depots"
    title_tab5 = f"Month Wise Actual vs Actuals of {loc_text} {prep_text} the month of {month}"
    st.markdown(f'<div class="title-bar">{title_tab5}</div>', unsafe_allow_html=True)

    # Filter data matching base filters (Depot, Route, Product, MHL/NMHL, RTC/HIRE)
    m_base = df[base_mask].copy()

    if len(m_base) == 0:
        st.warning("No data found for selected filters.")
    else:
        # Dynamically determine months present in the filtered dataset in chronological order
        unique_months = m_base["Month_Name"].dropna().unique()
        def parse_m_sort(m_str):
            try:
                return pd.to_datetime(m_str, format="%b-%y")
            except Exception:
                try:
                    return pd.to_datetime(m_str, format="%b-%Y")
                except Exception:
                    return pd.to_datetime(m_str, errors="coerce")

        months_order = sorted([str(m).strip() for m in unique_months if str(m).strip()], key=parse_m_sort)

        if not months_order:
            st.warning("No valid month data available.")
        else:
            # 1. Aggregate Month-wise values for CY
            m_cy_agg = m_base.groupby("Month_Name").agg(
                kms=("Optd_KMs", "sum"),
                earn_tot=(earn_tot, "sum"),
                earn_fpd=(earn_fpd, "sum"),
                earn_mhl=(earn_mhl, "sum"),
                pax=(pax_col, "sum")
            ).reindex(months_order).fillna(0)

            # Calculated Monthly Metrics
            m_chart = pd.DataFrame(index=months_order)
            m_chart["Month_Group"] = m_chart.index
            m_chart["Kms_Lakhs"] = (m_cy_agg["kms"] / 100000).round(2)
            m_chart["Earn_Lakhs"] = (m_cy_agg["earn_tot"] / 100000).round(2)
            m_chart["Pax_Lakhs"] = (m_cy_agg["pax"] / 100000).round(0).astype(int)
            m_chart["Tot_EPK"] = np.where(m_cy_agg["kms"] > 0, (m_cy_agg["earn_tot"] / m_cy_agg["kms"]).round(2), 0.0)
            m_chart["FPD_EPK"] = np.where(m_cy_agg["kms"] > 0, (m_cy_agg["earn_fpd"] / m_cy_agg["kms"]).round(2), 0.0)
            m_chart["MHL_EPK"] = np.where(m_cy_agg["kms"] > 0, (m_cy_agg["earn_mhl"] / m_cy_agg["kms"]).round(2), 0.0)

            # 2. Render 5 Top Bar Charts
            def create_card_chart(df_chart, y_col, chart_title, bar_color="#1d4ed8", is_pax=False):
                fig = px.bar(
                    df_chart,
                    x="Month_Group",
                    y=y_col,
                    title=f"<b>{chart_title}</b>",
                    text=y_col,
                    color_discrete_sequence=[bar_color]
                )
                text_fmt = '%{text:.0f}' if is_pax else '%{text:.2f}'
                fig.update_traces(
                    texttemplate=text_fmt,
                    textposition='outside',
                    textangle=-90
                )
                fig.update_layout(
                    height=250,
                    margin=dict(l=5, r=5, t=30, b=5),
                    xaxis_title="",
                    yaxis_title="",
                    xaxis=dict(tickangle=-90, showgrid=False),
                    yaxis=dict(showgrid=True, showticklabels=False),
                    template="plotly_white",
                    title_font_size=11,
                    title_x=0.5
                )
                return fig

            ch1, ch2, ch3, ch4, ch5 = st.columns(5)
            with ch1:
                st.plotly_chart(create_card_chart(m_chart, "Kms_Lakhs", "KILOMETERS (in lks.)", "#0284c7"), use_container_width=True)
            with ch2:
                st.plotly_chart(create_card_chart(m_chart, "Tot_EPK", f"{prefix} TOT. E.P.K (in Ps/kms.)", "#2563eb"), use_container_width=True)
            with ch3:
                st.plotly_chart(create_card_chart(m_chart, "FPD_EPK", f"{prefix} FPD. E.P.K (in Ps/kms.)", "#16a34a"), use_container_width=True)
            with ch4:
                st.plotly_chart(create_card_chart(m_chart, "MHL_EPK", f"{prefix} MHL. E.P.K (in Ps/kms.)", "#9333ea"), use_container_width=True)
            with ch5:
                pax_title = {"FPD": "FPD PASSENGERS (in lks.)", "MHL": "MHL PASSENGERS (in lks.)"}.get(passengers, "TOTAL PASSENGERS (in lks.)")
                st.plotly_chart(create_card_chart(m_chart, "Pax_Lakhs", pax_title, "#db2777", is_pax=True), use_container_width=True)

            # 3. Cumulative CY & LY Totals using Global Filtering Rules
            cy_kms_tot = cy_data["Optd_KMs"].sum() / 100000 if len(cy_data) else 0
            ly_kms_tot = ly_data["Optd_KMs"].sum() / 100000 if len(ly_data) else 0

            cy_earn_tot = cy_data[earn_tot].sum() / 100000 if len(cy_data) else 0
            ly_earn_tot = ly_data[earn_tot].sum() / 100000 if len(ly_data) else 0

            cy_epk_tot = (cy_data[earn_tot].sum() / cy_data["Optd_KMs"].sum()) if len(cy_data) and cy_data["Optd_KMs"].sum() > 0 else np.nan
            ly_epk_tot = (ly_data[earn_tot].sum() / ly_data["Optd_KMs"].sum()) if len(ly_data) and ly_data["Optd_KMs"].sum() > 0 else np.nan

            cy_epk_fpd = (cy_data[earn_fpd].sum() / cy_data["Optd_KMs"].sum()) if len(cy_data) and cy_data["Optd_KMs"].sum() > 0 else np.nan
            ly_epk_fpd = (ly_data[earn_fpd].sum() / ly_data["Optd_KMs"].sum()) if len(ly_data) and ly_data["Optd_KMs"].sum() > 0 else np.nan

            cy_epk_mhl = (cy_data[earn_mhl].sum() / cy_data["Optd_KMs"].sum()) if len(cy_data) and cy_data["Optd_KMs"].sum() > 0 else np.nan
            ly_epk_mhl = (ly_data[earn_mhl].sum() / ly_data["Optd_KMs"].sum()) if len(ly_data) and ly_data["Optd_KMs"].sum() > 0 else np.nan

            cy_pax_tot = round(cy_data[pax_col].sum() / 100000) if len(cy_data) else 0
            ly_pax_tot = round(ly_data[pax_col].sum() / 100000) if len(ly_data) else 0

            # Construct Month-Wise Dashboard Rows
            rows_data = [
                ("KILOMETERS (in lks.)", m_chart["Kms_Lakhs"], cy_kms_tot, ly_kms_tot, False),
                (f"{prefix} EARNINGS (in lks.)", m_chart["Earn_Lakhs"], cy_earn_tot, ly_earn_tot, False),
                (f"{prefix} TOT. E.P.K (in Ps/kms.)", m_chart["Tot_EPK"], cy_epk_tot, ly_epk_tot, False),
                (f"{prefix} FPD. E.P.K (in Ps/kms.)", m_chart["FPD_EPK"], cy_epk_fpd, ly_epk_fpd, False),
                (f"{prefix} MHL. E.P.K (in Ps/kms.)", m_chart["MHL_EPK"], cy_epk_mhl, ly_epk_mhl, False),
                (pax_title, m_chart["Pax_Lakhs"], cy_pax_tot, ly_pax_tot, True),
            ]

            # 4. Render HTML Matrix Table
            html = ['<div style="overflow-x:auto;"><table class="excel-table">']
            html.append("<tr>")
            html.append('<th class="header-left">PERFORMANCE PARAMETER</th>')
            for m_name in months_order:
                html.append(f'<th class="header-sub">{m_name}</th>')
            html.append(f'<th class="header-tot">{for_upto} CY</th>')
            html.append(f'<th class="header-tot">{for_upto} LY</th>')
            html.append('<th class="header-tot">VAR</th>')
            html.append('<th class="header-tot">% ⬆/⬇</th>')
            html.append("</tr>")

            export_rows = []
            for param_label, month_vals, cy_val, ly_val, is_pax in rows_data:
                var_val = cy_val - ly_val if pd.notna(cy_val) and pd.notna(ly_val) else np.nan
                pct_val = (var_val * 100 / ly_val) if (pd.notna(ly_val) and ly_val != 0) else np.nan

                row_dict = {"Parameter": param_label}
                html.append("<tr>")
                html.append(f'<td style="font-weight:700; text-align:left;">{param_label}</td>')

                for m_name in months_order:
                    v = month_vals.get(m_name, 0)
                    cell_fmt = fmt_pax(v) if is_pax else fmt(v)
                    html.append(f'<td>{cell_fmt}</td>')
                    row_dict[m_name] = v

                # Append Totals & Variance
                c_str = fmt_pax(cy_val) if is_pax else fmt(cy_val)
                l_str = fmt_pax(ly_val) if is_pax else fmt(ly_val)
                v_str = fmt_pax(var_val) if is_pax else fmt(var_val)
                p_str = fmt_growth(pct_val)

                html.append(f'<td style="font-weight:700; background:#f1f5f9;">{c_str}</td>')
                html.append(f'<td style="font-weight:700; background:#f1f5f9;">{l_str}</td>')
                html.append(f'<td class="{var_class(var_val)}">{v_str}</td>')
                html.append(f'<td class="{var_class(pct_val)}">{p_str}</td>')
                html.append("</tr>")

                row_dict["CY"] = cy_val
                row_dict["LY"] = ly_val
                row_dict["VAR"] = var_val
                row_dict["PCT"] = pct_val
                export_rows.append(row_dict)

            html.append("</table></div>")
            st.markdown("".join(html), unsafe_allow_html=True)

            # CSV Export with Unique Key
            export_df = pd.DataFrame(export_rows)
            st.download_button(
                "Download CSV",
                export_df.to_csv(index=False).encode("utf-8"),
                f"Monthwise_ACT_VS_ACT_{month}_{depot}.csv",
                "text/csv",
                key=f"dl5_tab5_{month}_{depot}"
            )

# ==================== TAB 7 ====================
with tab7:
    title7 = f"Service WISE, Day Wise {prefix} EPK ({for_upto}) - {month}"
    if depot != "ALL":
        title7 += f" of {depot} Depot"
    st.markdown(f'<div class="title-bar">{title7}</div>', unsafe_allow_html=True)
    st.caption(f"CY rows: {len(cy_data):,} | LY rows: {len(ly_data):,}")
    if len(cy_data) == 0:
        st.warning("No data for selected filters.")
    elif not service_col:
        st.warning("Service column not found in dataset.")
    else:
        all_parts = []
        for label, earn_col in [("TOT", earn_tot), ("FPD", earn_fpd), ("MHL", earn_mhl)]:
            cy = weighted_epk(cy_data, earn_col, group_cols=["DEPOT", service_col, "ROUTEE", "PRODUCT"])
            ly = weighted_epk(ly_data, earn_col, group_cols=["DEPOT", service_col, "ROUTEE", "PRODUCT"])
            all_idx = cy.index.union(ly.index) if len(ly) > 0 else cy.index
            cy = cy.reindex(all_idx)
            ly = ly.reindex(all_idx) if len(ly) > 0 else pd.DataFrame(np.nan, index=all_idx, columns=day_order + ["UPTO"], dtype=float)
            block = pd.DataFrame(index=all_idx)
            for d in day_order:
                block[f"{label}_{d}"] = pd.to_numeric(cy[d] if d in cy.columns else np.nan, errors="coerce")
            block[f"{label}_CY"] = pd.to_numeric(cy["UPTO"] if "UPTO" in cy.columns else np.nan, errors="coerce")
            block[f"{label}_LY"] = pd.to_numeric(ly["UPTO"] if "UPTO" in ly.columns else np.nan, errors="coerce")
            block[f"{label}_Var"] = (block[f"{label}_CY"] - block[f"{label}_LY"]).astype(float).round(2)
            all_parts.append(block)
        result7 = pd.concat(all_parts, axis=1).reset_index()
        result7 = result7.rename(columns={"ROUTEE": "ROUTE", service_col: "SERVICE NO"})
        result7 = result7.sort_values(["DEPOT", "SERVICE NO", "ROUTE", "PRODUCT"]).reset_index(drop=True)
        html = ['<div style="overflow-x:auto;"><table class="excel-table">']
        html.append("<tr>")
        html.append('<th class="header-left" rowspan="2">DEPOT</th>')
        html.append('<th class="header-left" rowspan="2">SERVICE NO</th>')
        html.append('<th class="header-left" rowspan="2">ROUTE</th>')
        html.append('<th class="header-left" rowspan="2">PRODUCT</th>')
        html.append(f'<th class="header-tot" colspan="10">{prefix} TOT. E.P.K (in Ps/kms.)</th>')
        html.append(f'<th class="header-fpd" colspan="10">{prefix} FPD. E.P.K (in Ps/kms.)</th>')
        html.append(f'<th class="header-mhl" colspan="10">{prefix} MHL. E.P.K (in Ps/kms.)</th>')
        html.append("</tr><tr>")
        for _ in range(3):
            for d in day_short:
                html.append(f'<th class="header-sub">{d}</th>')
            html.append(f'<th class="header-sub">{for_upto} CY</th>')
            html.append(f'<th class="header-sub">{for_upto} LY</th>')
            html.append('<th class="header-sub">Var</th>')
        html.append("</tr>")
        for _, row in result7.iterrows():
            html.append("<tr>")
            html.append(f'<td>{row["DEPOT"]}</td>')
            html.append(f'<td>{row["SERVICE NO"]}</td>')
            html.append(f'<td>{row["ROUTE"]}</td>')
            html.append(f'<td>{row["PRODUCT"]}</td>')
            for label in ["TOT", "FPD", "MHL"]:
                # Peak / Slack among Mon-Sun for this metric on this row
                day_vals = {}
                for d in day_order:
                    v = row.get(f"{label}_{d}")
                    try:
                        fv = float(v) if pd.notna(v) else None
                    except Exception:
                        fv = None
                    if fv is not None and fv != 0:
                        day_vals[d] = fv
                peak_d = max(day_vals, key=day_vals.get) if day_vals else None
                slack_d = min(day_vals, key=day_vals.get) if day_vals else None
                for d in day_order:
                    val = row.get(f"{label}_{d}")
                    style = ""
                    if peak_d and d == peak_d:
                        style = 'background-color:#c6efce; color:#006100; font-weight:600;'
                    elif slack_d and d == slack_d:
                        style = 'background-color:#ffc7ce; color:#9c0006; font-weight:600;'
                    html.append(f'<td style="{style}">{fmt(val)}</td>')
                html.append(f'<td>{fmt(row.get(f"{label}_CY"))}</td>')
                html.append(f'<td>{fmt(row.get(f"{label}_LY"))}</td>')
                var_val = row.get(f"{label}_Var")
                html.append(f'<td class="{var_class(var_val)}">{fmt(var_val)}</td>')
            html.append("</tr>")
        html.append("</table></div>")
        st.markdown("".join(html), unsafe_allow_html=True)
        st.markdown("""
<span style="background:#c6efce; padding:2px 8px;">Peak day</span> = Highest EPK among Mon–Sun &nbsp;
<span style="background:#ffc7ce; padding:2px 8px;">Slack day</span> = Lowest EPK among Mon–Sun
&nbsp;(per row, for each of TOT / FPD / MHL)
""", unsafe_allow_html=True)
        st.download_button("Download CSV", result7.to_csv(index=False).encode("utf-8"), f"Service_Performance_{month}_{depot}.csv", "text/csv", key="dl7")

# ==================== TAB 8 ====================
with tab8:
    st.markdown('<div class="title-bar">Date-Range Set vs Date-Range Set Variance Analysis</div>', unsafe_allow_html=True)
    
    # Base dataset filtered by cascading filters (excluding month/upto date limitations)
    filtered_df = df[base_mask].copy()
    
    if filtered_df.empty or "Date" not in filtered_df.columns:
        st.warning("No data available for date range selection under current filters.")
    else:
        min_date = filtered_df["Date"].min().date()
        max_date = filtered_df["Date"].max().date()
        
        c_set1, c_set2 = st.columns(2)
        with c_set1:
            st.markdown("##### Period Set 1 (Target Period)")
            set1_range = st.date_input(
                "Select Date Range for Set 1",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                key="tab8_set1_dates"
            )
        with c_set2:
            st.markdown("##### Period Set 2 (Comparison Period)")
            set2_range = st.date_input(
                "Select Date Range for Set 2",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                key="tab8_set2_dates"
            )
            
        if isinstance(set1_range, (list, tuple)) and len(set1_range) == 2 and isinstance(set2_range, (list, tuple)) and len(set2_range) == 2:
            s1_start, s1_end = pd.Timestamp(set1_range[0]), pd.Timestamp(set1_range[1])
            s2_start, s2_end = pd.Timestamp(set2_range[0]), pd.Timestamp(set2_range[1])
            
            data_set1 = filtered_df[(filtered_df["Date"] >= s1_start) & (filtered_df["Date"] <= s1_end)]
            data_set2 = filtered_df[(filtered_df["Date"] >= s2_start) & (filtered_df["Date"] <= s2_end)]
            
            p1_label = f"{s1_start.strftime('%d-%b-%Y')} to {s1_end.strftime('%d-%b-%Y')}"
            p2_label = f"{s2_start.strftime('%d-%b-%Y')} to {s2_end.strftime('%d-%b-%Y')}"
            
            st.caption(f"Comparing **{p1_label}** vs **{p2_label}**")
            
            html_str, merged_df = build_act_vs_act_table(
                group_col="DEPOT",
                data_cy=data_set1,
                data_ly=data_set2,
                cy_label="SET 1",
                ly_label="SET 2"
            )
            
            if html_str is None:
                st.warning("No data matching the selected date ranges.")
            else:
                st.markdown(html_str, unsafe_allow_html=True)
                st.download_button(
                    "Download CSV",
                    merged_df.to_csv(index=False).encode("utf-8"),
                    f"Period_Set_Comparison_{depot}.csv",
                    "text/csv",
                    key="dl8"
                )
        else:
            st.info("Please select valid start and end dates for both Period Set 1 and Period Set 2.")

st.caption("Cascading filters • Weighted EPK • Self-hosted on your PC")