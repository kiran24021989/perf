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
    .block-container { padding-top: 2rem !important; padding-bottom: 1rem !important; max-width: 100% !important; overflow: visible !important; }

    div[data-testid="stVerticalBlock"] {
        gap: 0.4rem !important;
    }

    /* Filter Panel */
    div[data-testid="stHorizontalBlock"]:has(div[data-testid="stSelectbox"]) {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 10px 14px !important;
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.15);
        border: 1px solid #334155;
        margin-top: 8px !important;
        margin-bottom: 12px !important;
        gap: 8px !important;
    }

    div[data-testid="stSelectbox"] label p {
        color: #FDFBF7 !important;
        font-family: 'Segoe UI', Roboto, sans-serif !important;
        font-weight: 700 !important;
        font-size: 12px !important;
        letter-spacing: 0.4px;
        text-transform: uppercase;
        margin-bottom: 4px !important;
        line-height: 1.2 !important;
        white-space: nowrap !important;
    }

    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 12px !important;
        font-family: 'Segoe UI', Roboto, sans-serif !important;
        min-height: 32px !important;
        height: 32px !important;
        padding-top: 0px !important;
        padding-bottom: 0px !important;
        padding-left: 8px !important;
        line-height: 32px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }

    div[data-testid="stSelectbox"] div[data-baseweb="select"] svg {
        width: 14px !important;
        height: 14px !important;
    }

    div[data-testid="stTabs"] {
        margin-top: 4px !important;
        padding-top: 0px !important;
    }

    div[data-testid="stTabs"] div[role="tablist"] {
        gap: 6px;
        border-bottom: none !important;
        flex-wrap: wrap !important;
    }

    div[data-testid="stTabs"] button {
        font-family: 'Segoe UI', Roboto, sans-serif !important;
        font-weight: 700 !important;
        font-size: 12px !important;
        padding: 6px 14px !important;
        color: #ffffff !important;
        background-color: #e67e22 !important;
        border-radius: 4px !important;
        border: 1px solid #d35400 !important;
        margin-right: 2px !important;
    }
    div[data-testid="stTabs"] button p { color: #ffffff !important; }
    div[data-testid="stTabs"] button:hover {
        background-color: #f39c12 !important;
        color: #ffffff !important;
    }
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: #ffffff !important;
        background-color: #27ae60 !important;
        border-color: #1e8449 !important;
    }
    div[data-testid="stTabs"] button[aria-selected="true"] p { color: #ffffff !important; }
    div[data-testid="stTabs"] [data-baseweb="tab-highlight"] { display: none !important; }

    div[data-testid="stDownloadButton"] button {
        padding: 4px 12px !important;
        font-size: 12px !important;
        min-height: 32px !important;
    }

    hr { margin: 8px 0 !important; }
    h4 { margin: 8px 0 6px 0 !important; font-size: 16px !important; }

    .excel-table {
        border-collapse: collapse;
        width: 100%;
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 12px;
        margin-top: 6px;
    }
    .excel-table th, .excel-table td {
        border: 1px solid #e2e8f0;
        padding: 4px 6px;
        text-align: center;
        white-space: nowrap;
        font-size: 12px;
    }
    .excel-table th { font-weight: 700; font-size: 11px; text-transform: uppercase; }
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
        padding: 6px;
        border-radius: 4px;
        border: 1px solid #fef08a;
        margin-bottom: 8px;
        font-family: 'Segoe UI', Roboto, sans-serif !important;
    }

    .freeze-wrap, .op-wrap {
        max-height: 65vh;
        overflow: auto;
        border: 1px solid #e2e8f0;
        margin-bottom: 10px;
        position: relative;
        z-index: 0;
    }
    .freeze-table, .op-table {
        border-collapse: separate;
        border-spacing: 0;
        font-size: 12px;
        width: max-content;
    }
    .freeze-table th, .freeze-table td,
    .op-table th, .op-table td {
        border: 1px solid #e2e8f0;
        white-space: nowrap;
    }
    .freeze-table thead tr:nth-child(1) th { position: sticky; top: 0; z-index: 4; }
    .freeze-table thead tr:nth-child(2) th { position: sticky; top: 30px; z-index: 4; }
    .freeze-table thead tr:nth-child(3) th { position: sticky; top: 58px; z-index: 4; }
    .op-table thead tr:first-child th { position: sticky; top: 0; z-index: 4; }
    .op-table thead tr:nth-child(2) th { position: sticky; top: 34px; z-index: 4; }
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
st.markdown("<!-- spacer for streamlit header -->", unsafe_allow_html=True)
# HEADING ABOVE FILTERS
st.title("HISTORICAL ANALYSIS OF RANGAREDDY REGION")
st.markdown(
    '<div style="height:4px;background:#e67e22;margin:-8px 0 12px 0;border-radius:2px;"></div>',
    unsafe_allow_html=True,
)

st.markdown("<div style='font-size:13px;font-weight:600;color:#334155;margin-bottom:2px;'>Filters</div>", unsafe_allow_html=True)
temp = df.copy()
c1, c2, c3, c4, c5, c6, c7, c8, c9 = st.columns(9)
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
    pax_opts = ["TOT", "FPD", "MHL"]
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
with c8:
    for_upto = st.selectbox("For / Upto", ["UPTO", "FOR"], index=0)
with c9:
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

# Service column name for Service Performance tab only
service_col = next((col for col in ["SER_NO", "SERVICE_NO", "SERVICE", "SERVICE_NUMBER", "ServiceNo"] if col in df.columns), None)

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
    html = ['<div class="table-scroll"><table class="excel-table"><thead>']
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

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs(["Route Day-wise", "ACT VS ACT", "Product wise", "Day wise", "trends", "cy trends", "Service performance", "Period Comparison", "Schedules"])

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
        html = ['<div class="table-scroll"><table class="excel-table">']
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
        html.append("</thead><tbody>")
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
        html.append("</tbody></table></div>")
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

    # ===== Daily Depot table for selected month =====
    st.markdown("---")
    st.markdown(f'<div class="title-bar">Daily Depot Performance – {month}</div>', unsafe_allow_html=True)

    # Build calendar month date range from month filter (e.g. Apr-2026 → 2026-04-01 to 2026-04-30)
    mon_map = {
        "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
        "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
    }
    try:
        parts = month.replace(" ", "").split("-")
        m_num = mon_map.get(parts[0][:3], 1)
        y_num = int(parts[1]) if len(parts[1]) == 4 else int("20" + parts[1])
        month_start = pd.Timestamp(year=y_num, month=m_num, day=1)
        # last day of month
        if m_num == 12:
            month_end = pd.Timestamp(year=y_num, month=12, day=31)
        else:
            month_end = pd.Timestamp(year=y_num, month=m_num + 1, day=1) - pd.Timedelta(days=1)
    except Exception:
        month_start = selected_max_date.replace(day=1) if not pd.isna(selected_max_date) else None
        month_end = selected_max_date if not pd.isna(selected_max_date) else None

    if month_start is not None:
        daily_mask = base_mask & (df["Date"] >= month_start) & (df["Date"] <= month_end)
        daily_df = df[daily_mask].copy()
        if len(daily_df) == 0:
            st.warning("No daily data for selected month and filters.")
        else:
            g = daily_df.groupby(["DEPOT", "Date"]).agg(
                Kilometers=("Optd_KMs", "sum"),
                Earnings=(earn_tot, "sum"),
                Earn_FPD=(earn_fpd, "sum"),
                Earn_MHL=(earn_mhl, "sum"),
                Passengers=(pax_col, "sum"),
            ).reset_index()
            g["EPK_TOT"] = np.where(g["Kilometers"] > 0, g["Earnings"] / g["Kilometers"], np.nan)
            g["EPK_FPD"] = np.where(g["Kilometers"] > 0, g["Earn_FPD"] / g["Kilometers"], np.nan)
            g["EPK_MHL"] = np.where(g["Kilometers"] > 0, g["Earn_MHL"] / g["Kilometers"], np.nan)
            g = g.sort_values(["DEPOT", "Date"]).reset_index(drop=True)
            g["_dt"] = pd.to_datetime(g["Date"])
            g["Weekday"] = g["_dt"].dt.strftime("%a")  # Mon, Tue, ...
            g["Date"] = g["_dt"].dt.strftime("%d-%m-%Y")
            g["Kilometers"] = (g["Kilometers"] / 100000).round(2)
            g["Earnings"] = (g["Earnings"] / 100000).round(2)
            g["EPK_TOT"] = g["EPK_TOT"].round(2)
            g["EPK_FPD"] = g["EPK_FPD"].round(2)
            g["EPK_MHL"] = g["EPK_MHL"].round(2)
            g["Passengers"] = g["Passengers"].round(0)

            pax_heading = {"FPD": "FPD PASSENGERS", "MHL": "MHL PASSENGERS"}.get(passengers, "TOTAL PASSENGERS")

            # Compact table: ~header + 31 rows visible per viewport
            html_d = ['<div class="table-scroll" style="max-height: 75vh;"><table class="excel-table" style="font-size:10px;">']
            html_d.append("<tr>")
            html_d.append('<th class="header-left" style="padding:2px 3px; font-size:9px;">DEPOT</th>')
            html_d.append('<th class="header-left" style="padding:2px 3px; font-size:9px;">DATE</th>')
            html_d.append('<th class="header-left" style="padding:2px 3px; font-size:9px;">WEEKDAY</th>')
            html_d.append('<th class="header-km" style="padding:2px 3px; font-size:9px;">KMs (lks.)</th>')
            html_d.append(f'<th class="header-earn" style="padding:2px 3px; font-size:9px;">{prefix} EARN (lks.)</th>')
            html_d.append(f'<th class="header-tot" style="padding:2px 3px; font-size:9px;">{prefix} TOT EPK</th>')
            html_d.append(f'<th class="header-fpd" style="padding:2px 3px; font-size:9px;">{prefix} FPD EPK</th>')
            html_d.append(f'<th class="header-mhl" style="padding:2px 3px; font-size:9px;">{prefix} MHL EPK</th>')
            html_d.append(f'<th class="header-left" style="padding:2px 3px; font-size:9px;">{pax_heading}</th>')
            html_d.append("</tr>")
            # Peak / Slack by TOT EPK within each depot
            peak_dates = set()
            slack_dates = set()
            for dep, grp in g.groupby("DEPOT"):
                vals = grp[["Date", "EPK_TOT"]].dropna()
                vals = vals[vals["EPK_TOT"] != 0]
                if len(vals) == 0:
                    continue
                # top 2 peak, bottom 2 slack
                sorted_v = vals.sort_values("EPK_TOT", ascending=False)
                for d in sorted_v.head(2)["Date"]:
                    peak_dates.add((dep, d))
                for d in sorted_v.tail(2)["Date"]:
                    slack_dates.add((dep, d))
            # avoid same day marked both if few days
            slack_dates = slack_dates - peak_dates

            for _, r in g.iterrows():
                key = (r["DEPOT"], r["Date"])
                row_style = ""
                date_style = "font-weight:700;"
                if key in peak_dates:
                    row_style = "background-color:#c6efce;"
                    date_style = "font-weight:700; color:#006100;"
                elif key in slack_dates:
                    row_style = "background-color:#ffc7ce;"
                    date_style = "font-weight:700; color:#9c0006;"
                html_d.append(f'<tr style="{row_style}">')
                html_d.append(f'<td style="font-weight:700; padding:1px 3px;">{r["DEPOT"]}</td>')
                html_d.append(f'<td style="{date_style} padding:1px 3px;">{r["Date"]}</td>')
                html_d.append(f'<td style="padding:1px 3px;">{r["Weekday"]}</td>')
                html_d.append(f'<td style="padding:1px 3px;">{fmt(r["Kilometers"])}</td>')
                html_d.append(f'<td style="padding:1px 3px;">{fmt(r["Earnings"])}</td>')
                html_d.append(f'<td style="font-weight:700; padding:1px 3px;">{fmt(r["EPK_TOT"])}</td>')
                html_d.append(f'<td style="padding:1px 3px;">{fmt(r["EPK_FPD"])}</td>')
                html_d.append(f'<td style="padding:1px 3px;">{fmt(r["EPK_MHL"])}</td>')
                html_d.append(f'<td style="padding:1px 3px;">{fmt_pax(r["Passengers"])}</td>')
                html_d.append("</tr>")
            html_d.append("</table></div>")
            st.markdown("""
<span style="background:#c6efce; padding:2px 8px; font-weight:700;">Peak</span> = Top 2 TOT EPK days per depot &nbsp;
<span style="background:#ffc7ce; padding:2px 8px; font-weight:700;">Slack</span> = Bottom 2 TOT EPK days per depot
""", unsafe_allow_html=True)
            st.markdown("".join(html_d), unsafe_allow_html=True)
            st.caption(f"Date range: {month_start.strftime('%d-%m-%Y')} to {month_end.strftime('%d-%m-%Y')} | Rows: {len(g):,}")
            st.download_button(
                "Download Daily CSV",
                g.to_csv(index=False).encode("utf-8"),
                f"Daily_Depot_{month}.csv",
                "text/csv",
                key="dl2_daily"
            )


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
            html = ['<div class="table-scroll"><table class="excel-table">']
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
            html = ['<div class="table-scroll"><table class="excel-table">']
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
            html = ['<div class="table-scroll"><table class="excel-table">']
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

    if not service_col:
        st.warning("Service column not found in dataset.")
    else:
        # Simple service filter from performance data (SER_NO) only
        svc_vals = set()
        if len(cy_data) and service_col in cy_data.columns:
            svc_vals.update(cy_data[service_col].dropna().unique().tolist())
        if len(ly_data) and service_col in ly_data.columns:
            svc_vals.update(ly_data[service_col].dropna().unique().tolist())

        def _svc_key(v):
            s = str(v).strip()
            if s.endswith(".0"):
                s = s[:-2]
            try:
                return (0, int(float(s)))
            except Exception:
                return (1, s)

        svc_opts = ["ALL"] + sorted(
            [x for x in svc_vals if str(x).strip() and str(x).lower() != "nan"],
            key=_svc_key,
        )
        service_no = st.selectbox("SERVICE NO", svc_opts, index=0, key="tab7_service")

        if service_no == "ALL":
            cy_svc = cy_data.copy()
            ly_svc = ly_data.copy()
        else:
            # match as string and numeric
            def _match(data, svc):
                if len(data) == 0:
                    return data.copy()
                s = data[service_col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
                target = str(svc).strip()
                if target.endswith(".0"):
                    target = target[:-2]
                m = data[s == target]
                if len(m) == 0:
                    try:
                        tnum = float(target)
                        m = data[pd.to_numeric(data[service_col], errors="coerce") == tnum]
                    except Exception:
                        pass
                return m.copy()
            cy_svc = _match(cy_data, service_no)
            ly_svc = _match(ly_data, service_no)

        st.caption(f"CY rows: {len(cy_svc):,} | LY rows: {len(ly_svc):,} | service={service_no} | col={service_col}")
        if len(cy_svc) == 0 and len(ly_svc) == 0:
            st.warning("No data for selected filters.")
        else:
            all_parts = []
            for label, earn_col in [("TOT", earn_tot), ("FPD", earn_fpd), ("MHL", earn_mhl)]:
                cy = weighted_epk(cy_svc, earn_col, group_cols=["DEPOT", service_col, "ROUTEE", "PRODUCT"])
                ly = weighted_epk(ly_svc, earn_col, group_cols=["DEPOT", service_col, "ROUTEE", "PRODUCT"])
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
            html = ['<div class="table-scroll"><table class="excel-table">']
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





# ==================== TAB 9: Schedules (SMASTER) ====================
with tab9:
    st.markdown('<div class="title-bar">Schedules – SCHs / SERVICES / SCH KMS</div>', unsafe_allow_html=True)

    SCHEDULE_EXCEL = r"D:\MONTHLY\SROS.xlsx"
    SCHEDULE_SHEET = "SMASTER"

    @st.cache_data(ttl=300)
    def load_smaster(path, sheet):
        p = Path(path)
        if not p.exists():
            return None, f"File not found: {path}"
        try:
            raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
            header_row = 0
            for i in range(min(20, len(raw))):
                vals = [str(v).strip().upper().replace(" ", "") for v in raw.iloc[i].tolist()]
                joined = "|".join(vals)
                if "SERVICENO" in joined and ("DEPOT" in joined or "PRODUCT" in joined):
                    header_row = i
                    break
            sdf = pd.read_excel(path, sheet_name=sheet, header=header_row, engine="openpyxl")
            sdf.columns = [str(c).strip() for c in sdf.columns]
            return sdf.dropna(axis=1, how="all"), None
        except Exception as e:
            return None, str(e)

    sched_raw, sched_err = load_smaster(SCHEDULE_EXCEL, SCHEDULE_SHEET)
    if sched_err:
        st.warning(sched_err)
    elif sched_raw is None or len(sched_raw) == 0:
        st.warning("SMASTER empty")
    else:
        def find_col(cands):
            def normkey(s):
                return str(s).strip().lower().replace(" ", "").replace("_", "").replace("/", "").replace(".", "")
            norm = {normkey(c): c for c in sched_raw.columns}
            for cand in cands:
                k = normkey(cand)
                if k in norm:
                    return norm[k]
            for c in sched_raw.columns:
                cl = normkey(c)
                for cand in cands:
                    if normkey(cand) in cl or cl in normkey(cand):
                        return c
            # last resort: strip-only match on display name
            strip_map = {str(c).strip().upper(): c for c in sched_raw.columns}
            for cand in cands:
                if cand.strip().upper() in strip_map:
                    return strip_map[cand.strip().upper()]
            return None

        col_depot = find_col(["DEPOT"])
        col_product = find_col(["PRODUCT"])
        col_mhl = find_col(["MHL/NMHL", "MHL_NMHL", "MHLNMHL"])
        col_month = find_col(["MONTH", "Month"])
        col_year = find_col(["YEAR", "Year"])
        col_rtc = find_col(["RTC/HIRE", "RTC_HIRE"])
        col_sch = find_col(["NoOfSchedules", "NoOfSchedule"])
        col_svc = find_col(["ServiceNo", "SERVICENO"])
        col_kms = find_col(["RevenueKms", "Revenue Kms"])
        col_dtype = find_col(["D.TYPE", "DTYPE", "DutyType", "OtService", "ServiceType"])

        required = [col_depot, col_product, col_rtc, col_sch, col_svc, col_kms]
        if any(x is None for x in required):
            st.error("Missing required columns")
            st.write(list(sched_raw.columns))
        else:
            sdf = sched_raw.copy()
            if col_month and col_year:
                def make_mon(row):
                    try:
                        m = str(row[col_month]).strip()[:3].title()
                        y = int(float(row[col_year]))
                        return f"{m}-{y}"
                    except Exception:
                        return str(row[col_month])
                sdf["_MonthKey"] = sdf.apply(make_mon, axis=1)
            elif col_month:
                sdf["_MonthKey"] = sdf[col_month].astype(str).str.strip()
            else:
                sdf["_MonthKey"] = "ALL"

            def norm_rtc(v):
                s = str(v).strip().upper()
                if s in ("R", "RTC"):
                    return "RTC"
                if s in ("H", "HIRE"):
                    return "HIRE"
                return s

            sdf["_RTC"] = sdf[col_rtc].map(norm_rtc)
            sdf["_SCH"] = pd.to_numeric(sdf[col_sch], errors="coerce").fillna(0)
            sdf["_KMS"] = pd.to_numeric(sdf[col_kms], errors="coerce").fillna(0)
            sdf["_DEPOT"] = sdf[col_depot].astype(str).str.strip()
            sdf["_PRODUCT"] = sdf[col_product].astype(str).str.strip()
            sdf["_MHL"] = sdf[col_mhl].astype(str).str.strip() if col_mhl else "ALL"
            if col_dtype:
                sdf["_DTYPE"] = sdf[col_dtype].astype(str).str.strip().str.upper()
            else:
                sdf["_DTYPE"] = ""


            # ---- Shared filters (Table A uses only these) ----
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                d_opts = ["ALL", "REGION"] + sorted([str(x).strip() for x in sdf["_DEPOT"].dropna().unique() if str(x).strip() and str(x).strip().lower() != "nan"])
                f_depot = st.selectbox("DEPOT", d_opts, key="sch_depot")
            with fc2:
                mon_opts = sorted(
                    [x for x in sdf["_MonthKey"].dropna().unique() if x and x != "ALL"],
                    key=lambda m: pd.to_datetime(str(m), format="%b-%Y", errors="coerce"),
                    reverse=True,
                ) or ["ALL"]
                f_month = st.selectbox("MONTH", mon_opts, key="sch_month")
            with fc3:
                f_compare = st.selectbox("COMPARE WITH", ["LAST YEAR", "PREVIOUS MONTH"], key="sch_compare")

            # Base for Table A: only depot + month (no product/mhl/route filter)
            base_a = sdf.copy()
            if f_depot not in ("ALL", "REGION"):
                base_a = base_a[base_a["_DEPOT"] == f_depot]

            cy_sdf = base_a[base_a["_MonthKey"] == f_month].copy() if f_month != "ALL" else base_a.copy()
            ly_sdf = base_a.iloc[0:0].copy()
            ly_key = ""
            try:
                parts = str(f_month).split("-")
                mon_abbr, yr = parts[0], int(parts[1])
                mon_map = {"Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,"Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12}
                inv_map = {v: k for k, v in mon_map.items()}
                if f_compare == "LAST YEAR":
                    ly_key = f"{mon_abbr}-{yr-1}"
                else:
                    mnum = mon_map.get(mon_abbr[:3], 1)
                    if mnum == 1:
                        pm, py = 12, yr - 1
                    else:
                        pm, py = mnum - 1, yr
                    ly_key = f"{inv_map[pm]}-{py}"
                ly_sdf = base_a[base_a["_MonthKey"] == ly_key].copy()
            except Exception:
                pass

            st.caption(f"Table A filters: Depot={f_depot} | Month={f_month} | Compare={f_compare} ({ly_key or 'N/A'})")


            st.markdown("#### Product-wise Summary")

            # Labels for comparison columns
            if f_compare == "PREVIOUS MONTH":
                h_cy, h_ly, h_var = "CM", "PM", "VAR"
            else:
                h_cy, h_ly, h_var = "CY", "LY", "VAR"

            # Ensure OT / muster cols exist on cy_sdf / ly_sdf
            def ensure_ot_cols(data):
                data = data.copy()
                def parse_time_to_minutes(v):
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        return 0.0
                    if isinstance(v, (int, float)):
                        return float(v)
                    try:
                        import datetime as _dt
                        if isinstance(v, _dt.time):
                            return v.hour * 60 + v.minute + v.second / 60.0
                        if isinstance(v, _dt.timedelta):
                            return v.total_seconds() / 60.0
                        if isinstance(v, _dt.datetime):
                            return v.hour * 60 + v.minute
                    except Exception:
                        pass
                    s = str(v).strip()
                    if not s or s.lower() in ("nan", "nat", "none", "null"):
                        return 0.0
                    if ":" in s:
                        parts = s.split(":")
                        try:
                            h = int(parts[0]); m = int(parts[1]) if len(parts) > 1 else 0
                            return h * 60 + m
                        except Exception:
                            return 0.0
                    try:
                        return float(s)
                    except Exception:
                        return 0.0

                c_cond = find_col(["COND MUSTERS", "CONDMUSTERS", "COND MUSTER"])
                c_dri = find_col(["DRI MUS", "DRIMUS", "DRI MUSTER", "DRI MUSTERS"])
                c_cnd_ot = find_col(["CND OT", "CNDOT"])
                c_drv_ot = find_col(["DRV OT", "DRVOT"])
                if c_cnd_ot is None:
                    for c in data.columns:
                        if str(c).strip().upper().replace(" ", "") == "CNDOT":
                            c_cnd_ot = c; break
                if c_drv_ot is None:
                    for c in data.columns:
                        if str(c).strip().upper().replace(" ", "") == "DRVOT":
                            c_drv_ot = c; break
                data["_COND_MUS"] = pd.to_numeric(data[c_cond], errors="coerce").fillna(0) if c_cond else 0
                data["_DRI_MUS"] = pd.to_numeric(data[c_dri], errors="coerce").fillna(0) if c_dri else 0
                data["_CND_OT"] = data[c_cnd_ot].map(parse_time_to_minutes) if c_cnd_ot else 0
                data["_DRV_OT"] = data[c_drv_ot].map(parse_time_to_minutes) if c_drv_ot else 0
                return data

            cy_p = ensure_ot_cols(cy_sdf)
            ly_p = ensure_ot_cols(ly_sdf) if len(ly_sdf) else ly_sdf


            PRODUCT_ORDER = [
                "AC-SLP", "AC-HBD", "e-GRD", "RJD", "N-HBD R", "SLX", "DLX",
                "EXP-H", "ME", "PVG-R", "PVG-H", "EXP-R", "GRD+", "N-HBD", "PVG",
            ]

            def agg_product(data):
                metric_keys = []
                for side in ["RTC", "HIRE"]:
                    metric_keys += [f"SCH_{side}", f"SVC_{side}", f"KMS_{side}",
                                   f"CREW_COND_{side}", f"CREW_DRI_{side}",
                                   f"OT_COND_{side}", f"OT_DRI_{side}"]
                metric_keys += ["SCH_TOTAL", "SVC_TOTAL", "KMS_TOTAL"]
                base_cols = ["DEPOT", "PRODUCT", "MHL_NMHL"]
                if data is None or len(data) == 0:
                    return pd.DataFrame(columns=base_cols + metric_keys)
                rows = []
                if f_depot == "REGION":
                    group_keys = ["_PRODUCT", "_MHL"]
                else:
                    group_keys = ["_DEPOT", "_PRODUCT", "_MHL"]
                for keys, grp in data.groupby(group_keys):
                    if f_depot == "REGION":
                        if isinstance(keys, tuple):
                            prod, mhl = keys[0], keys[1]
                        else:
                            prod, mhl = keys, ""
                        dep = "REGION"
                    else:
                        dep, prod, mhl = keys
                    rec = {"DEPOT": dep, "PRODUCT": prod, "MHL_NMHL": mhl}
                    for side in ["RTC", "HIRE"]:
                        sub = grp[grp["_RTC"] == side]
                        rec[f"SCH_{side}"] = sub["_SCH"].sum()
                        rec[f"SVC_{side}"] = sub[col_svc].nunique()
                        rec[f"KMS_{side}"] = sub["_KMS"].sum()
                        rec[f"CREW_COND_{side}"] = sub["_COND_MUS"].sum() * 1.3
                        rec[f"CREW_DRI_{side}"] = sub["_DRI_MUS"].sum() * 1.3
                        rec[f"OT_COND_{side}"] = sub["_CND_OT"].sum()
                        rec[f"OT_DRI_{side}"] = sub["_DRV_OT"].sum()
                    rec["SCH_TOTAL"] = rec["SCH_RTC"] + rec["SCH_HIRE"]
                    rec["SVC_TOTAL"] = rec["SVC_RTC"] + rec["SVC_HIRE"]
                    rec["KMS_TOTAL"] = rec["KMS_RTC"] + rec["KMS_HIRE"]
                    rows.append(rec)
                return pd.DataFrame(rows)

            def sort_products(df):
                if len(df) == 0:
                    return df
                order_map = {p: i for i, p in enumerate(PRODUCT_ORDER)}
                df = df.copy()
                df["_ord"] = df["PRODUCT"].map(lambda x: order_map.get(str(x).strip(), 500))
                df = df.sort_values(["DEPOT", "_ord", "PRODUCT"]).drop(columns=["_ord"])
                return df.reset_index(drop=True)


            cy_agg = agg_product(cy_p)
            ly_agg = agg_product(ly_p)
            metrics = [
                "SCH_RTC", "SCH_HIRE", "SCH_TOTAL",
                "SVC_RTC", "SVC_HIRE", "SVC_TOTAL",
                "KMS_RTC", "KMS_HIRE", "KMS_TOTAL",
                "CREW_COND_RTC", "CREW_DRI_RTC", "CREW_COND_HIRE",
                "OT_COND_RTC", "OT_DRI_RTC", "OT_COND_HIRE",
            ]

            def minutes_to_hhmm(mins):
                if mins is None:
                    return ""
                try:
                    mins = float(mins)
                except Exception:
                    return ""
                if mins == 0:
                    return ""
                mins = int(round(mins))
                h, m = divmod(abs(mins), 60)
                sign = "-" if mins < 0 else ""
                return f"{sign}{h}:{m:02d}"

            if len(cy_agg) == 0 and len(ly_agg) == 0:
                st.warning("No product-wise data for selection.")
            else:
                cy_agg = sort_products(cy_agg)
                ly_agg = sort_products(ly_agg)
                merged = cy_agg.merge(ly_agg, on=["DEPOT", "PRODUCT", "MHL_NMHL"], how="outer", suffixes=("_CY", "_LY"))
                if "MHL_NMHL" not in merged.columns:
                    merged["MHL_NMHL"] = ""
                merged["MHL_NMHL"] = merged["MHL_NMHL"].fillna("")
                for m in metrics:
                    for s in ("_CY", "_LY"):
                        c = f"{m}{s}"
                        if c not in merged.columns:
                            merged[c] = 0
                        merged[c] = pd.to_numeric(merged[c], errors="coerce").fillna(0)
                    merged[f"{m}_VAR"] = merged[f"{m}_CY"] - merged[f"{m}_LY"]
                merged = sort_products(merged)

                extra_rows = []
                dep_label = str(merged["DEPOT"].iloc[0]) if len(merged) else (f_depot if f_depot != "REGION" else "REGION")
                # Only detail product rows (exclude any prior totals)
                detail = merged[~merged["PRODUCT"].astype(str).str.upper().isin(["TOTAL", "MHL", "NMHL"])].copy()
                # TOTAL = all products
                tot = {"DEPOT": dep_label, "PRODUCT": "TOTAL", "MHL_NMHL": ""}
                for m in metrics:
                    tot[f"{m}_CY"] = detail[f"{m}_CY"].sum()
                    tot[f"{m}_LY"] = detail[f"{m}_LY"].sum()
                    tot[f"{m}_VAR"] = tot[f"{m}_CY"] - tot[f"{m}_LY"]
                extra_rows.append(tot)
                # MHL = sum where MHL_NMHL is exactly MHL (not NMHL)
                # NMHL = sum where MHL_NMHL is exactly NMHL
                for label in ["MHL", "NMHL"]:
                    mask = detail["MHL_NMHL"].astype(str).str.strip().str.upper() == label
                    sub = detail[mask]
                    row = {"DEPOT": dep_label, "PRODUCT": label, "MHL_NMHL": label}
                    for m in metrics:
                        row[f"{m}_CY"] = sub[f"{m}_CY"].sum() if len(sub) else 0
                        row[f"{m}_LY"] = sub[f"{m}_LY"].sum() if len(sub) else 0
                        row[f"{m}_VAR"] = row[f"{m}_CY"] - row[f"{m}_LY"]
                    extra_rows.append(row)
                merged = pd.concat([detail, pd.DataFrame(extra_rows)], ignore_index=True)

                def cell_num(v, is_var=False, is_ot=False):
                    try:
                        fv = float(v)
                    except Exception:
                        return "<td style='padding:4px 5px;font-size:12px;'></td>"
                    if fv == 0:
                        return "<td style='padding:4px 5px;font-size:12px;'></td>"
                    if is_ot:
                        s = minutes_to_hhmm(fv)
                    else:
                        s = f"{fv:,.0f}" if abs(fv - round(fv)) < 1e-6 else f"{fv:,.1f}"
                    if is_var:
                        cls = "pos" if fv > 0 else "neg"
                        return f"<td class='{cls}' style='padding:4px 5px;font-size:12px;'>{s}</td>"
                    return f"<td style='padding:4px 5px;font-size:12px;'>{s}</td>"

                # No freeze - normal scroll, larger font & row height
                
                # ---- Build clean 3-row header + body ----
                # Column plan (after DEPOT, PRODUCT):
                # SCH: RTC×3, HIRE×3, TOTAL×3 = 9
                # SVC: 9
                # KMS: 9
                # CREW: RTC COND×3, RTC DRI×3, HIRE COND×3, HIRE DRI×3 = 12
                # OT: same 12
                # Total data cols = 9+9+9+12+12 = 51

                def th(text, colspan=1, rowspan=1, bg="#334155", extra=""):
                    return (
                        f'<th colspan="{colspan}" rowspan="{rowspan}" '
                        f'style="background:{bg};color:#fff;padding:6px 4px;font-size:11px;'
                        f'text-align:center;border:1px solid #94a3b8;{extra}">{text}</th>'
                    )

                def th_sub(text):
                    return (
                        f'<th style="background:#e2e8f0;color:#0f172a;padding:5px 4px;font-size:10px;'
                        f'text-align:center;border:1px solid #94a3b8;">{text}</th>'
                    )

                html = ['<div class="freeze-wrap"><table class="freeze-table"><thead>']

                # Row 1 – group titles
                html.append("<tr>")
                html.append(th("DEPOT", rowspan=3, bg="#0369a1", extra="position:sticky;left:0;z-index:6;min-width:70px;"))
                html.append(th("PRODUCT", rowspan=3, bg="#0369a1", extra="position:sticky;left:70px;z-index:6;min-width:80px;"))
                html.append(th("MHL/NMHL", rowspan=3, bg="#0369a1", extra="position:sticky;left:150px;z-index:6;min-width:60px;"))
                html.append(th("SCHs", colspan=9, bg="#b91c1c"))
                html.append(th("SERVICES", colspan=9, bg="#a21caf"))
                html.append(th("SCH KMS", colspan=9, bg="#15803d"))
                html.append(th("CREW REQUIREMENT", colspan=9, bg="#1d4ed8"))
                html.append(th("SCH OVER TIME", colspan=9, bg="#c2410c"))
                html.append("</tr>")

                # Row 2 – RTC / HIRE / TOTAL (and COND/DRI under crew & OT)
                html.append("<tr>")
                for _ in range(3):  # SCH, SVC, KMS
                    html.append(th("RTC", colspan=3, bg="#475569"))
                    html.append(th("HIRE", colspan=3, bg="#475569"))
                    html.append(th("TOTAL", colspan=3, bg="#475569"))
                # CREW: RTC-COND, RTC-DRI, HIRE-COND (no HIRE DRI)
                html.append(th("RTC COND", colspan=3, bg="#1e40af"))
                html.append(th("RTC DRI", colspan=3, bg="#1e40af"))
                html.append(th("HIRE COND", colspan=3, bg="#1e40af"))
                # OT: RTC-COND, RTC-DRI, HIRE-COND (no HIRE DRI)
                html.append(th("RTC COND", colspan=3, bg="#9a3412"))
                html.append(th("RTC DRI", colspan=3, bg="#9a3412"))
                html.append(th("HIRE COND", colspan=3, bg="#9a3412"))
                html.append("</tr>")

                # Row 3 – CM/PM/VAR or CY/LY/VAR
                html.append("<tr>")
                for _ in range(15):  # 9 (sch/svc/kms) + 3 crew + 3 ot
                    for h in (h_cy, h_ly, h_var):
                        html.append(th_sub(h))
                html.append("</tr>")
                html.append("</thead><tbody>")

                ot_metrics = {"OT_COND_RTC", "OT_DRI_RTC", "OT_COND_HIRE"}
                for _, row in merged.iterrows():
                    prod = str(row.get("PRODUCT", ""))
                    is_tot = prod in ("TOTAL", "MHL", "NMHL")
                    style = "font-weight:bold;background:#e2efda;" if is_tot else ""
                    html.append(f'<tr style="{style}">')
                    html.append(
                        f'<td style="position:sticky;left:0;z-index:2;background:#e0f2fe;'
                        f'padding:5px 6px;font-size:12px;font-weight:600;min-width:70px;">{row["DEPOT"]}</td>'
                    )
                    html.append(
                        f'<td style="position:sticky;left:70px;z-index:2;background:#f0f9ff;'
                        f'padding:5px 6px;font-size:12px;min-width:80px;">{row["PRODUCT"]}</td>'
                    )
                    html.append(
                        f'<td style="position:sticky;left:150px;z-index:2;background:#f8fafc;'
                        f'padding:5px 6px;font-size:11px;min-width:60px;">{row.get("MHL_NMHL", "")}</td>'
                    )
                    for m in metrics:
                        is_ot = m in ot_metrics
                        if m.startswith("SCH"):
                            bg = "#fef2f2"
                        elif m.startswith("SVC"):
                            bg = "#f5f3ff"
                        elif m.startswith("KMS"):
                            bg = "#f0fdf4"
                        elif m.startswith("CREW"):
                            bg = "#eff6ff"
                        elif m.startswith("OT"):
                            bg = "#fff7ed"
                        else:
                            bg = "#fff"
                        def cell_bg(v, is_var=False, is_ot=False, bg="#fff"):
                            try:
                                fv = float(v)
                            except Exception:
                                return f"<td style='padding:4px 5px;font-size:12px;background:{bg};'></td>"
                            if fv == 0:
                                return f"<td style='padding:4px 5px;font-size:12px;background:{bg};'></td>"
                            if is_ot:
                                s = minutes_to_hhmm(fv)
                            else:
                                s = f"{fv:,.0f}" if abs(fv - round(fv)) < 1e-6 else f"{fv:,.1f}"
                            if is_var:
                                cls = "pos" if fv > 0 else "neg"
                                return f"<td class='{cls}' style='padding:4px 5px;font-size:12px;background:{bg};'>{s}</td>"
                            return f"<td style='padding:4px 5px;font-size:12px;background:{bg};'>{s}</td>"
                        html.append(cell_bg(row[f"{m}_CY"], is_ot=is_ot, bg=bg))
                        html.append(cell_bg(row[f"{m}_LY"], is_ot=is_ot, bg=bg))
                        html.append(cell_bg(row[f"{m}_VAR"], is_var=True, is_ot=is_ot, bg=bg))
                    html.append("</tr>")
                html.append("</tbody></table></div>")

                st.markdown("".join(html), unsafe_allow_html=True)
                st.caption(
                    f"Headings: {h_cy}/{h_ly}/{h_var} | Crew REQ = MUSTERS×1.3 | OT = sum of CND OT / DRV OT (HH:MM)"
                )
                st.download_button(
                    "Download Product-wise CSV",
                    merged.to_csv(index=False).encode("utf-8"),
                    f"Schedules_Product_{f_month}.csv",
                    "text/csv",
                    key="dl9a",
                )


            # ========== TABLE 2: Depot-wise RTC/HIRE Summary (like your image TOTAL section) ==========
            
            # ========== TABLE 2: Depot-wise full summary (DO/SC/SO/NO) ==========
            st.markdown('<hr style="margin:4px 0;border:none;border-top:1px solid #e2e8f0;">', unsafe_allow_html=True)
            st.markdown(f"#### Depot Wise Summary of SCHs and Services — {f_month}")

            col_route = find_col(["ROUTEE", "ROUTE", "RouteName", "Route"])
            col_dtype = find_col(["D.TYPE", "DTYPE", "D.Type", "DutyType"])
            col_cond_mus = find_col(["COND MUSTERS", "CONDMUSTERS", "COND_MUSTERS", "ConductorMuster", "CND MUSTER", "COND MUSTER"])
            col_dri_mus = find_col(["DRI MUS", "DRIMUS", "DRI_MUS", "DriverMuster", "DRV MUSTER", "DRI MUSTERS", "DRI MUSTER"])
            # CND OT / DRV OT often have leading spaces in Excel headers e.g. " CND OT"
            col_cnd_ot = find_col(["CND OT", "CNDOT", "CND_OT", "ConductorOT", "CNDOT"])
            col_drv_ot = find_col(["DRV OT", "DRVOT", "DRV_OT", "DriverOT", "DRVOT"])
            if col_cnd_ot is None:
                for c in sched_raw.columns:
                    if str(c).strip().upper().replace(" ", "") in ("CNDOT", "CNDOUT"):
                        col_cnd_ot = c
                        break
            if col_drv_ot is None:
                for c in sched_raw.columns:
                    if str(c).strip().upper().replace(" ", "") in ("DRVOT", "DRVOUT"):
                        col_drv_ot = c
                        break
            col_muster3 = find_col(["3 DAYS MUSTER", "3DAYSMUSTER", "Muster3"])

            # ---- Table B extra filters: MHL/NMHL, ROUTE, PRODUCT ----
            st.markdown("##### Table B filters")
            bf1, bf2, bf3 = st.columns(3)
            with bf1:
                b_mhl_opts = ["ALL"] + sorted([x for x in sdf["_MHL"].dropna().unique() if x and str(x).lower() != "nan"])
                b_mhl = st.selectbox("MHL / NMHL", b_mhl_opts, key="sch_b_mhl")
            with bf2:
                col_route = find_col(["ROUTEE", "ROUTE", "RouteName", "Route"])
                if col_route:
                    sdf["_ROUTE"] = sdf[col_route].astype(str).str.strip()
                    # month-scoped route list
                    route_base = sdf[sdf["_MonthKey"] == f_month] if f_month != "ALL" else sdf
                    b_route_opts = ["ALL"] + sorted([x for x in route_base["_ROUTE"].dropna().unique() if str(x).strip() and str(x).strip().lower() != "nan"])
                else:
                    b_route_opts = ["ALL"]
                b_route = st.selectbox("ROUTE", b_route_opts, key="sch_b_route")
            with bf3:
                prod_base = sdf[sdf["_MonthKey"] == f_month] if f_month != "ALL" else sdf
                if b_mhl != "ALL":
                    prod_base = prod_base[prod_base["_MHL"] == b_mhl]
                if col_route and b_route != "ALL":
                    prod_base = prod_base[prod_base.get("_ROUTE", prod_base.columns[0]) == b_route] if "_ROUTE" in prod_base.columns else prod_base
                b_prod_opts = ["ALL"] + sorted([x for x in prod_base["_PRODUCT"].dropna().unique() if x and str(x).lower() != "nan"])
                b_product = st.selectbox("PRODUCT", b_prod_opts, key="sch_b_product")

            # Source data for depot summary: month + shared depot + table B filters
            dep_src = sdf[sdf["_MonthKey"] == f_month].copy() if f_month != "ALL" else sdf.copy()
            if f_depot not in ("ALL", "REGION"):
                dep_src = dep_src[dep_src["_DEPOT"] == f_depot]
            if b_mhl != "ALL":
                dep_src = dep_src[dep_src["_MHL"] == b_mhl]
            if col_route and b_route != "ALL" and "_ROUTE" in dep_src.columns:
                dep_src = dep_src[dep_src["_ROUTE"] == b_route]
            if b_product != "ALL":
                dep_src = dep_src[dep_src["_PRODUCT"] == b_product]

            if len(dep_src) == 0:
                st.warning("No depot-wise data.")
            else:
                # Normalize D.TYPE to DO/SC/SO/NO
                if col_dtype:
                    dep_src = dep_src.copy()
                    dep_src["_DTYPE"] = dep_src[col_dtype].astype(str).str.strip().str.upper()
                    # map common variants
                    dep_src["_DTYPE"] = dep_src["_DTYPE"].replace({
                        "D.O": "DO", "D/O": "DO", "DO.": "DO",
                        "S.C": "SC", "S/C": "SC",
                        "S.O": "SO", "S/O": "SO",
                        "N.O": "NO", "N/O": "NO",
                    })
                else:
                    dep_src["_DTYPE"] = "OTHER"

                def parse_time_to_minutes(v):
                    """Convert HH:MM / H:MM / datetime.time / timedelta to total minutes."""
                    if v is None or (isinstance(v, float) and pd.isna(v)):
                        return 0.0
                    if isinstance(v, (int, float)):
                        # already numeric - treat as hours if small, or minutes if large
                        return float(v)
                    # datetime.time
                    try:
                        import datetime as _dt
                        if isinstance(v, _dt.time):
                            return v.hour * 60 + v.minute + v.second / 60.0
                        if isinstance(v, _dt.timedelta):
                            return v.total_seconds() / 60.0
                        if isinstance(v, _dt.datetime):
                            return v.hour * 60 + v.minute
                    except Exception:
                        pass
                    s = str(v).strip()
                    if not s or s.lower() in ("nan", "nat", "none", "null"):
                        return 0.0
                    # formats: 38:42, 38:42:00, 1:05
                    if ":" in s:
                        parts = s.split(":")
                        try:
                            h = int(parts[0])
                            m = int(parts[1]) if len(parts) > 1 else 0
                            sec = int(parts[2]) if len(parts) > 2 else 0
                            return h * 60 + m + sec / 60.0
                        except Exception:
                            return 0.0
                    try:
                        return float(s)
                    except Exception:
                        return 0.0

                def minutes_to_hhmm(mins):
                    if mins is None or mins == 0:
                        return ""
                    mins = int(round(float(mins)))
                    h, m = divmod(abs(mins), 60)
                    sign = "-" if mins < 0 else ""
                    return f"{sign}{h}:{m:02d}"

                for c, alias in [
                    (col_cond_mus, "_COND_MUS"),
                    (col_dri_mus, "_DRI_MUS"),
                ]:
                    if c:
                        dep_src[alias] = pd.to_numeric(dep_src[c], errors="coerce").fillna(0)
                    else:
                        dep_src[alias] = 0
                # OT columns are time (HH:MM) → minutes
                if col_cnd_ot:
                    dep_src["_CND_OT"] = dep_src[col_cnd_ot].map(parse_time_to_minutes)
                else:
                    dep_src["_CND_OT"] = 0
                if col_drv_ot:
                    dep_src["_DRV_OT"] = dep_src[col_drv_ot].map(parse_time_to_minutes)
                else:
                    dep_src["_DRV_OT"] = 0

                DTYPES = ["DO", "SC", "SO", "NO"]

                def side_metrics(sub):
                    schs = sub["_SCH"].sum()
                    ser = sub[col_svc].nunique()
                    kms = sub["_KMS"].sum()
                    cnd_req = sub["_COND_MUS"].sum() * 1.3
                    drv_req = sub["_DRI_MUS"].sum() * 1.3
                    cnd_ot = sub["_CND_OT"].sum()
                    drv_ot = sub["_DRV_OT"].sum()
                    return {
                        "SCHS": schs, "SER": ser, "SCH_KMS": kms,
                        "CND_REQ": cnd_req, "DRV_REQ": drv_req,
                        "CND_OT": cnd_ot, "DRV_OT": drv_ot,
                    }

                def build_side_row(dep, side, base):
                    """Build one REGION/RTC_HIRE row. TOTAL SER/SCHS/KMS = RTC+HIRE (Excel style)."""
                    rec = {"REGION": dep, "RTC_HIRE": side}
                    if side == "TOTAL":
                        m_rtc = side_metrics(base[base["_RTC"] == "RTC"])
                        m_hire = side_metrics(base[base["_RTC"] == "HIRE"])
                        # Additive totals matching Excel
                        rec["T_SCHS"] = m_rtc["SCHS"] + m_hire["SCHS"]
                        rec["T_SER"] = m_rtc["SER"] + m_hire["SER"]
                        rec["T_KMS"] = m_rtc["SCH_KMS"] + m_hire["SCH_KMS"]
                        rec["T_CND_REQ"] = m_rtc["CND_REQ"] + m_hire["CND_REQ"]
                        rec["T_DRV_REQ"] = m_rtc["DRV_REQ"] + m_hire["DRV_REQ"]
                        for dt in DTYPES:
                            sub_r = base[(base["_RTC"] == "RTC") & (base["_DTYPE"] == dt)]
                            sub_h = base[(base["_RTC"] == "HIRE") & (base["_DTYPE"] == dt)]
                            mr, mh = side_metrics(sub_r), side_metrics(sub_h)
                            rec[f"{dt}_SCHS"] = mr["SCHS"] + mh["SCHS"]
                            rec[f"{dt}_SER"] = mr["SER"] + mh["SER"]
                            rec[f"{dt}_KMS"] = mr["SCH_KMS"] + mh["SCH_KMS"]
                            rec[f"{dt}_CND_REQ"] = mr["CND_REQ"] + mh["CND_REQ"]
                            rec[f"{dt}_DRV_REQ"] = mr["DRV_REQ"] + mh["DRV_REQ"]
                            rec[f"{dt}_CND_OT"] = mr["CND_OT"] + mh["CND_OT"]
                            rec[f"{dt}_DRV_OT"] = mr["DRV_OT"] + mh["DRV_OT"]
                    else:
                        sub = base[base["_RTC"] == side]
                        tot_m = side_metrics(sub)
                        rec["T_SCHS"] = tot_m["SCHS"]
                        rec["T_SER"] = tot_m["SER"]
                        rec["T_KMS"] = tot_m["SCH_KMS"]
                        rec["T_CND_REQ"] = tot_m["CND_REQ"]
                        rec["T_DRV_REQ"] = tot_m["DRV_REQ"]
                        for dt in DTYPES:
                            dsub = sub[sub["_DTYPE"] == dt]
                            m = side_metrics(dsub)
                            rec[f"{dt}_SCHS"] = m["SCHS"]
                            rec[f"{dt}_SER"] = m["SER"]
                            rec[f"{dt}_KMS"] = m["SCH_KMS"]
                            rec[f"{dt}_CND_REQ"] = m["CND_REQ"]
                            rec[f"{dt}_DRV_REQ"] = m["DRV_REQ"]
                            rec[f"{dt}_CND_OT"] = m["CND_OT"]
                            rec[f"{dt}_DRV_OT"] = m["DRV_OT"]
                    return rec

                def fmt_num(v, is_time=False):
                    try:
                        fv = float(v)
                    except Exception:
                        return ""
                    if fv == 0:
                        return ""
                    if is_time:
                        # if stored as hours decimal or minutes - show as number for now
                        return f"{fv:,.0f}" if abs(fv - round(fv)) < 1e-6 else f"{fv:,.2f}"
                    return f"{fv:,.0f}" if abs(fv - round(fv)) < 1e-6 else f"{fv:,.1f}"

                rows_out = []
                depots = sorted([str(d).strip() for d in dep_src["_DEPOT"].dropna().unique() if str(d).strip() and str(d).strip().lower() != "nan"])
                for dep in depots:
                    dgrp = dep_src[dep_src["_DEPOT"] == dep]
                    for side in ["RTC", "HIRE", "TOTAL"]:
                        rows_out.append(build_side_row(dep, side, dgrp))

                # REGION totals = SUM of depot rows (Excel: REGION SER = sum of depot SERs)
                import copy
                depot_rows = [r for r in rows_out if r["REGION"] != "REGION"]
                for side in ["RTC", "HIRE", "TOTAL"]:
                    side_rows = [r for r in depot_rows if r["RTC_HIRE"] == side]
                    rec = {"REGION": "REGION", "RTC_HIRE": side}
                    if not side_rows:
                        rows_out.append(build_side_row("REGION", side, dep_src))
                        continue
                    # sum all numeric fields
                    keys = [k for k in side_rows[0].keys() if k not in ("REGION", "RTC_HIRE")]
                    for k in keys:
                        rec[k] = sum(float(r.get(k, 0) or 0) for r in side_rows)
                    rows_out.append(rec)

                dsum = pd.DataFrame(rows_out)

                # HTML - compact
                html2 = ['<div class="table-scroll"><table class="excel-table" style="table-layout:auto;font-size:9px;">']
                # header row 1
                html2.append("<tr>")
                html2.append('<th rowspan="2" style="background:#0f172a;color:white;padding:6px;">REGION</th>')
                html2.append('<th rowspan="2" style="background:#0f172a;color:white;padding:6px;">RTC/HIRE</th>')
                html2.append('<th colspan="3" style="background:#b91c1c;color:white;padding:6px;">DO</th>')
                html2.append('<th colspan="7" style="background:#7c3aed;color:white;padding:6px;">SC</th>')
                html2.append('<th colspan="5" style="background:#15803d;color:white;padding:6px;">SO</th>')
                html2.append('<th colspan="5" style="background:#0369a1;color:white;padding:6px;">NO</th>')
                html2.append('<th colspan="3" style="background:#c2410c;color:white;padding:6px;">TOTAL</th>')
                html2.append('<th colspan="2" style="background:#166534;color:white;padding:6px;">REQMT CREW</th>')
                html2.append("</tr>")
                # header row 2
                html2.append("<tr>")
                # DO: SCHS SER SCH KMS
                for h in ["SCHS", "SER", "SCH KMS"]:
                    html2.append(f'<th class="header-sub">{h}</th>')
                # SC: SCHS SER SCH KMS CND REQ DRV REQ CNDOT DRVOT
                for h in ["SCHS", "SER", "SCH KMS", "CND REQ", "DRV REQ", "CNDOT", "DRVOT"]:
                    html2.append(f'<th class="header-sub">{h}</th>')
                # SO
                for h in ["SCHS", "SER", "SCH KMS", "CND REQ", "DRV REQ"]:
                    html2.append(f'<th class="header-sub">{h}</th>')
                # NO
                for h in ["SCHS", "SER", "SCH KMS", "CND REQ", "DRV REQ"]:
                    html2.append(f'<th class="header-sub">{h}</th>')
                # TOTAL
                for h in ["SCHS", "SER", "SCH KMS"]:
                    html2.append(f'<th class="header-sub">{h}</th>')
                # REQMT
                for h in ["COND", "DRI"]:
                    html2.append(f'<th class="header-sub">{h}</th>')
                html2.append("</tr>")

                for _, r in dsum.iterrows():
                    is_tot = str(r["RTC_HIRE"]) == "TOTAL" or str(r["REGION"]) == "REGION"
                    style = "font-weight:bold;background:#e2efda;" if is_tot else ""
                    html2.append(f'<tr style="{style}">')
                    html2.append(f'<td>{r["REGION"]}</td>')
                    html2.append(f'<td>{r["RTC_HIRE"]}</td>')
                    def td(v, bg, is_ot=False):
                        val = minutes_to_hhmm(v) if is_ot else fmt_num(v)
                        return f'<td style="background:{bg};padding:4px;font-size:11px;">{val}</td>'
                    # DO
                    for k in ["DO_SCHS", "DO_SER", "DO_KMS"]:
                        html2.append(td(r.get(k, 0), "#fef2f2"))
                    # SC
                    for k in ["SC_SCHS", "SC_SER", "SC_KMS"]:
                        html2.append(td(r.get(k, 0), "#f5f3ff"))
                    for k in ["SC_CND_REQ", "SC_DRV_REQ"]:
                        html2.append(td(r.get(k, 0), "#eff6ff"))
                    html2.append(td(r.get("SC_CND_OT", 0), "#fff7ed", is_ot=True))
                    html2.append(td(r.get("SC_DRV_OT", 0), "#fff7ed", is_ot=True))
                    # SO
                    for k in ["SO_SCHS", "SO_SER", "SO_KMS"]:
                        html2.append(td(r.get(k, 0), "#f0fdf4"))
                    for k in ["SO_CND_REQ", "SO_DRV_REQ"]:
                        html2.append(td(r.get(k, 0), "#eff6ff"))
                    # NO
                    for k in ["NO_SCHS", "NO_SER", "NO_KMS"]:
                        html2.append(td(r.get(k, 0), "#fefce8"))
                    for k in ["NO_CND_REQ", "NO_DRV_REQ"]:
                        html2.append(td(r.get(k, 0), "#eff6ff"))
                    # TOTAL
                    for k in ["T_SCHS", "T_SER", "T_KMS"]:
                        html2.append(td(r.get(k, 0), "#e2e8f0"))
                    # REQMT crew
                    html2.append(td(r.get("T_CND_REQ", 0), "#dbeafe"))
                    html2.append(td(r.get("T_DRV_REQ", 0), "#dbeafe"))
                    html2.append("</tr>")
                html2.append("</table></div>")
                st.markdown("".join(html2), unsafe_allow_html=True)
                st.download_button(
                    "Download Depot-wise CSV",
                    dsum.to_csv(index=False).encode("utf-8"),
                    f"Schedules_DepotDetail_{f_month}.csv",
                    "text/csv",
                    key="dl9b",
                )


            # ========== TABLE 3: Summary of Operation ==========
            st.markdown('<hr style="margin:4px 0;border:none;border-top:1px solid #e2e8f0;">', unsafe_allow_html=True)
            st.markdown(f"#### Summary of Operation — {f_month}")

            # Reuse Table B filtered source (dep_src) if available, else rebuild
            try:
                op_src = dep_src.copy()
            except NameError:
                op_src = sdf[sdf["_MonthKey"] == f_month].copy() if f_month != "ALL" else sdf.copy()
                if f_depot not in ("ALL", "REGION"):
                    op_src = op_src[op_src["_DEPOT"] == f_depot]

            col_route = find_col(["ROUTEE", "ROUTE", "RouteName", "Route"])
            if col_route and "_ROUTE" not in op_src.columns:
                op_src["_ROUTE"] = op_src[col_route].astype(str).str.strip()
            elif "_ROUTE" not in op_src.columns:
                op_src["_ROUTE"] = ""

            _dtype_col = find_col(["D.TYPE", "DTYPE", "D.Type", "DutyType"])
            if _dtype_col:
                op_src["_DTYPE"] = op_src[_dtype_col].astype(str).str.strip().str.upper()
                op_src["_DTYPE"] = op_src["_DTYPE"].replace({
                    "D.O": "DO", "D/O": "DO", "DO.": "DO",
                    "S.C": "SC", "S/C": "SC",
                    "S.O": "SO", "S/O": "SO",
                    "N.O": "NO", "N/O": "NO",
                })
            else:
                op_src["_DTYPE"] = ""

            DTYPES3 = ["DO", "NO", "SC", "SO"]

            # Crew + OT columns for Table C
            def _parse_ot_min(v):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return 0.0
                if isinstance(v, (int, float)):
                    return float(v)
                try:
                    import datetime as _dt
                    if isinstance(v, _dt.time):
                        return v.hour * 60 + v.minute
                    if isinstance(v, _dt.timedelta):
                        return v.total_seconds() / 60.0
                    if isinstance(v, _dt.datetime):
                        return v.hour * 60 + v.minute
                except Exception:
                    pass
                s = str(v).strip()
                if not s or s.lower() in ("nan", "nat", "none", "null"):
                    return 0.0
                if ":" in s:
                    parts = s.split(":")
                    try:
                        return int(parts[0]) * 60 + int(parts[1])
                    except Exception:
                        return 0.0
                try:
                    return float(s)
                except Exception:
                    return 0.0

            def _min_to_hhmm(mins):
                try:
                    mins = int(round(float(mins)))
                except Exception:
                    return ""
                if mins == 0:
                    return ""
                h, m = divmod(abs(mins), 60)
                return f"{'-' if mins < 0 else ''}{h}:{m:02d}"

            c_cond = find_col(["COND MUSTERS", "CONDMUSTERS", "COND MUSTER"])
            c_dri = find_col(["DRI MUS", "DRIMUS", "DRI MUSTER", "DRI MUSTERS"])
            c_cnd_ot = find_col(["CND OT", "CNDOT"])
            c_drv_ot = find_col(["DRV OT", "DRVOT"])
            if c_cnd_ot is None:
                for c in op_src.columns:
                    if str(c).strip().upper().replace(" ", "") == "CNDOT":
                        c_cnd_ot = c
                        break
            if c_drv_ot is None:
                for c in op_src.columns:
                    if str(c).strip().upper().replace(" ", "") == "DRVOT":
                        c_drv_ot = c
                        break
            op_src = op_src.copy()
            op_src["_COND_MUS"] = pd.to_numeric(op_src[c_cond], errors="coerce").fillna(0) if c_cond else 0
            op_src["_DRI_MUS"] = pd.to_numeric(op_src[c_dri], errors="coerce").fillna(0) if c_dri else 0
            op_src["_CND_OT"] = op_src[c_cnd_ot].map(_parse_ot_min) if c_cnd_ot else 0
            op_src["_DRV_OT"] = op_src[c_drv_ot].map(_parse_ot_min) if c_drv_ot else 0

            if len(op_src) == 0:
                st.warning("No data for Summary of Operation.")
            else:
                rows3 = []
                sno = 1
                group_cols = ["_DEPOT", "_ROUTE", "_PRODUCT"]
                for keys, grp in op_src.groupby(group_cols, dropna=False):
                    dep, route, prod = keys
                    svcs = sorted({str(x).strip() for x in grp[col_svc].dropna().unique() if str(x).strip() and str(x).lower() != "nan"})
                    svc_str = ",".join(svcs) + ("," if svcs else "")
                    rec = {
                        "SNo": sno,
                        "DEPOT": dep,
                        "ROUTE": route,
                        "PRODUCT": prod,
                        "SERVICE_NUMBERS": svc_str,
                    }
                    for dt in DTYPES3:
                        sub = grp[grp["_DTYPE"] == dt]
                        rec[f"SCH_{dt}"] = sub["_SCH"].sum()
                        rec[f"SER_{dt}"] = sub[col_svc].nunique()
                        rec[f"KMS_{dt}"] = sub["_KMS"].sum()
                    rec["SCH_TOTAL"] = sum(rec[f"SCH_{dt}"] for dt in DTYPES3)
                    rec["SER_TOTAL"] = sum(rec[f"SER_{dt}"] for dt in DTYPES3)
                    rec["KMS_TOTAL"] = sum(rec[f"KMS_{dt}"] for dt in DTYPES3)
                    # Crew requirement (musters x 1.3) and Schedule OT
                    rec["CREW_COND"] = grp["_COND_MUS"].sum() * 1.3
                    rec["CREW_DRI"] = grp["_DRI_MUS"].sum() * 1.3
                    rec["OT_COND"] = grp["_CND_OT"].sum()
                    rec["OT_DRI"] = grp["_DRV_OT"].sum()
                    rows3.append(rec)
                    sno += 1

                df3 = pd.DataFrame(rows3)

                # TOTAL row
                if len(df3) > 0:
                    tot = {"SNo": "", "DEPOT": "TOTAL", "ROUTE": "", "PRODUCT": "", "SERVICE_NUMBERS": ""}
                    for dt in DTYPES3:
                        tot[f"SCH_{dt}"] = df3[f"SCH_{dt}"].sum()
                        tot[f"SER_{dt}"] = df3[f"SER_{dt}"].sum()
                        tot[f"KMS_{dt}"] = df3[f"KMS_{dt}"].sum()
                    tot["SCH_TOTAL"] = df3["SCH_TOTAL"].sum()
                    tot["SER_TOTAL"] = df3["SER_TOTAL"].sum()
                    tot["KMS_TOTAL"] = df3["KMS_TOTAL"].sum()
                    tot["CREW_COND"] = df3["CREW_COND"].sum()
                    tot["CREW_DRI"] = df3["CREW_DRI"].sum()
                    tot["OT_COND"] = df3["OT_COND"].sum()
                    tot["OT_DRI"] = df3["OT_DRI"].sum()
                    df3 = pd.concat([df3, pd.DataFrame([tot])], ignore_index=True)

                def n3(v):
                    try:
                        fv = float(v)
                        return "" if fv == 0 else (f"{fv:,.0f}" if abs(fv - round(fv)) < 1e-6 else f"{fv:,.1f}")
                    except Exception:
                        return ""

                html3 = ['<div class="op-wrap"><table class="op-table"><thead>']
                html3.append("<tr>")
                html3.append('<th rowspan="2" style="position:sticky;left:0;top:0;z-index:3;background:#0f172a;color:white;padding:8px 6px;font-size:13px;min-width:40px;">S.No</th>')
                html3.append('<th rowspan="2" style="position:sticky;left:40px;top:0;z-index:3;background:#0f172a;color:white;padding:8px 6px;font-size:13px;min-width:60px;">DEPOT</th>')
                html3.append('<th rowspan="2" style="position:sticky;left:100px;top:0;z-index:3;background:#0f172a;color:white;padding:8px 6px;font-size:13px;min-width:80px;">ROUTE NO.</th>')
                html3.append('<th rowspan="2" style="position:sticky;left:180px;top:0;z-index:3;background:#0f172a;color:white;padding:8px 6px;font-size:13px;min-width:70px;">PRODUCT</th>')
                html3.append('<th rowspan="2" style="position:sticky;left:250px;top:0;z-index:3;background:#0f172a;color:white;padding:8px 6px;font-size:13px;min-width:190px;width:190px;">SERVICE NUMBERS</th>')
                html3.append('<th colspan="5" style="position:sticky;top:0;z-index:4;background:#b91c1c;color:white;padding:8px;font-size:13px;">NO OF SCHEDULES</th>')
                html3.append('<th colspan="5" style="position:sticky;top:0;z-index:4;background:#7c3aed;color:white;padding:8px;font-size:13px;">SERVICES</th>')
                html3.append('<th colspan="5" style="position:sticky;top:0;z-index:4;background:#15803d;color:white;padding:8px;font-size:13px;">SCHEDULE KILOMETERS</th>')
                html3.append('<th colspan="2" style="position:sticky;top:0;z-index:4;background:#1d4ed8;color:white;padding:8px;font-size:13px;">CREW REQUIREMENT</th>')
                html3.append('<th colspan="2" style="position:sticky;top:0;z-index:4;background:#c2410c;color:white;padding:8px;font-size:13px;">SCH OVER TIME</th>')
                html3.append("</tr><tr>")
                for bg in ("#fecaca", "#ddd6fe", "#bbf7d0"):
                    for h in ["DO", "NO", "SC", "SO", "TOTAL"]:
                        html3.append(f'<th style="position:sticky;top:36px;z-index:4;background:{bg};color:#0f172a;padding:6px;font-size:12px;">{h}</th>')
                for h in ["COND", "DRI"]:
                    html3.append(f'<th style="position:sticky;top:36px;z-index:4;background:#dbeafe;color:#0f172a;padding:6px;font-size:12px;">{h}</th>')
                for h in ["COND", "DRI"]:
                    html3.append(f'<th style="position:sticky;top:36px;z-index:4;background:#ffedd5;color:#0f172a;padding:6px;font-size:12px;">{h}</th>')
                html3.append("</tr></thead><tbody>")

                sticky_base = 'position:sticky;z-index:2;padding:6px;font-size:13px;'
                lefts = [0, 45, 110, 200, 280]  # approx sticky offsets for 5 cols
                for _, r in df3.iterrows():
                    is_tot = str(r["DEPOT"]) == "TOTAL"
                    style = "font-weight:bold;background:#e2efda;" if is_tot else ""
                    html3.append(f'<tr style="{style}">')
                    bg0 = "#e2efda" if is_tot else "#e0f2fe"
                    html3.append(f'<td style="{sticky_base}left:0;background:{bg0};min-width:40px;">{r["SNo"]}</td>')
                    html3.append(f'<td style="{sticky_base}left:40px;background:{bg0};min-width:60px;">{r["DEPOT"]}</td>')
                    html3.append(f'<td style="{sticky_base}left:100px;background:{bg0};min-width:80px;">{r["ROUTE"]}</td>')
                    html3.append(f'<td style="{sticky_base}left:180px;background:{bg0};min-width:70px;">{r["PRODUCT"]}</td>')
                    html3.append(f'<td style="{sticky_base}left:250px;background:{bg0};min-width:70px;max-width:90px;text-align:left;white-space:normal;font-size:12px;">{r["SERVICE_NUMBERS"]}</td>')
                    for dt in DTYPES3:
                        html3.append(f'<td style="padding:6px;font-size:13px;background:#fef2f2;">{n3(r[f"SCH_{dt}"])}</td>')
                    html3.append(f'<td style="padding:6px;font-size:13px;background:#fee2e2;font-weight:600;">{n3(r["SCH_TOTAL"])}</td>')
                    for dt in DTYPES3:
                        html3.append(f'<td style="padding:6px;font-size:13px;background:#f5f3ff;">{n3(r[f"SER_{dt}"])}</td>')
                    html3.append(f'<td style="padding:6px;font-size:13px;background:#ede9fe;font-weight:600;">{n3(r["SER_TOTAL"])}</td>')
                    for dt in DTYPES3:
                        html3.append(f'<td style="padding:6px;font-size:13px;background:#f0fdf4;">{n3(r[f"KMS_{dt}"])}</td>')
                    html3.append(f'<td style="padding:6px;font-size:13px;background:#dcfce7;font-weight:600;">{n3(r["KMS_TOTAL"])}</td>')
                    html3.append(f'<td style="padding:6px;font-size:13px;background:#eff6ff;">{n3(r.get("CREW_COND", 0))}</td>')
                    html3.append(f'<td style="padding:6px;font-size:13px;background:#eff6ff;">{n3(r.get("CREW_DRI", 0))}</td>')
                    html3.append(f'<td style="padding:6px;font-size:13px;background:#fff7ed;">{_min_to_hhmm(r.get("OT_COND", 0))}</td>')
                    html3.append(f'<td style="padding:6px;font-size:13px;background:#fff7ed;">{_min_to_hhmm(r.get("OT_DRI", 0))}</td>')
                    html3.append("</tr>")
                html3.append("</tbody></table></div>")
                st.markdown("".join(html3), unsafe_allow_html=True)
                st.download_button(
                    "Download Summary of Operation CSV",
                    df3.to_csv(index=False).encode("utf-8"),
                    f"Summary_Operation_{f_month}.csv",
                    "text/csv",
                    key="dl9c",
                )



st.caption("Cascading filters • Weighted EPK • Self-hosted on your PC")