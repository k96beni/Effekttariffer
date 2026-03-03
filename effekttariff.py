import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Elnätsavgift Kalkylator", page_icon="⚡", layout="wide")

# --- DATA LOADING ---
@st.cache_data
def load_data():
    base = os.path.dirname(os.path.abspath(__file__))

    # Load both files
    ef_raw = pd.read_excel(os.path.join(base, "Effektkunder-elnätsföretagens-elnätsavgifter.xlsx"), header=None, sheet_name=0)
    sf_raw = pd.read_excel(os.path.join(base, "Större-säkringskunder-elnätsföretagens-elnätsavgifter.xlsx"), header=None, sheet_name=0)

    return ef_raw, sf_raw

@st.cache_data
def parse_companies(df):
    """Extract company list from row 4 onwards."""
    companies = {}
    for i in range(4, len(df)):
        code = df.iloc[i, 0]
        name = df.iloc[i, 1]
        if pd.notna(name) and str(name).strip():
            companies[str(name).strip()] = {"row": i, "code": str(code).strip() if pd.notna(code) else ""}
    return companies

@st.cache_data
def parse_sakring_tariffs(df, row_idx):
    """Parse tariff data for Större säkringskunder."""
    categories = {
        "35A, 30 000 kWh/år": {"start_col": 3, "cols": {
            "myndighetsavgift": 3, "fast_avgift": 8, "rorlig1": 13, "rorlig2": 18, "total": 23
        }},
        "63A, 50 000 kWh/år": {"start_col": 28, "cols": {
            "myndighetsavgift": 28, "fast_avgift": 33, "rorlig1": 38, "rorlig2": 43, "total": 48
        }},
        "80A, 80 000 kWh/år": {"start_col": 53, "cols": {
            "myndighetsavgift": 53, "fast_avgift": 58, "rorlig1": 63, "rorlig2": 68, "total": 73
        }},
        "100A, 100 000 kWh/år": {"start_col": 78, "cols": {
            "myndighetsavgift": 78, "fast_avgift": 83, "rorlig1": 88, "rorlig2": 93, "total": 98
        }},
        "125A, 125 000 kWh/år": {"start_col": 103, "cols": {
            "myndighetsavgift": 103, "fast_avgift": 108, "rorlig1": 113, "rorlig2": 118, "total": 123
        }},
        "160A, 190 000 kWh/år": {"start_col": 128, "cols": {
            "myndighetsavgift": 128, "fast_avgift": 133, "rorlig1": 138, "rorlig2": 143, "total": 148
        }},
    }

    result = {}
    for cat_name, cat_info in categories.items():
        cols = cat_info["cols"]
        data = {}
        for field, col in cols.items():
            val = df.iloc[row_idx, col]
            data[field] = float(val) if pd.notna(val) else None
        # Only include if at least fast_avgift exists
        if data.get("fast_avgift") is not None:
            result[cat_name] = data
    return result

@st.cache_data
def parse_effekt_tariffs(df, row_idx):
    """Parse tariff data for Effektkunder."""
    categories = {
        "100 kW, 350 MWh/år": {"cols": {
            "myndighetsavgift": 3, "fast_avgift": 8, "abonnerad_effekt": 13,
            "hogbelast_effekt": 18, "antal_matvarden": 23, "intervall_red": 28,
            "vinter_hog": 33, "vinter_lag": 38, "var_host_hog": 43,
            "var_host_lag": 48, "sommar_hog": 53, "sommar_lag": 58, "total": 63
        }},
        "1 MW, 5 GWh/år": {"cols": {
            "myndighetsavgift": 68, "fast_avgift": 73, "abonnerad_effekt": 78,
            "hogbelast_effekt": 83, "antal_matvarden": 88, "intervall_red": 93,
            "vinter_hog": 98, "vinter_lag": 103, "var_host_hog": 108,
            "var_host_lag": 113, "sommar_hog": 118, "sommar_lag": 123, "total": 128
        }},
        "20 MW, 140 GWh/år": {"cols": {
            "myndighetsavgift": 133, "fast_avgift": 138, "abonnerad_effekt": 143,
            "hogbelast_effekt": 148, "antal_matvarden": 153, "intervall_red": 158,
            "vinter_hog": 163, "vinter_lag": 168, "var_host_hog": 173,
            "var_host_lag": 178, "sommar_hog": 183, "sommar_lag": 188, "total": 193
        }},
    }

    result = {}
    for cat_name, cat_info in categories.items():
        cols = cat_info["cols"]
        data = {}
        for field, col in cols.items():
            val = df.iloc[row_idx, col]
            data[field] = float(val) if pd.notna(val) else None
        if data.get("fast_avgift") is not None or data.get("hogbelast_effekt") is not None:
            result[cat_name] = data
    return result

def calc_sakring(tariff, kwh_year):
    """Calculate annual cost for säkringskund."""
    mynd = tariff.get("myndighetsavgift") or 0
    fast = tariff.get("fast_avgift") or 0
    r1 = tariff.get("rorlig1") or 0  # öre/kWh
    r2 = tariff.get("rorlig2")
    # Use average of r1 and r2 if both exist, otherwise just r1
    if r2 is not None and r2 > 0:
        rorlig = (r1 + r2) / 2
    else:
        rorlig = r1
    rorlig_kr = kwh_year * rorlig / 100
    total = mynd + fast + rorlig_kr
    return {
        "myndighetsavgift": mynd,
        "fast_avgift": fast,
        "rorlig_ore_kwh": rorlig,
        "rorlig_kr": rorlig_kr,
        "total": total
    }

def calc_effekt(tariff, kwh_year, max_kw):
    """Calculate annual cost for effektkund."""
    mynd = tariff.get("myndighetsavgift") or 0
    fast = tariff.get("fast_avgift") or 0
    abon = tariff.get("abonnerad_effekt") or 0
    hogbelast = tariff.get("hogbelast_effekt") or 0

    # Rörlig: weighted average across seasons
    # Approximate: vinter 4 mån, vår/höst 4 mån, sommar 4 mån
    season_weights = {"vinter": 4/12, "var_host": 4/12, "sommar": 4/12}
    rorlig_avg = 0
    for season, weight in season_weights.items():
        hog = tariff.get(f"{season}_hog") or 0
        lag = tariff.get(f"{season}_lag") or 0
        # Assume 50/50 hög/låg within each season
        rorlig_avg += weight * (hog + lag) / 2

    abon_kr = max_kw * abon * 12  # monthly
    hogbelast_kr = max_kw * hogbelast  # annual (based on antal_matvarden)
    rorlig_kr = kwh_year * rorlig_avg / 100
    total = mynd + fast + abon_kr + hogbelast_kr + rorlig_kr

    return {
        "myndighetsavgift": mynd,
        "fast_avgift": fast,
        "abonnerad_effekt_kr_kw": abon,
        "abonnerad_effekt_kr": abon_kr,
        "hogbelast_effekt_kr_kw": hogbelast,
        "hogbelast_effekt_kr": hogbelast_kr,
        "rorlig_ore_kwh": rorlig_avg,
        "rorlig_kr": rorlig_kr,
        "total": total
    }

# --- UI ---
st.title("⚡ Elnätsavgift Kalkylator")
st.caption("Baserad på Energimarknadsinspektionens tariffdata (2025)")

try:
    ef_raw, sf_raw = load_data()
except Exception as e:
    st.error(f"Kunde inte ladda datafiler: {e}")
    st.stop()

# Get companies from both files
companies_sf = parse_companies(sf_raw)
companies_ef = parse_companies(ef_raw)

# Merge — use sf as primary (more companies)
all_companies = {}
for name, info in companies_sf.items():
    all_companies[name] = {"sf_row": info["row"], "ef_row": None, "code": info["code"]}
for name, info in companies_ef.items():
    if name in all_companies:
        all_companies[name]["ef_row"] = info["row"]
    else:
        all_companies[name] = {"sf_row": None, "ef_row": info["row"], "code": info["code"]}

company_names = sorted(all_companies.keys())

# --- SIDEBAR ---
st.sidebar.header("Inställningar")

selected_company = st.sidebar.selectbox(
    "Välj elnätsföretag",
    company_names,
    index=company_names.index("Kraftringen Nät AB") if "Kraftringen Nät AB" in company_names else 0
)

st.sidebar.divider()
st.sidebar.subheader("Förbrukningsdata")

num_outlets = st.sidebar.number_input("Antal ladduttag", min_value=1, max_value=100, value=4)
kwh_per_outlet_month = st.sidebar.number_input("kWh per uttag per månad", min_value=0, max_value=10000, value=223)
max_power_kw = st.sidebar.number_input("Max effekt (kW/mån)", min_value=1, max_value=20000, value=15)

total_kwh_year = num_outlets * kwh_per_outlet_month * 12
total_kwh_month = num_outlets * kwh_per_outlet_month

st.sidebar.divider()
st.sidebar.metric("Total årsförbrukning", f"{total_kwh_year:,.0f} kWh/år")
st.sidebar.metric("Total månadsförbrukning", f"{total_kwh_month:,.0f} kWh/mån")

# --- MAIN CONTENT ---
company_info = all_companies[selected_company]

col1, col2 = st.columns(2)

# --- Säkringstariff ---
with col1:
    st.subheader("🔌 Säkringstariff")
    if company_info["sf_row"] is not None:
        sakring_tariffs = parse_sakring_tariffs(sf_raw, company_info["sf_row"])
        if sakring_tariffs:
            selected_sakring = st.selectbox(
                "Säkringskategori",
                list(sakring_tariffs.keys()),
                key="sakring_cat"
            )
            tariff = sakring_tariffs[selected_sakring]
            result = calc_sakring(tariff, total_kwh_year)

            st.markdown("**Tariffkomponenter (2025)**")
            tariff_df = pd.DataFrame({
                "Komponent": ["Myndighetsavgift", "Fast avgift", f"Rörlig ({result['rorlig_ore_kwh']:.2f} öre/kWh)"],
                "kr/år": [result["myndighetsavgift"], result["fast_avgift"], result["rorlig_kr"]]
            })
            tariff_df["kr/år"] = tariff_df["kr/år"].map(lambda x: f"{x:,.0f}")
            st.dataframe(tariff_df, use_container_width=True, hide_index=True)

            st.metric("Total elnätsavgift (exkl. moms)", f"{result['total']:,.0f} kr/år")

            subcol1, subcol2 = st.columns(2)
            subcol1.metric("Per månad", f"{result['total']/12:,.0f} kr")
            subcol2.metric("Per kWh", f"{result['total']/total_kwh_year*100:.1f} öre")
        else:
            st.warning("Ingen säkringstariffdata finns för detta företag.")
    else:
        st.warning("Ingen säkringstariffdata finns för detta företag.")

# --- Effekttariff ---
with col2:
    st.subheader("⚡ Effekttariff")
    if company_info["ef_row"] is not None:
        effekt_tariffs = parse_effekt_tariffs(ef_raw, company_info["ef_row"])
        if effekt_tariffs:
            selected_effekt = st.selectbox(
                "Effektkategori",
                list(effekt_tariffs.keys()),
                key="effekt_cat"
            )
            tariff_e = effekt_tariffs[selected_effekt]
            result_e = calc_effekt(tariff_e, total_kwh_year, max_power_kw)

            st.markdown("**Tariffkomponenter (2025)**")
            tariff_e_df = pd.DataFrame({
                "Komponent": [
                    "Myndighetsavgift",
                    "Fast avgift",
                    f"Abonnerad effekt ({result_e['abonnerad_effekt_kr_kw']:.0f} kr/kW × {max_power_kw} kW)",
                    f"Högbelasteffekt ({result_e['hogbelast_effekt_kr_kw']:.0f} kr/kW × {max_power_kw} kW)",
                    f"Rörlig ({result_e['rorlig_ore_kwh']:.2f} öre/kWh)"
                ],
                "kr/år": [
                    result_e["myndighetsavgift"],
                    result_e["fast_avgift"],
                    result_e["abonnerad_effekt_kr"],
                    result_e["hogbelast_effekt_kr"],
                    result_e["rorlig_kr"]
                ]
            })
            tariff_e_df["kr/år"] = tariff_e_df["kr/år"].map(lambda x: f"{x:,.0f}")
            st.dataframe(tariff_e_df, use_container_width=True, hide_index=True)

            st.metric("Total elnätsavgift (exkl. moms)", f"{result_e['total']:,.0f} kr/år")

            subcol1, subcol2 = st.columns(2)
            subcol1.metric("Per månad", f"{result_e['total']/12:,.0f} kr")
            subcol2.metric("Per kWh", f"{result_e['total']/total_kwh_year*100:.1f} öre")
        else:
            st.warning("Ingen effekttariffdata finns för detta företag.")
    else:
        st.warning("Ingen effekttariffdata finns för detta företag.")

# --- Jämförelse ---
st.divider()
st.subheader("📊 Jämförelse")

if company_info["sf_row"] is not None and company_info["ef_row"] is not None:
    sakring_tariffs = parse_sakring_tariffs(sf_raw, company_info["sf_row"])
    effekt_tariffs = parse_effekt_tariffs(ef_raw, company_info["ef_row"])

    comparison_data = []
    for cat, tariff in sakring_tariffs.items():
        r = calc_sakring(tariff, total_kwh_year)
        comparison_data.append({
            "Tarifftyp": "Säkring",
            "Kategori": cat,
            "Total kr/år": r["total"],
            "kr/mån": r["total"] / 12,
            "öre/kWh": r["total"] / total_kwh_year * 100 if total_kwh_year > 0 else 0
        })
    for cat, tariff in effekt_tariffs.items():
        r = calc_effekt(tariff, total_kwh_year, max_power_kw)
        comparison_data.append({
            "Tarifftyp": "Effekt",
            "Kategori": cat,
            "Total kr/år": r["total"],
            "kr/mån": r["total"] / 12,
            "öre/kWh": r["total"] / total_kwh_year * 100 if total_kwh_year > 0 else 0
        })

    if comparison_data:
        comp_df = pd.DataFrame(comparison_data).sort_values("Total kr/år")

        comp_df["Total kr/år"] = comp_df["Total kr/år"].map(lambda x: f"{x:,.0f}")
        comp_df["kr/mån"] = comp_df["kr/mån"].map(lambda x: f"{x:,.0f}")
        comp_df["öre/kWh"] = comp_df["öre/kWh"].map(lambda x: f"{x:.1f}")
        st.dataframe(comp_df, use_container_width=True, hide_index=True)

        cheapest = comp_df.iloc[0]
        st.success(f"Billigaste alternativ: **{cheapest['Tarifftyp']} — {cheapest['Kategori']}** = {cheapest['Total kr/år']} kr/år ({cheapest['öre/kWh']} öre/kWh)")

# --- Sensitivity: Effekt vs förbrukning ---
st.divider()
st.subheader("📈 Känslighetsanalys: Effekttariff vid varierande effekt")

if company_info["ef_row"] is not None:
    effekt_tariffs = parse_effekt_tariffs(ef_raw, company_info["ef_row"])
    if effekt_tariffs:
        first_cat = list(effekt_tariffs.keys())[0]
        tariff_sens = effekt_tariffs[first_cat]

        kw_range = list(range(5, min(max_power_kw * 4 + 1, 201), 5))
        sens_data = []
        for kw in kw_range:
            r = calc_effekt(tariff_sens, total_kwh_year, kw)
            sens_data.append({"Max effekt (kW)": kw, "Total kr/år": r["total"], "öre/kWh": r["total"] / total_kwh_year * 100 if total_kwh_year > 0 else 0})

        sens_df = pd.DataFrame(sens_data)
        st.line_chart(sens_df.set_index("Max effekt (kW)")["Total kr/år"], use_container_width=True)
        st.caption(f"Effekttariff ({first_cat}) vid {total_kwh_year:,} kWh/år, varierande max-effekt")

st.divider()
st.caption("Data: Energimarknadsinspektionen (Ei) — Elnätsföretagens nätavgifter 2025. Beräkningarna är approximationer baserade på typkundskategorier.")