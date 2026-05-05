import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from briefing_generator import generate_pdf, generate_pptx
import requests
import io

# ==========================================
# 1. S&P INDICATIVE RATING MATRIX
# ==========================================
# Baris: Profil Flexibility & Performance | Kolom: Profil Institutional & Economic
MATRIX = {
    (1.0, 1.7): {1.0:'AAA', 1.5:'AAA', 2.0:'AAA', 2.5:'AA+', 3.0:'AA', 3.5:'A+', 4.0:'A', 4.5:'A-', 5.0:'BBB+', 5.5:'BB+', 6.0:'BB-'},
    (1.8, 2.2): {1.0:'AAA', 1.5:'AAA', 2.0:'AA+', 2.5:'AA', 3.0:'AA-', 3.5:'A', 4.0:'A-', 4.5:'BBB+', 5.0:'BBB', 5.5:'BB+', 6.0:'BB-'},
    (2.3, 2.7): {1.0:'AAA', 1.5:'AA+', 2.0:'AA', 2.5:'AA-', 3.0:'A', 3.5:'A-', 4.0:'BBB+', 4.5:'BBB', 5.0:'BB+', 5.5:'BB', 6.0:'B+'},
    (2.8, 3.2): {1.0:'AA+', 1.5:'AA', 2.0:'AA-', 2.5:'A+', 3.0:'A-', 3.5:'BBB', 4.0:'BBB-', 4.5:'BB+', 5.0:'BB', 5.5:'BB-', 6.0:'B+'},
    (3.3, 3.7): {1.0:'AA', 1.5:'AA-', 2.0:'A+', 2.5:'A', 3.0:'BBB+', 3.5:'BBB-', 4.0:'BB+', 4.5:'BB', 5.0:'BB-', 5.5:'B+', 6.0:'B'},
    (3.8, 4.2): {1.0:'AA-', 1.5:'A+', 2.0:'A', 2.5:'BBB+', 3.0:'BBB', 3.5:'BB+', 4.0:'BB', 4.5:'BB-', 5.0:'B+', 5.5:'B', 6.0:'B'},
    (4.3, 4.7): {1.0:'A-', 1.5:'BBB+', 2.0:'BBB', 2.5:'BB+', 3.0:'BB', 3.5:'BB-', 4.0:'B+', 4.5:'B', 5.0:'B-', 5.5:'B-', 6.0:'B-'},
    (4.8, 5.2): {1.0:'BBB', 1.5:'BBB', 2.0:'BBB-', 2.5:'BB+', 3.0:'BB', 3.5:'BB-', 4.0:'B+', 4.5:'B', 5.0:'B', 5.5:'B-', 6.0:'B-'},
    (5.3, 6.0): {1.0:'BB+', 1.5:'BB+', 2.0:'BB', 2.5:'BB-', 3.0:'B+', 3.5:'B', 4.0:'B', 4.5:'B-', 5.0:'B-', 5.5:'B-', 6.0:'B-'}
}

# Pemetaan Rating ke Nilai Numerik
RATING_TO_NUM = {
    'AAA':16, 'AA+':15, 'AA':14, 'AA-':13, 'A+':12, 'A':11, 'A-':10,
    'BBB+':9, 'BBB':8, 'BBB-':7, 'BB+':6, 'BB':5, 'BB-':4, 'B+':3, 'B':2, 'B-':1, 'CCC':0
}

def get_indicative_rating(ie_prof, fp_prof):
    """Menentukan rating indikatif berdasarkan perpotongan profil di matriks S&P."""
    row_key = next((k for k in MATRIX.keys() if k[0] <= round(fp_prof, 1) <= k[1]), (5.3, 6.0))
    col_key = max(1.0, min(6.0, round(ie_prof * 2) / 2))
    return MATRIX[row_key].get(col_key, 'B-').upper()

# ==========================================
# 2. CALIBRATED SCORING ENGINE
# ==========================================
def score_economic(gdp_pc, growth, diversification="Standard"):
    """Skoring pilar Ekonomi dengan ambang batas (threshold) yang dikalibrasi."""
    if gdp_pc > 48000: s = 1
    elif gdp_pc > 34000: s = 2
    elif gdp_pc > 20000: s = 3
    elif gdp_pc > 7000: s = 4
    elif gdp_pc > 1500: s = 5 
    else: s = 6
    if growth > 4.0: s = max(1, s - 1) 
    if diversification == "High": s = max(1, s - 1)
    return s

def score_external(gefn, niip, car_receipts=20, reserves_months=6.0):
    """Skoring pilar Eksternal berdasarkan likuiditas dan posisi NIIP."""
    l_s = 1 if gefn <= 50 or reserves_months > 6 else (2 if gefn <= 75 else 3)
    d_s = 1 if niip >= 0 else (2 if niip > -20 else (4 if niip > -100 else 6))
    return (l_s + d_s) / 2, l_s, d_s

def score_fiscal(debt_gdp, int_rev, balance, flex="Neutral"):
    """Skoring pilar Fiskal yang disesuaikan dengan profil utang Indonesia."""
    perf = 1 if balance >= 0 else (2 if balance > -3 else 4)
    burd = 2 if debt_gdp <= 45 and int_rev <= 15 else (4 if debt_gdp <= 60 else 6)
    final_score = (perf + burd) / 2
    if flex == "High": final_score = max(1, final_score - 0.5)
    return final_score, perf, burd

def score_monetary(regime, credibility, inflation, depth):
    """Skoring pilar Moneter berdasarkan kredibilitas dan tingkat inflasi."""
    if inflation <= 3 and depth > 40 and credibility == "High": return 2
    if regime == "Floating" and inflation <= 10: return 3
    return 5

#### CHUNK 2

# ==========================================
# 3. DUAL-SOURCE DATA ENGINE (FIXED CLOUD FETCH)
# ==========================================

import gdown
import os


@st.cache_data(ttl=3600, show_spinner=False)  # Cache 1 jam
def get_wgi_data(file_id):
    """Download dan proses WGI data - hanya dijalankan sekali per jam"""
    output_path = f"temp_{file_id}.xlsx"
    url = f"https://drive.google.com/uc?id={file_id}"
    
    gdown.download(url, output_path, quiet=True)
    
    xl_wgi = pd.ExcelFile(output_path, engine='openpyxl')
    sheet_name = next((s for s in xl_wgi.sheet_names if 'Data' in s), xl_wgi.sheet_names[0])
    df_fitch = pd.read_excel(xl_wgi, sheet_name=sheet_name, header=None, engine='openpyxl')
    
    header_row_idx = None
    for i in range(min(15, len(df_fitch))):
        if df_fitch.iloc[i].astype(str).str.contains('Governance', case=False, na=False).any():
            header_row_idx = i
            break
    
    header_row = df_fitch.iloc[header_row_idx].astype(str)
    wgi_idx = header_row.str.contains('Governance', case=False, na=False).idxmax()
    
    wgi_lookup = df_fitch.iloc[header_row_idx+1:, [6, wgi_idx]]
    wgi_lookup.columns = ['Country', 'WGI_Score']
    wgi_lookup['Country'] = wgi_lookup['Country'].str.strip()
    wgi_lookup['WGI_Score'] = pd.to_numeric(wgi_lookup['WGI_Score'], errors='coerce')
    wgi_lookup = wgi_lookup.dropna(subset=['Country'])
    
    if os.path.exists(output_path):
        os.remove(output_path)
    
    return wgi_lookup

def process_full_data(sp_xlsx):
    # 1. RETRIEVE WGI DARI GOOGLE DRIVE (PAKAI GDOWN)
    file_id = "1P5j1E2eAD9mEp25mskOdCTbLbVVfx2TZ"
    
    try:
        wgi_lookup = get_wgi_data(file_id)  # ← pakai cache
        
    except Exception as e:
        st.error(f"⚠️ Gagal memuat WGI: {e}")
        return None
    

    # 2. PROCESS S&P MACRO DATA (UPLOADED BY USER)
    try:
        xls = pd.ExcelFile(sp_xlsx)
        
        def clean_sp_sheet(name):
            """Helper untuk membersihkan sheet S&P yang memiliki multi-level header."""
            if name not in xls.sheet_names:
                return pd.DataFrame()
                
            df = pd.read_excel(xls, sheet_name=name, header=None)
            mask = df.apply(lambda row: row.astype(str).str.contains('Country Name', na=False).any(), axis=1)
            if not mask.any(): return pd.DataFrame()
            
            h_idx = df[mask].index[0]
            metrics = df.iloc[h_idx-1].ffill()
            years = df.iloc[h_idx].apply(lambda x: str(int(x)) if isinstance(x, (int,float)) and not np.isnan(x) else str(x))
            
            cols = [f"{m_s}_{y_s}".replace('nan_', '') if pd.notna(m_s) and y_s != "nan" else str(y_s).replace('nan_', '') 
                    for m_s, y_s in zip(metrics, years)]
            
            df.columns = cols
            return df.iloc[h_idx+1:].reset_index(drop=True)

        # Muat indikator utama sesuai metodologi S&P
        e = clean_sp_sheet('Economic Data')
        g = clean_sp_sheet('General Government Data')
        m = clean_sp_sheet('Monetary Data')
        b = clean_sp_sheet('Balance-Of-Payments Data')
        xb = clean_sp_sheet('External Balance Sheet')

        # Cari kolom secara dinamis
        country_col = [c for c in e.columns if 'Country Name' in c][0]
        rating_col = [c for c in e.columns if 'LT FC rating' in c][0]
        f_c = lambda df, kw: [c for c in df.columns if kw in c and "2024" in c][0]

        # 3. KONSOLIDASI DATA MASTER
        master = pd.DataFrame({
            'Country': e[country_col].str.strip(),
            'Actual_Rating': e[rating_col].str.replace('*','').str.strip(),
            'GDP_PC': pd.to_numeric(e[f_c(e, 'GDP per capita')], errors='coerce').fillna(0) * 1000,
            'Growth': pd.to_numeric(e[f_c(e, 'Real GDP growth')], errors='coerce').fillna(0),
            'Debt_GDP': pd.to_numeric(g[f_c(g, 'Net GG debt/GDP')], errors='coerce').fillna(0),
            'Balance': pd.to_numeric(g[f_c(g, 'GG balance/GDP')], errors='coerce').fillna(0),
            'Int_Rev': pd.to_numeric(g[f_c(g, 'interest expenditure/revenues')], errors='coerce').fillna(0),
            'CPI': pd.to_numeric(m[f_c(m, 'CPI growth')], errors='coerce').fillna(0),
            'Fin_Depth': pd.to_numeric(m[f_c(m, "non-gov't sector/GDP")], errors='coerce').fillna(0),
            'GEFN': pd.to_numeric(b[f_c(b, 'Gross ext. fin. needs')], errors='coerce').fillna(100),
            'Reserves': pd.to_numeric(b[f_c(b, 'Usable reserves/CAPs')], errors='coerce').fillna(3),
            'NIIP_CAR': pd.to_numeric(xb[f_c(xb, 'Net ext. liabilities/CARs')], errors='coerce').fillna(0) * -1
        })
        
        return master.merge(wgi_lookup, on='Country', how='left')
        
    except Exception as e:
        st.error(f"⚠️ Kesalahan format file S&P: {e}")
        return None

# ==========================================
# 5. TREND DATA EXTRACTOR (Dynamic Year Detection)
# ==========================================

@st.cache_data(ttl=3600, show_spinner=False)
def extract_trend_data(_xls_bytes):
    """
    Extracts all historical indicators per country as a long-format DataFrame.
    Dynamically detects years — works for any 10-year window (e.g. 2020-2029f or 2021-2030f).
    Returns: DataFrame with columns [Country, Metric, Year, Value]
    """
    import io
    xls = pd.ExcelFile(io.BytesIO(_xls_bytes), engine='openpyxl')

    SHEET_METRICS = {
        'Economic Data': [
            'GDP per capita (000s $)', 'Real GDP growth (%)',
            'Investment/GDP (%)', 'Savings/GDP (%)', 'Exports/GDP (%)',
            'Unemployment rate (% of workforce)'
        ],
        'Monetary Data': [
            'CPI growth (%)', "Banks' claims on resident non-gov't sector/GDP",
        ],
        'General Government Data': [
            'GG balance/GDP (%)', 'Net GG debt/GDP (%)',
            'GG interest expenditure/revenues (%)', 'Gross GG debt/GDP (%)',
            'GG Revenues/GDP (%)', 'GG Expenditures/GDP (%)'
        ],
        'Balance-Of-Payments Data': [
            'Current account balance/GDP (%)',
            'Usable reserves/CAPs (months)',
            'Gross ext. fin. needs/(CAR + use. res.) (%)',
            'Net FDI/GDP (%)',
        ],
        'External Balance Sheet': [
            'Net ext. liabilities/CARs (%)',
            'Narrow net ext. debt/CARs (%)',
        ],
    }

    all_records = []

    for sheet_name, wanted_metrics in SHEET_METRICS.items():
        if sheet_name not in xls.sheet_names:
            continue

        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, engine='openpyxl')

        # Detect header rows
        metric_row_idx = None
        for i in range(min(10, len(df_raw))):
            row_vals = df_raw.iloc[i].astype(str)
            if row_vals.str.contains('Country Name', na=False).any():
                metric_row_idx = i - 1
                year_row_idx   = i
                data_start_idx = i + 1
                break

        if metric_row_idx is None:
            continue

        metrics_row = df_raw.iloc[metric_row_idx]
        years_row   = df_raw.iloc[year_row_idx]

        # --- Dynamically detect years (skip first 3 fixed cols) ---
        detected_years = []
        for val in years_row.iloc[3:]:
            if pd.isna(val):
                continue
            s = str(val).strip()
            if s in ('Country Name', 'Country Code', 'LT FC rating', 'nan', ''):
                continue
            # Normalise: "2021.0" → "2021", keep "2025e", "2026f" as-is
            try:
                s = str(int(float(s))) if s.replace('.','',1).isdigit() else s
            except Exception:
                pass
            detected_years.append(s)
        n_years = len(set(detected_years))  # unique years per metric group

        # --- Map each metric to its starting column ---
        metric_col_map = {}
        for col_idx, val in enumerate(metrics_row):
            if pd.notna(val):
                metric_name = str(val).strip()
                if metric_name in wanted_metrics:
                    metric_col_map[metric_name] = col_idx

        if not metric_col_map:
            continue

        # Determine year labels for this sheet — NORMALISE at source
        first_metric_col = min(metric_col_map.values())
        year_labels = []
        for val in years_row.iloc[first_metric_col: first_metric_col + n_years]:
                if pd.isna(val):
                    year_labels.append('?')
                else:
                    s = str(val).strip()
                    try:
                        # "2021.0" → "2021", integers → "2020"
                        s = str(int(float(s)))
                    except ValueError:
                        # "2025e", "2026f" — keep suffix intact
                        pass
                    year_labels.append(s)

        # Data rows
        data_df = df_raw.iloc[data_start_idx:].reset_index(drop=True)
        country_col_idx = 0  # Country Name is always col 0

        for _, row in data_df.iterrows():
            country = str(row.iloc[country_col_idx]).strip()
            if country in ('nan', '', 'None') or pd.isna(row.iloc[country_col_idx]):
                continue

            for metric_name, start_col in metric_col_map.items():
                for j, yr_label in enumerate(year_labels):
                    col_idx = start_col + j
                    if col_idx >= len(row):
                        continue
                    raw_val = row.iloc[col_idx]
                    try:
                        val = float(raw_val)
                    except (ValueError, TypeError):
                        val = np.nan

                    all_records.append({
                        'Country': country,
                        'Metric':  metric_name,
                        'Year':    yr_label,
                        'Value':   val,
                    })

    return pd.DataFrame(all_records)
### CHUNCK 3

# ==========================================
# 4. DASHBOARD INTERFACE & UI
# ==========================================

# Konfigurasi Halaman & Judul Profesional
st.set_page_config(page_title="Sovereign Rating Monitoring", layout="wide")

# ── ACCESS CONTROL ────────────────────────────────────────────────────────────
PASSCODE = st.secrets["APP_PASSCODE"]

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align:center; color:#1E3A8A;'>🔐 Sovereign Rating Monitoring</h2>", 
                unsafe_allow_html=True)
    st.markdown("<p style='text-align:center; color:#64748B;'>Enter passcode to access the dashboard</p>", 
                unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        entered = st.text_input("Passcode", type="password", placeholder="Enter passcode...")
        if st.button("Login", use_container_width=True):
            if entered == PASSCODE:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Incorrect passcode. Please try again.")
    st.stop()


# Branding Header
st.markdown("<h1 style='text-align: center; color: #1E3A8A;'>Sovereign Rating Monitoring</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: #64748B;'>Based on S&P Global Methodology</h4>", unsafe_allow_html=True)
st.markdown("""
<style>
    .stTabs [data-baseweb="tab"] p {
        font-size: 1.2rem !important;
        font-weight: 700 !important;
    }
</style>
""", unsafe_allow_html=True)
st.write("---")

import os

# ── Default dataset path (relative to script location) ───────────────────────
DEFAULT_DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "snp_data.xlsx")

# ── File source selector ──────────────────────────────────────────────────────
st.markdown("#### 📂 Data Source")
use_upload = st.toggle("Upload my own dataset", value=False,
                       help="Toggle on to upload a custom S&P macro dataset. "
                            "When off, the built-in default dataset is used.")

f_macro = None
data_source_label = ""

if use_upload:
    f_macro = st.file_uploader(
        "Upload S&P Macro Dataset (.xlsx)", type=["xlsx"],
        help="Your file must follow the S&P macro dataset format."
    )
    if f_macro:
        data_source_label = f"📤 Using uploaded file: **{f_macro.name}**"
    else:
        st.info("👆 Upload a file above, or toggle off to use the default dataset.")
else:
    if os.path.exists(DEFAULT_DATA_PATH):
        with open(DEFAULT_DATA_PATH, "rb") as f:
            f_macro = io.BytesIO(f.read())
            f_macro.name = "sp_macro_default.xlsx"   # give it a name attribute
        data_source_label = "📊 Using built-in default dataset"
    else:
        st.error(
            "Default dataset not found at `data/sp_macro_default.xlsx`. "
            "Please upload your file manually."
        )

if data_source_label:
    st.caption(data_source_label)

if f_macro:
    with st.spinner('Fetching cloud reference data and processing...'):
        df = process_full_data(f_macro)
    
    if df is not None:
        tabs = st.tabs(["📊 National Portfolio", "📉 Peer Comparison", "💡 Recommendation", "📄 Briefing Notes", "🧪 Methodology Simulator"])

        with tabs[0]: # NATION ANALYSIS & OVERLAY
            target = st.selectbox("Select Sovereign Target", df['Country'].unique(), 
                                   index=list(df['Country']).index('Indonesia') if 'Indonesia' in df['Country'].values else 0)
            r = df[df['Country'] == target].iloc[0]
            
            # Kalkulasi Skor pilar berdasarkan logika terkalibrasi di Chunk 1
            s_inst = max(1.0, min(6.0, 6 - (float(r['WGI_Score']) / 20))) if pd.notna(r['WGI_Score']) else 3.5
            s_eco = score_economic(r['GDP_PC'], r['Growth'], "Standard")
            s_fis, _, _ = score_fiscal(r['Debt_GDP'], r['Int_Rev'], r['Balance'], "Neutral")
            s_ext, _, _ = score_external(r['GEFN'], r['NIIP_CAR'], 20, r['Reserves'])
            s_mon = score_monetary("Floating", "High", r['CPI'], r['Fin_Depth'])

            # Hitung koordinat Matrix dan Rating Indikatif
            prof_ie = (s_inst + s_eco) / 2
            prof_fp = (s_fis + s_ext + s_mon) / 3
            srm_rating = get_indicative_rating(prof_ie, prof_fp)
            
            # Hitung selisih notch (Qualitative Overlay)
            actual_val = RATING_TO_NUM.get(r['Actual_Rating'], 8)
            srm_val = RATING_TO_NUM.get(srm_rating, 8)
            qo = actual_val - srm_val

            # Visualisasi Metrik Utama
            m1, m2, m3 = st.columns(3)
            m1.metric("Calculated SRM (Matrix)", srm_rating)
            m2.metric("Official S&P Rating", r['Actual_Rating'])
            m3.metric("Qualitative Adjustment", f"{int(qo):+} Notch", delta=int(qo), delta_color="normal")

            st.markdown("#### 🔍 In-Depth Pillar Analysis")

            pillar_tabs = st.tabs(["🏛 Economic", "🌐 External", "💰 Fiscal", "🏦 Monetary", "⚖️ Qualitative Overlay (QO)"])

            # ── ECONOMIC ──────────────────────────────────────────────
            with pillar_tabs[0]:
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("GDP per Capita (USD)", f"${r['GDP_PC']:,.0f}")
                    if r['GDP_PC'] > 48000:
                        st.success("🟢 High Income — Strong creditworthiness anchor")
                    elif r['GDP_PC'] > 20000:
                        st.info("🔵 Upper-Middle Income — Moderate resilience")
                    elif r['GDP_PC'] > 7000:
                        st.warning("🟡 Lower-Middle Income — Vulnerability to shocks")
                    else:
                        st.error("🔴 Low Income — Significant structural weakness")
                
                with col2:
                    st.metric("Trend Real GDP Growth (%)", f"{r['Growth']:.1f}%")
                    if r['Growth'] > 5:
                        st.success("🟢 High Growth — Positive rating momentum")
                    elif r['Growth'] > 3:
                        st.info("🔵 Moderate Growth — Stable trajectory")
                    elif r['Growth'] > 0:
                        st.warning("🟡 Low Growth — Limited fiscal space")
                    else:
                        st.error("🔴 Contraction — Severe credit pressure")
                
                with col3:
                    st.metric("Economic Score", f"{s_eco:.2f} / 6.00")
                    score_pct = (1 - (s_eco - 1) / 5) * 100
                    st.progress(score_pct / 100)
                    st.caption(f"Pillar strength: {score_pct:.0f}% (higher = better)")
                
                st.divider()
                st.markdown("**S&P Assessment Logic — Economic**")
                st.markdown(f"""
                | Factor | Value | S&P Implication |
                |--------|-------|-----------------|
                | GDP per Capita | ${r['GDP_PC']:,.0f} | {'Above $48K threshold — top tier' if r['GDP_PC']>48000 else 'Below $48K — score penalized'} |
                | Real GDP Growth | {r['Growth']:.1f}% | {'Above 4% — 1 notch bonus applied' if r['Growth']>4 else 'Below 4% — no growth bonus'} |
                | Diversification | Standard | No diversification bonus applied |
                """)

            # ── EXTERNAL ──────────────────────────────────────────────
            with pillar_tabs[1]:
                col1, col2, col3 = st.columns(3)
                s_ext_full, s_liq, s_debt_pos = score_external(r['GEFN'], r['NIIP_CAR'], 20, r['Reserves'])
                
                with col1:
                    st.metric("GEFN (% of CAR)", f"{r['GEFN']:.1f}%")
                    st.metric("Liquidity Score", f"{s_liq:.1f} / 6.0")
                    if r['GEFN'] <= 50:
                        st.success("🟢 Low external financing need")
                    elif r['GEFN'] <= 75:
                        st.info("🔵 Moderate — manageable rollover risk")
                    else:
                        st.error("🔴 High GEFN — elevated rollover vulnerability")
                
                with col2:
                    st.metric("NIIP (% of GDP)", f"{r['NIIP_CAR']:.1f}%")
                    st.metric("Debt Position Score", f"{s_debt_pos:.1f} / 6.0")
                    if r['NIIP_CAR'] >= 0:
                        st.success("🟢 Net creditor position")
                    elif r['NIIP_CAR'] > -20:
                        st.info("🔵 Slight net debtor — low risk")
                    elif r['NIIP_CAR'] > -100:
                        st.warning("🟡 Moderate net debtor position")
                    else:
                        st.error("🔴 Large net debtor — structural vulnerability")
                
                with col3:
                    st.metric("Reserves (Months of CXP)", f"{r['Reserves']:.1f} mo")
                    st.metric("External Score", f"{s_ext_full:.2f} / 6.00")
                    score_pct = (1 - (s_ext_full - 1) / 5) * 100
                    st.progress(score_pct / 100)
                    if r['Reserves'] > 6:
                        st.success("🟢 Adequate reserve buffer")
                    else:
                        st.warning("🟡 Below 6-month threshold")
                
                st.divider()
                st.markdown("**S&P Assessment Logic — External**")
                st.markdown(f"""
                | Factor | Value | S&P Implication |
                |--------|-------|-----------------|
                | GEFN | {r['GEFN']:.1f}% | Liquidity score: {s_liq:.1f} |
                | NIIP | {r['NIIP_CAR']:.1f}% | Debt position score: {s_debt_pos:.1f} |
                | Reserves | {r['Reserves']:.1f} months | {'Above 6mo — liquidity buffer adequate' if r['Reserves']>6 else 'Below 6mo — limited buffer'} |
                | Combined External Score | {s_ext_full:.2f} | Average of liquidity + debt position |
                """)

            # ── FISCAL ────────────────────────────────────────────────
            with pillar_tabs[2]:
                col1, col2, col3 = st.columns(3)
                s_fis_full, s_perf, s_burd = score_fiscal(r['Debt_GDP'], r['Int_Rev'], r['Balance'], "Neutral")
                
                with col1:
                    st.metric("Fiscal Balance (% GDP)", f"{r['Balance']:.1f}%")
                    st.metric("Performance Score", f"{s_perf:.1f} / 6.0")
                    if r['Balance'] >= 0:
                        st.success("🟢 Surplus — fiscal consolidation achieved")
                    elif r['Balance'] > -3:
                        st.info("🔵 Deficit within -3% — manageable")
                    else:
                        st.error("🔴 Deficit exceeds -3% — fiscal pressure")
                
                with col2:
                    st.metric("Debt-to-GDP (%)", f"{r['Debt_GDP']:.1f}%")
                    st.metric("Interest-to-Revenue (%)", f"{r['Int_Rev']:.1f}%")
                    st.metric("Burden Score", f"{s_burd:.1f} / 6.0")
                    if r['Debt_GDP'] <= 45 and r['Int_Rev'] <= 15:
                        st.success("🟢 Low debt burden")
                    elif r['Debt_GDP'] <= 60:
                        st.warning("🟡 Moderate debt — watch interest costs")
                    else:
                        st.error("🔴 High debt burden — crowd-out risk")
                
                with col3:
                    st.metric("Fiscal Score", f"{s_fis_full:.2f} / 6.00")
                    score_pct = (1 - (s_fis_full - 1) / 5) * 100
                    st.progress(score_pct / 100)
                    st.caption(f"Pillar strength: {score_pct:.0f}%")
                
                st.divider()
                st.markdown("**S&P Assessment Logic — Fiscal**")
                st.markdown(f"""
                | Factor | Value | S&P Implication |
                |--------|-------|-----------------|
                | Fiscal Balance | {r['Balance']:.1f}% | Performance score: {s_perf:.1f} |
                | Debt-to-GDP | {r['Debt_GDP']:.1f}% | {'≤45% — low burden' if r['Debt_GDP']<=45 else ('≤60% — moderate' if r['Debt_GDP']<=60 else '>60% — high burden')} |
                | Interest/Revenue | {r['Int_Rev']:.1f}% | {'≤15% — affordable' if r['Int_Rev']<=15 else '>15% — affordability concern'} |
                | Burden Score | {s_burd:.1f} | Combined debt & interest assessment |
                """)

            # ── MONETARY ──────────────────────────────────────────────
            with pillar_tabs[3]:
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("CPI Inflation (%)", f"{r['CPI']:.1f}%")
                    if r['CPI'] <= 3:
                        st.success("🟢 Price stability achieved")
                    elif r['CPI'] <= 6:
                        st.info("🔵 Moderate inflation — within tolerance")
                    elif r['CPI'] <= 10:
                        st.warning("🟡 Elevated inflation — credibility at risk")
                    else:
                        st.error("🔴 High inflation — monetary instability")
                
                with col2:
                    st.metric("Financial Depth (% GDP)", f"{r['Fin_Depth']:.1f}%")
                    if r['Fin_Depth'] > 40:
                        st.success("🟢 Deep financial market")
                    elif r['Fin_Depth'] > 20:
                        st.info("🔵 Moderate financial depth")
                    else:
                        st.warning("🟡 Shallow — limited monetary transmission")
                
                with col3:
                    st.metric("Monetary Score", f"{s_mon:.2f} / 6.00")
                    score_pct = (1 - (s_mon - 1) / 5) * 100
                    st.progress(score_pct / 100)
                    st.caption(f"Pillar strength: {score_pct:.0f}%")
                
                st.divider()
                st.markdown("**S&P Assessment Logic — Monetary**")
                st.markdown(f"""
                | Factor | Value | S&P Implication |
                |--------|-------|-----------------|
                | FX Regime | Floating | Floating regime — full score eligibility |
                | Inflation | {r['CPI']:.1f}% | {'≤3% + deep market → score 2 (best)' if r['CPI']<=3 and r['Fin_Depth']>40 else ('≤10% floating → score 3' if r['CPI']<=10 else 'High inflation → score 5')} |
                | Financial Depth | {r['Fin_Depth']:.1f}% | {'Above 40% threshold — bonus applied' if r['Fin_Depth']>40 else 'Below 40% — no depth bonus'} |
                | Monetary Credibility | High | Assumed high per regime classification |
                """)

            
            # ── QUALITATIVE OVERLAY (QO) ──────────────────────────────
            with pillar_tabs[4]:
                st.markdown("### ⚖️ Qualitative Overlay (QO) — Full Decomposition")
                st.markdown("""
                > **S&P Methodology Note:** The Qualitative Overlay reflects analyst judgment applied *after* the Sovereign Rating Model (SRM) produces an indicative rating. 
                > It aggregates up to **±3 notches** of adjustment across **5 qualitative factors**, each scored independently by the rating committee.
                """)

                col1, col2 = st.columns([3, 2])

                with col1:
                    st.markdown("#### 📋 Five QO Adjustment Factors")
                    st.markdown(f"""
                    | # | Factor | What S&P Examines | Typical Range |
                    |---|--------|-------------------|---------------|
                    | 1 | **Contingent Liabilities** | GRE debt crystallization risk, banking sector bailout burden, explicit/implicit SOE guarantees, off-budget fiscal operations | −2 to 0 notch |
                    | 2 | **Monetary Flexibility** | Effectiveness of monetary policy transmission, central bank independence, depth of local capital markets beyond SRM capture | −1 to +1 notch |
                    | 3 | **External Liquidity & IIP** | Reserve adequacy adjustments, quality of reserve assets, current account trajectory beyond GEFN/NIIP metrics | −1 to +1 notch |
                    | 4 | **Fiscal Flexibility** | Revenue mobilization capacity, tax base breadth, spending rigidity, pension/healthcare liabilities not in headline debt | −1 to +1 notch |
                    | 5 | **Political & Governance Risks** | Institutional stability, rule of law, political transition risk, geopolitical exposure beyond WGI score | −2 to 0 notch |
                    """)

                    st.divider()
                    st.markdown("#### 🔍 Indonesia-Specific QO Considerations")
                    st.markdown(f"""
                    | Factor | Indonesia Assessment | Direction |
                    |--------|---------------------|-----------|
                    | Contingent Liabilities | Significant SOE universe (PLN, Pertamina, BUMN); banking sector moderately capitalized but NPL risk exists | ⚠️ Mild negative |
                    | Monetary Flexibility | Bank Indonesia has credible inflation targeting; IDR is floating but subject to EM capital flow volatility | ➡️ Neutral |
                    | External Liquidity & IIP | Reserves adequate but current account structurally in deficit; high non-resident ownership of govt bonds (SBN) | ⚠️ Mild negative |
                    | Fiscal Flexibility | Constitutional 3% deficit ceiling limits fiscal space; narrow tax base (tax ratio ~10% GDP); commodity revenue cyclicality | ⚠️ Mild negative |
                    | Political & Governance Risks | Democratic stability maintained; decentralization creates subnational fiscal risks; institutional quality improving but below IG peers | ➡️ Neutral to mild negative |
                    """)

                with col2:
                    st.markdown("#### 📐 QO Summary for This Sovereign")

                    qo_label = "Upgrade" if qo > 0 else ("Penalty" if qo < 0 else "Neutral")
                    qo_color = "#10b981" if qo > 0 else ("#ef4444" if qo < 0 else "#6b7280")

                    st.markdown(f"""
                    <div style="padding:20px; border-radius:10px; border-left: 6px solid {qo_color}; background:#f8f9fa; margin-bottom:16px;">
                        <div style="font-size:13px; color:#666; margin-bottom:4px;">Net Qualitative Overlay</div>
                        <div style="font-size:36px; font-weight:800; color:{qo_color};">{int(qo):+} Notch</div>
                        <div style="font-size:14px; color:#444; margin-top:4px;">{qo_label} vs SRM output</div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.markdown(f"""
                    <div style="padding:16px; border-radius:8px; background:#f0f4ff; margin-bottom:12px;">
                        <div style="font-size:13px; color:#666;">SRM Indicative Rating</div>
                        <div style="font-size:24px; font-weight:700; color:#1e3a8a;">{srm_rating}</div>
                    </div>
                    <div style="padding:16px; border-radius:8px; background:#f0fdf4; margin-bottom:12px;">
                        <div style="font-size:13px; color:#666;">Official S&P Rating (after QO)</div>
                        <div style="font-size:24px; font-weight:700; color:#065f46;">{r['Actual_Rating']}</div>
                    </div>
                    """, unsafe_allow_html=True)

                    st.divider()
                    st.markdown("#### 📌 Interpretation Guide")

                    if qo > 0:
                        st.success(f"""
                        **Positive QO (+{int(qo)} notch)**  
                        S&P's committee judged qualitative factors to be **stronger** than the SRM model implies.  
                        Possible reasons: stronger institutions than WGI captures, better monetary credibility, 
                        lower contingent liability risks, or superior fiscal flexibility.
                        """)
                    elif qo < 0:
                        st.error(f"""
                        **Negative QO ({int(qo)} notch)**  
                        S&P's committee identified **hidden risks** not fully captured by the quantitative model.  
                        Possible reasons: SOE/GRE contingent liabilities, narrow fiscal space, 
                        external vulnerability beyond GEFN/NIIP, or governance concerns above WGI score.
                        """)
                    else:
                        st.info("""
                        **Neutral QO (0 notch)**  
                        The SRM model output aligns with S&P's full qualitative assessment.  
                        No material adjustments identified across the five qualitative factors.
                        """)

                    st.divider()
                    st.caption("⚠️ Note: Individual QO factor decomposition is not publicly disclosed by S&P. The net QO shown here is reverse-engineered from the difference between the SRM output and the published rating.")

            # ── SUMMARY SCORECARD ─────────────────────────────────────
            st.divider()
            st.markdown("#### 📊 Pillar Scorecard Summary")

            summary_cols = st.columns(5)
            pillars = [
                ("🏛 Institutional", s_inst, "IE Profile"),
                ("📈 Economic", s_eco, "IE Profile"),
                ("💰 Fiscal", s_fis, "FP Profile"),
                ("🌐 External", s_ext, "FP Profile"),
                ("🏦 Monetary", s_mon, "FP Profile"),
            ]

            for col, (name, score, profile) in zip(summary_cols, pillars):
                strength = (1 - (score - 1) / 5) * 100
                color = "#2ecc71" if strength >= 60 else ("#f39c12" if strength >= 40 else "#e74c3c")
                col.markdown(f"""
                <div style="text-align:center; padding:12px; border-radius:8px; border: 1px solid #ddd;">
                    <div style="font-size:13px; color:#666;">{name}</div>
                    <div style="font-size:24px; font-weight:bold; color:{color};">{score:.1f}</div>
                    <div style="font-size:11px; color:#999;">/ 6.0 · {profile}</div>
                    <div style="font-size:12px; color:{color};">{strength:.0f}% strength</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown(f"""
            <div style="margin-top:12px; padding:12px; background:#f0f4ff; border-radius:8px; text-align:center;">
                <b>IE Profile (Institutional + Economic):</b> {prof_ie:.2f} &nbsp;|&nbsp; 
                <b>FP Profile (Fiscal + External + Monetary):</b> {prof_fp:.2f} &nbsp;|&nbsp;
                <b>→ SRM Rating: {srm_rating}</b>
            </div>
            """, unsafe_allow_html=True)

        with tabs[1]:  # COMPARATIVE ANALYSIS
            st.header("Benchmarking Side-by-Side")

            # ── Rating Group Filter ──────────────────────────────────────
            RATING_GROUPS = {
                "All": None,
                "AAA": ["AAA"],
                "AA (AA+, AA, AA-)": ["AA+", "AA", "AA-"],
                "A (A+, A, A-)": ["A+", "A", "A-"],
                "BBB (BBB+, BBB, BBB-)": ["BBB+", "BBB", "BBB-"],
                "BB (BB+, BB, BB-)": ["BB+", "BB", "BB-"],
                "B (B+, B, B-)": ["B+", "B", "B-"],
                "CCC & Below": ["CCC"],
            }

            col_filter, col_multi = st.columns([1, 3])
            with col_filter:
                selected_group = st.selectbox("Filter by Rating Group", list(RATING_GROUPS.keys()), index=0)

            group_ratings = RATING_GROUPS[selected_group]
            if group_ratings:
                available_countries = df[df['Actual_Rating'].isin(group_ratings)]['Country'].unique().tolist()
            else:
                available_countries = df['Country'].unique().tolist()

            default_sel = [target] if target in available_countries else (available_countries[:1] if available_countries else [])

            with col_multi:
                sel_nations = st.multiselect(
                    "Select Countries to Compare (max 7)",
                    available_countries,
                    default=default_sel,
                    max_selections=7,
                )

            if not sel_nations:
                st.info("Select at least 1 country to begin comparison.")
                st.stop()

            # ── Rating badges ────────────────────────────────────────────
            badge_cols = st.columns(len(sel_nations))
            for i, n in enumerate(sel_nations):
                nr = df[df['Country'] == n].iloc[0]
                rating = nr['Actual_Rating']
                num = RATING_TO_NUM.get(rating, 0)
                color = "#2ecc71" if num >= 10 else ("#f39c12" if num >= 6 else "#e74c3c")
                badge_cols[i].markdown(f"""
                <div style="text-align:center;padding:8px;border-radius:6px;background:{color}22;border:1px solid {color};">
                    <div style="font-weight:700;color:{color};font-size:18px;">{rating}</div>
                    <div style="font-size:12px;color:#555;">{n}</div>
                </div>""", unsafe_allow_html=True)

            st.markdown("")

            # ── Build pillar scores ──────────────────────────────────────
            comp_list = []
            for n in sel_nations:
                nr = df[df['Country'] == n].iloc[0]
                s_inst_n = max(1.0, min(6.0, 6 - (float(nr['WGI_Score'])/20))) if pd.notna(nr['WGI_Score']) else 3.5
                s_eco_n  = score_economic(nr['GDP_PC'], nr['Growth'], "Standard")
                s_fis_n  = score_fiscal(nr['Debt_GDP'], nr['Int_Rev'], nr['Balance'], "Neutral")[0]
                s_ext_n  = score_external(nr['GEFN'], nr['NIIP_CAR'], 20, nr['Reserves'])[0]
                s_mon_n  = score_monetary("Floating", "High", nr['CPI'], nr['Fin_Depth'])
                ie_n     = (s_inst_n + s_eco_n) / 2
                fp_n     = (s_fis_n + s_ext_n + s_mon_n) / 3
                srm_n    = get_indicative_rating(ie_n, fp_n)
                qo_n     = RATING_TO_NUM.get(nr['Actual_Rating'], 8) - RATING_TO_NUM.get(srm_n, 8)
                comp_list.append({
                    "Country": n, "🏛 Institutional": s_inst_n,
                    "📈 Economic": s_eco_n, "💰 Fiscal": s_fis_n,
                    "🌐 External": s_ext_n, "🏦 Monetary": s_mon_n,
                    "_ie": ie_n, "_fp": fp_n, "_srm": srm_n, "_qo": qo_n, "_nr": nr,
                })

            # ── Sub-tabs ─────────────────────────────────────────────────
            bench_tabs = st.tabs([
                "📊 Pillar Scores",
                "📋 Metrics Snapshot",
                "📈 Historical Trends",
            ])

            # ════════════════════════════════════════════════════════════
            # SUB-TAB 1 — PILLAR SCORES
            # ════════════════════════════════════════════════════════════
            with bench_tabs[0]:
                chart_df = pd.DataFrame([{
                    "Country": c["Country"],
                    "🏛 Institutional": c["🏛 Institutional"],
                    "📈 Economic": c["📈 Economic"],
                    "💰 Fiscal": c["💰 Fiscal"],
                    "🌐 External": c["🌐 External"],
                    "🏦 Monetary": c["🏦 Monetary"],
                } for c in comp_list])

                melted = chart_df.melt(id_vars='Country', var_name='Pillar', value_name='Score')
                fig_bar = px.bar(
                    melted, x='Pillar', y='Score', color='Country',
                    barmode='group', range_y=[0, 6],
                    labels={'Pillar': 'Pillar', 'Score': 'Score (1=Best, 6=Worst)'},
                    color_discrete_sequence=px.colors.qualitative.Prism,
                    text_auto='.1f',
                )
                fig_bar.update_layout(plot_bgcolor='rgba(0,0,0,0)', yaxis=dict(gridcolor='#e5e7eb'))
                st.plotly_chart(fig_bar, use_container_width=True)

                st.markdown("##### Pillar Score Heatmap")
                score_df = chart_df.set_index("Country").round(2)
                st.dataframe(
                    score_df.style.background_gradient(cmap="RdYlGn_r", axis=None, vmin=1, vmax=6),
                    use_container_width=True,
                )

            # ════════════════════════════════════════════════════════════
            # SUB-TAB 2 — METRICS SNAPSHOT (color-coded)
            # ════════════════════════════════════════════════════════════
            with bench_tabs[1]:
                METRIC_THRESHOLDS = {
                    "WGI Score":                      (True,  [60, 40]),
                    "GDP per Capita (USD)":            (True,  [20000, 7000]),
                    "Real GDP Growth (%)":             (True,  [4.0, 1.0]),
                    "Fiscal Balance (% GDP)":          (True,  [-3.0, -5.0]),
                    "Debt-to-GDP (%)":                 (False, [45, 60]),
                    "Interest/Revenue (%)":            (False, [15, 25]),
                    "GEFN (% CAR)":                   (False, [50, 75]),
                    "NIIP (% GDP)":                   (True,  [0, -50]),
                    "Reserves (months)":               (True,  [6.0, 3.0]),
                    "CPI Inflation (%)":               (False, [3.0, 6.0]),
                    "Financial Depth (% GDP)":         (True,  [40, 20]),
                    "IE Profile":                      (False, [2.5, 3.5]),
                    "FP Profile":                      (False, [2.5, 3.5]),
                    "QO (Notch)":                      (True,  [1, 0]),
                }

                GREEN  = "background-color:#d1fae5;color:#065f46;"
                YELLOW = "background-color:#fef3c7;color:#78350f;"
                RED    = "background-color:#fee2e2;color:#991b1b;"
                BOLD   = "font-weight:600;"

                metric_rows = []
                for c in comp_list:
                    nr = c["_nr"]
                    metric_rows.append({
                        "Country":                  c["Country"],
                        "Actual Rating":            nr['Actual_Rating'],
                        "SRM (Model)":              c["_srm"],
                        "QO (Notch)":               int(c["_qo"]),
                        "IE Profile":               round(c["_ie"], 2),
                        "FP Profile":               round(c["_fp"], 2),
                        "WGI Score":                round(float(nr['WGI_Score']), 1) if pd.notna(nr['WGI_Score']) else np.nan,
                        "GDP per Capita (USD)":     round(nr['GDP_PC'], 0),
                        "Real GDP Growth (%)":      round(nr['Growth'], 1),
                        "Fiscal Balance (% GDP)":   round(nr['Balance'], 1),
                        "Debt-to-GDP (%)":          round(nr['Debt_GDP'], 1),
                        "Interest/Revenue (%)":     round(nr['Int_Rev'], 1),
                        "GEFN (% CAR)":             round(nr['GEFN'], 1),
                        "NIIP (% GDP)":             round(nr['NIIP_CAR'], 1),
                        "Reserves (months)":        round(nr['Reserves'], 1),
                        "CPI Inflation (%)":        round(nr['CPI'], 1),
                        "Financial Depth (% GDP)":  round(nr['Fin_Depth'], 1),
                    })

                metric_df_raw = pd.DataFrame(metric_rows).set_index("Country").T

                FORMAT_MAP = {
                    "GDP per Capita (USD)":    lambda v: f"${float(v):,.0f}",
                    "Real GDP Growth (%)":     lambda v: f"{float(v):.1f}%",
                    "Fiscal Balance (% GDP)":  lambda v: f"{float(v):.1f}%",
                    "Debt-to-GDP (%)":         lambda v: f"{float(v):.1f}%",
                    "Interest/Revenue (%)":    lambda v: f"{float(v):.1f}%",
                    "GEFN (% CAR)":            lambda v: f"{float(v):.1f}%",
                    "NIIP (% GDP)":            lambda v: f"{float(v):.1f}%",
                    "Reserves (months)":       lambda v: f"{float(v):.1f} mo",
                    "CPI Inflation (%)":       lambda v: f"{float(v):.1f}%",
                    "Financial Depth (% GDP)": lambda v: f"{float(v):.1f}%",
                    "QO (Notch)":              lambda v: f"{int(float(v)):+}",
                }
                display_df = metric_df_raw.copy()
                for row, fmt in FORMAT_MAP.items():
                    if row in display_df.index:
                        display_df.loc[row] = display_df.loc[row].apply(
                            lambda v: fmt(v) if not (isinstance(v, float) and np.isnan(v)) else "N/A"
                        )

                def style_full_table(df):
                    styles = pd.DataFrame("", index=df.index, columns=df.columns)
                    for row in df.index:
                        for col in df.columns:
                            if row in ("Actual Rating", "SRM (Model)"):
                                num = RATING_TO_NUM.get(str(df.loc[row, col]).replace("*","").strip(), -1)
                                styles.loc[row, col] = (GREEN if num >= 10 else YELLOW if num >= 6 else RED) + BOLD
                                continue
                            if row not in METRIC_THRESHOLDS:
                                continue
                            try:
                                val = float(metric_df_raw.loc[row, col])
                            except (ValueError, TypeError):
                                continue
                            higher_better, (g_thresh, y_thresh) = METRIC_THRESHOLDS[row]
                            if higher_better:
                                styles.loc[row, col] = (GREEN if val >= g_thresh else YELLOW if val >= y_thresh else RED) + BOLD
                            else:
                                styles.loc[row, col] = (GREEN if val <= g_thresh else YELLOW if val <= y_thresh else RED) + BOLD
                    return styles

                st.markdown("""
                <div style="display:flex;gap:16px;margin-bottom:8px;font-size:13px;">
                    <span style="background:#d1fae5;color:#065f46;padding:3px 10px;border-radius:4px;font-weight:600;">🟢 Strong</span>
                    <span style="background:#fef3c7;color:#78350f;padding:3px 10px;border-radius:4px;font-weight:600;">🟡 Moderate</span>
                    <span style="background:#fee2e2;color:#991b1b;padding:3px 10px;border-radius:4px;font-weight:600;">🔴 Weak</span>
                </div>""", unsafe_allow_html=True)

                st.dataframe(
                    display_df.style.apply(style_full_table, axis=None),
                    use_container_width=True, height=620,
                )

            # ════════════════════════════════════════════════════════════
            # SUB-TAB 3 — HISTORICAL TRENDS
            # ════════════════════════════════════════════════════════════
            with bench_tabs[2]:
                st.markdown("#### 📈 Historical & Forecast Indicator Trends")
                st.caption("Data spans all available years from the uploaded dataset. Shaded years are estimates/forecasts.")

                # Load trend data (cached)
                with st.spinner("Extracting trend data..."):
                    if hasattr(f_macro, 'getvalue'):
                        xls_bytes = f_macro.getvalue()   # Streamlit UploadedFile
                    else:
                        f_macro.seek(0)
                        xls_bytes = f_macro.read()       # BytesIO from default file
                        f_macro.seek(0)                  # reset for any subsequent reads
                    trend_df = extract_trend_data(xls_bytes)

                if trend_df.empty:
                    st.warning("Could not extract trend data from this file.")
                else:
                    # Filter to selected countries
                    trend_filtered = trend_df[trend_df['Country'].isin(sel_nations)].copy()

                    # Detect all available metrics
                    available_metrics = sorted(trend_filtered['Metric'].dropna().unique().tolist())

                    # Detect all years dynamically
                    all_years = trend_filtered['Year'].unique().tolist()

                    # Friendly metric selector with groupings
                    METRIC_GROUPS = {
                        "🏛 Institutional": [],
                        "📈 Economic": [
                            'GDP per capita (000s $)', 'Real GDP growth (%)',
                            'Investment/GDP (%)', 'Savings/GDP (%)',
                            'Exports/GDP (%)', 'Unemployment rate (% of workforce)'
                        ],
                        "💰 Fiscal": [
                            'GG balance/GDP (%)', 'Net GG debt/GDP (%)',
                            'Gross GG debt/GDP (%)', 'GG interest expenditure/revenues (%)',
                            'GG Revenues/GDP (%)', 'GG Expenditures/GDP (%)'
                        ],
                        "🌐 External": [
                            'Current account balance/GDP (%)',
                            'Usable reserves/CAPs (months)',
                            'Gross ext. fin. needs/(CAR + use. res.) (%)',
                            'Net FDI/GDP (%)',
                            'Net ext. liabilities/CARs (%)',
                            'Narrow net ext. debt/CARs (%)',
                        ],
                        "🏦 Monetary": [
                            'CPI growth (%)',
                            "Banks' claims on resident non-gov't sector/GDP",
                        ],
                    }

                    # Build flat selector options (only those present in data)
                    selector_options = [m for m in available_metrics]

                    tc1, tc2 = st.columns([1, 3])
                    with tc1:
                        st.markdown("**Select Pillar**")
                        pillar_filter = st.radio(
                            "Pillar",
                            ["All"] + [k for k in METRIC_GROUPS if k != "🏛 Institutional"],
                            label_visibility="collapsed"
                        )
                    with tc2:
                        # Filter metrics by pillar
                        if pillar_filter == "All":
                            filtered_metrics = selector_options
                        else:
                            filtered_metrics = [m for m in METRIC_GROUPS.get(pillar_filter, []) if m in selector_options]

                        if not filtered_metrics:
                            st.info("No metrics available for this pillar.")
                        else:
                            selected_metrics = st.multiselect(
                                "Select Indicators to Plot",
                                filtered_metrics,
                                default=filtered_metrics[:2],
                            )

                    if filtered_metrics and selected_metrics:
                        # Detect forecast years (contain 'e' or 'f')
                        forecast_years = [y for y in all_years if isinstance(y, str) and (y.endswith('e') or y.endswith('f'))]
                        historical_years = [y for y in all_years if y not in forecast_years]

                    # ── Forecast toggle ──────────────────────────────────
                    show_forecast = st.toggle(
                        "Show Estimates (e) & Forecasts (f)",
                        value=True,
                        help="'e' = current year estimate, 'f' = multi-year forecast"
                    )

                    # ── Helper functions (defined once outside the loop) ──
                    def year_to_num(y):
                        try:
                            return int(float(str(y).replace('e', '').replace('f', '')))
                        except:
                            return 9999

                    def is_estimate(y):
                        return str(y).strip().endswith('e')

                    def is_forecast_year(y):
                        return str(y).strip().endswith('f')

                    def is_projected(y):
                        return is_estimate(y) or is_forecast_year(y)

                    PRISM_COLORS = px.colors.qualitative.Prism

                    for metric in selected_metrics:
                        metric_data = trend_filtered[trend_filtered['Metric'] == metric].copy()
                        metric_data = metric_data.dropna(subset=['Value'])

                        if metric_data.empty:
                            st.caption(f"No data for: {metric}")
                            continue

                        # ── Assign columns ────────────────────────────────
                        metric_data['Year']        = metric_data['Year'].astype(str)
                        metric_data['XNum']        = metric_data['Year'].apply(year_to_num)
                        metric_data['IsEstimate']  = metric_data['Year'].apply(is_estimate)
                        metric_data['IsForecast']  = metric_data['Year'].apply(is_forecast_year)
                        metric_data['IsProjected'] = metric_data['IsEstimate'] | metric_data['IsForecast']
                        metric_data                = metric_data.sort_values('XNum').reset_index(drop=True)

                        # ── Build tick labels BEFORE any filtering ────────
                        all_years_sorted = metric_data.drop_duplicates('XNum').sort_values('XNum')
                        tickvals = all_years_sorted['XNum'].tolist()
                        ticktext = all_years_sorted['Year'].tolist()

                        # ── Filter if toggle off ──────────────────────────
                        if not show_forecast:
                            metric_data = metric_data[~metric_data['IsProjected']].reset_index(drop=True)
                            tickvals = [v for v, t in zip(tickvals, ticktext) if not is_projected(t)]
                            ticktext = [t for t in ticktext if not is_projected(t)]

                        if metric_data.empty:
                            st.caption(f"No data for: {metric}")
                            continue

                        primary_country = sel_nations[0]
                        fig = go.Figure()

                        for i, country in enumerate(sel_nations):
                            cdata = metric_data[metric_data['Country'] == country].sort_values('XNum').reset_index(drop=True)
                            if cdata.empty:
                                continue

                            color      = PRISM_COLORS[i % len(PRISM_COLORS)]
                            is_primary = (country == primary_country)
                            lw         = 2.5 if is_primary else 1.8

                            hist = cdata[~cdata['IsProjected']]
                            est  = cdata[cdata['IsEstimate']]
                            fore = cdata[cdata['IsForecast']]

                            # ── 1. Historical: solid ──────────────────────
                            if not hist.empty:
                                fig.add_trace(go.Scatter(
                                    x=hist['XNum'].tolist(),
                                    y=hist['Value'].tolist(),
                                    mode='lines+markers',
                                    name=country,
                                    line=dict(color=color, width=lw, dash='solid'),
                                    marker=dict(size=6, color=color),
                                    legendgroup=country,
                                    showlegend=True,
                                    hovertemplate=f"<b>{country}</b><br>%{{x}}: %{{y:.2f}}<extra></extra>",
                                ))

                            # ── 2. Estimate (e): dotted ───────────────────
                            if show_forecast and not est.empty:
                                if not hist.empty:
                                    # invisible connector, no hover
                                    fig.add_trace(go.Scatter(
                                        x=[int(hist['XNum'].iloc[-1])] + est['XNum'].tolist()[:1],
                                        y=[float(hist['Value'].iloc[-1])] + est['Value'].tolist()[:1],
                                        mode='lines',
                                        line=dict(color=color, width=lw, dash='dot'),
                                        legendgroup=country,
                                        showlegend=False,
                                        hoverinfo='skip',
                                    ))
                                fig.add_trace(go.Scatter(
                                    x=est['XNum'].tolist(),
                                    y=est['Value'].tolist(),
                                    mode='lines+markers',
                                    name=f"{country} (estimate)",
                                    line=dict(color=color, width=lw, dash='dot'),
                                    marker=dict(size=5, color=color, symbol='diamond'),
                                    legendgroup=country,
                                    showlegend=False,
                                    hovertemplate=f"<b>{country} (est)</b><br>%{{x}}: %{{y:.2f}}<extra></extra>",
                                ))

                            # ── 3. Forecast (f): dashed ───────────────────
                            if show_forecast and not fore.empty:
                                if not est.empty:
                                    anchor_x = int(est['XNum'].iloc[-1])
                                    anchor_y = float(est['Value'].iloc[-1])
                                elif not hist.empty:
                                    anchor_x = int(hist['XNum'].iloc[-1])
                                    anchor_y = float(hist['Value'].iloc[-1])
                                else:
                                    anchor_x = None
                                    anchor_y = None

                                if anchor_x is not None:
                                    # invisible connector, no hover
                                    fig.add_trace(go.Scatter(
                                        x=[anchor_x] + fore['XNum'].tolist()[:1],
                                        y=[anchor_y] + fore['Value'].tolist()[:1],
                                        mode='lines',
                                        line=dict(color=color, width=lw, dash='dash'),
                                        legendgroup=country,
                                        showlegend=False,
                                        hoverinfo='skip',
                                    ))

                                fx = fore['XNum'].tolist()
                                fy = fore['Value'].tolist()
                                y_floor = min(0.0, float(metric_data['Value'].min()))

                                # shaded area
                                fig.add_trace(go.Scatter(
                                    x=fx + fx[::-1],
                                    y=fy + [y_floor] * len(fx),
                                    fill='toself',
                                    fillcolor=f'rgba(99,102,241,{0.09 if is_primary else 0.04})',
                                    line=dict(width=0),
                                    legendgroup=country,
                                    showlegend=False,
                                    hoverinfo='skip',
                                ))

                                # forecast line
                                fig.add_trace(go.Scatter(
                                    x=fx, y=fy,
                                    mode='lines+markers',
                                    name=f"{country} (forecast)",
                                    line=dict(color=color, width=lw, dash='dash'),
                                    marker=dict(size=5, color=color, symbol='circle-open'),
                                    legendgroup=country,
                                    showlegend=False,
                                    hovertemplate=f"<b>{country} (fcst)</b><br>%{{x}}: %{{y:.2f}}<extra></extra>",
                                ))

                        # ── Boundary lines ────────────────────────────────
                        if show_forecast:
                            hist_only = metric_data[~metric_data['IsProjected']]
                            est_only  = metric_data[metric_data['IsEstimate']]
                            fore_only = metric_data[metric_data['IsForecast']]

                            if not hist_only.empty and not est_only.empty:
                                fig.add_shape(
                                    type="line", xref="x", yref="paper",
                                    x0=int(hist_only['XNum'].max()),
                                    x1=int(hist_only['XNum'].max()),
                                    y0=0, y1=1,
                                    line=dict(color="rgba(234,179,8,0.5)", width=1.2, dash="dot"),
                                )

                            if not est_only.empty and not fore_only.empty:
                                fig.add_shape(
                                    type="line", xref="x", yref="paper",
                                    x0=int(est_only['XNum'].max()),
                                    x1=int(est_only['XNum'].max()),
                                    y0=0, y1=1,
                                    line=dict(color="rgba(99,102,241,0.45)", width=1.2, dash="dot"),
                                )

                        fig.update_layout(
                            title=dict(text=f"<b>{metric}</b>", font=dict(size=14)),
                            xaxis=dict(
                                tickmode='array',
                                tickvals=tickvals,
                                ticktext=ticktext,
                                showgrid=False,
                                tickangle=-30,
                            ),
                            yaxis=dict(
                                gridcolor='#e5e7eb',
                                zeroline=True,
                                zerolinecolor='#d1d5db',
                            ),
                            plot_bgcolor='rgba(0,0,0,0)',
                            paper_bgcolor='rgba(0,0,0,0)',
                            legend_title_text='Country',
                            hovermode='x unified',
                            height=400,
                            margin=dict(t=55, b=40, l=10, r=10),
                        )
                        st.plotly_chart(fig, use_container_width=True)

                    # ── Summary stats table for selected metrics ─────────
                    st.divider()
                    st.markdown("##### 📊 Latest Available Value per Indicator")

                    RAW_METRIC_THRESHOLDS = {
                        'GDP per capita (000s $)':                          (True,  [20, 7]),
                        'Real GDP growth (%)':                              (True,  [4.0, 1.0]),
                        'Investment/GDP (%)':                               (True,  [25, 15]),
                        'Savings/GDP (%)':                                  (True,  [25, 15]),
                        'Exports/GDP (%)':                                  (True,  [30, 15]),
                        'Unemployment rate (% of workforce)':               (False, [5, 10]),
                        'CPI growth (%)':                                   (False, [3.0, 6.0]),
                        "Banks' claims on resident non-gov't sector/GDP":   (True,  [40, 20]),
                        'GG balance/GDP (%)':                               (True,  [-3.0, -5.0]),
                        'Net GG debt/GDP (%)':                              (False, [45, 60]),
                        'Gross GG debt/GDP (%)':                            (False, [60, 90]),
                        'GG interest expenditure/revenues (%)':             (False, [15, 25]),
                        'GG Revenues/GDP (%)':                              (True,  [30, 20]),
                        'GG Expenditures/GDP (%)':                          (False, [35, 45]),
                        'Current account balance/GDP (%)':                  (True,  [0, -3.0]),
                        'Usable reserves/CAPs (months)':                    (True,  [6.0, 3.0]),
                        'Gross ext. fin. needs/(CAR + use. res.) (%)':      (False, [50, 75]),
                        'Net FDI/GDP (%)':                                  (True,  [3.0, 0]),
                        'Net ext. liabilities/CARs (%)':                    (False, [100, 200]),
                        'Narrow net ext. debt/CARs (%)':                    (False, [50, 100]),
                    }

                    summary_rows_num = []
                    summary_rows_disp = []
                    for metric in selected_metrics:
                        num_row  = {"Indicator": metric}
                        disp_row = {"Indicator": metric}
                        for n in sel_nations:
                            sub = trend_filtered[
                                (trend_filtered['Country'] == n) &
                                (trend_filtered['Metric'] == metric)
                            ].dropna(subset=['Value'])
                            if sub.empty:
                                num_row[n]  = np.nan
                                disp_row[n] = "N/A"
                            else:
                                def yk(y):
                                    try: return float(str(y).replace('e','').replace('f',''))
                                    except: return 0
                                sub2 = sub.copy()
                                sub2['_s'] = sub2['Year'].apply(yk)
                                latest = sub2.sort_values('_s').iloc[-1]
                                val = round(latest['Value'], 2)
                                yr  = latest['Year']
                                num_row[n]  = val
                                disp_row[n] = f"{val:,.2f}  ({yr})"
                        summary_rows_num.append(num_row)
                        summary_rows_disp.append(disp_row)

                    num_df  = pd.DataFrame(summary_rows_num).set_index("Indicator")
                    disp_df = pd.DataFrame(summary_rows_disp).set_index("Indicator")

                    GREEN_S  = "background-color:#d1fae5;color:#065f46;font-weight:600;"
                    YELLOW_S = "background-color:#fef3c7;color:#78350f;font-weight:600;"
                    RED_S    = "background-color:#fee2e2;color:#991b1b;font-weight:600;"
                    NONE_S   = ""

                    def style_summary_table(display):
                        styles = pd.DataFrame(NONE_S, index=display.index, columns=display.columns)
                        for metric in display.index:
                            if metric not in RAW_METRIC_THRESHOLDS:
                                continue
                            higher_better, (g_thresh, y_thresh) = RAW_METRIC_THRESHOLDS[metric]
                            for col in display.columns:
                                try:
                                    val = float(num_df.loc[metric, col])
                                except (ValueError, TypeError, KeyError):
                                    continue
                                if np.isnan(val):
                                    continue
                                if higher_better:
                                    color = GREEN_S if val >= g_thresh else (YELLOW_S if val >= y_thresh else RED_S)
                                else:
                                    color = GREEN_S if val <= g_thresh else (YELLOW_S if val <= y_thresh else RED_S)
                                styles.loc[metric, col] = color
                        return styles

                    st.markdown("""
                    <div style="display:flex;gap:16px;margin-bottom:8px;font-size:13px;">
                        <span style="background:#d1fae5;color:#065f46;padding:3px 10px;border-radius:4px;font-weight:600;">🟢 Strong</span>
                        <span style="background:#fef3c7;color:#78350f;padding:3px 10px;border-radius:4px;font-weight:600;">🟡 Moderate</span>
                        <span style="background:#fee2e2;color:#991b1b;padding:3px 10px;border-radius:4px;font-weight:600;">🔴 Weak</span>
                        <span style="color:#888;font-size:12px;margin-left:8px;">Values show latest available year in parentheses</span>
                    </div>""", unsafe_allow_html=True)

                    st.dataframe(
                        disp_df.style.apply(style_summary_table, axis=None),
                        use_container_width=True,
                    )

        with tabs[2]: # RECOMMENDATION
            st.markdown("<h2 style='color:#1E3A8A;'>💡 Sovereign Rating Recommendation</h2>", unsafe_allow_html=True)
            st.markdown(f"<h4 style='color:#64748B;'>Data-driven advisory for <b>{target}</b> — based on S&P methodology signals</h4>", unsafe_allow_html=True)

            # ── Re-use scores from tabs[0] ────────────────────────────
            pillars = {
                "Institutional": {"score": s_inst, "profile": "IE"},
                "Economic":      {"score": s_eco,  "profile": "IE"},
                "Fiscal":        {"score": s_fis,  "profile": "FP"},
                "External":      {"score": s_ext,  "profile": "FP"},
                "Monetary":      {"score": s_mon,  "profile": "FP"},
            }

            # ── Classify each pillar ──────────────────────────────────
            strengths   = {k: v for k, v in pillars.items() if v["score"] <= 2.5}
            moderates   = {k: v for k, v in pillars.items() if 2.5 < v["score"] <= 3.5}
            weaknesses  = {k: v for k, v in pillars.items() if v["score"] > 3.5}

            # ── Metric-level signals ──────────────────────────────────
            signals = []

            # Economic signals
            if r['GDP_PC'] < 7000:
                signals.append({"pillar": "Economic", "type": "weakness", "metric": "GDP per Capita",
                    "value": f"${r['GDP_PC']:,.0f}",
                    "message": "Below $7,000 threshold — S&P classifies this as low-income, a structural drag on the rating ceiling.",
                    "action": "Accelerate investment climate reforms to boost productivity and GDP per capita trajectory."})
            elif r['GDP_PC'] < 20000:
                signals.append({"pillar": "Economic", "type": "moderate", "metric": "GDP per Capita",
                    "value": f"${r['GDP_PC']:,.0f}",
                    "message": "In $7K–$20K range — lower-middle income band with limited shock absorption capacity.",
                    "action": "Structural reforms in manufacturing and services to move up the value chain."})
            else:
                signals.append({"pillar": "Economic", "type": "strength", "metric": "GDP per Capita",
                    "value": f"${r['GDP_PC']:,.0f}",
                    "message": "Above $20,000 — provides meaningful creditworthiness anchor.",
                    "action": "Maintain investment environment to sustain income level trajectory."})

            if r['Growth'] > 4.0:
                signals.append({"pillar": "Economic", "type": "strength", "metric": "Real GDP Growth",
                    "value": f"{r['Growth']:.1f}%",
                    "message": "Above 4% threshold — S&P applies 1-notch improvement bonus to economic score.",
                    "action": "Highlight growth consistency in investor relations narrative to reinforce positive momentum."})
            elif r['Growth'] < 1.0:
                signals.append({"pillar": "Economic", "type": "weakness", "metric": "Real GDP Growth",
                    "value": f"{r['Growth']:.1f}%",
                    "message": "Below 1% — near-stagnation signals fiscal and debt sustainability concerns.",
                    "action": "Credible fiscal stimulus or structural reform agenda needed to restore growth trajectory."})

            # Fiscal signals
            if r['Balance'] < -5.0:
                signals.append({"pillar": "Fiscal", "type": "weakness", "metric": "Fiscal Balance",
                    "value": f"{r['Balance']:.1f}% GDP",
                    "message": "Deficit exceeds -5% of GDP — materially above S&P's -3% manageable threshold.",
                    "action": "Present a credible medium-term fiscal consolidation path with specific revenue/expenditure targets."})
            elif r['Balance'] < -3.0:
                signals.append({"pillar": "Fiscal", "type": "moderate", "metric": "Fiscal Balance",
                    "value": f"{r['Balance']:.1f}% GDP",
                    "message": "Deficit between -3% and -5% — outside the manageable band, watch trend direction.",
                    "action": "Demonstrate fiscal consolidation trajectory even if ceiling breach is temporary."})
            else:
                signals.append({"pillar": "Fiscal", "type": "strength", "metric": "Fiscal Balance",
                    "value": f"{r['Balance']:.1f}% GDP",
                    "message": "Within or above -3% threshold — S&P views this as manageable fiscal performance.",
                    "action": "Communicate fiscal discipline narrative proactively in rating dialogue."})

            if r['Debt_GDP'] > 60:
                signals.append({"pillar": "Fiscal", "type": "weakness", "metric": "Debt-to-GDP",
                    "value": f"{r['Debt_GDP']:.1f}%",
                    "message": "Above 60% — S&P assigns high debt burden score, pressuring the FP profile.",
                    "action": "Commit to debt stabilization targets; highlight debt composition (local vs FX, maturity profile)."})
            elif r['Debt_GDP'] > 45:
                signals.append({"pillar": "Fiscal", "type": "moderate", "metric": "Debt-to-GDP",
                    "value": f"{r['Debt_GDP']:.1f}%",
                    "message": "Between 45%–60% — moderate burden band; direction of travel matters to S&P.",
                    "action": "Show declining debt/GDP trajectory in medium-term fiscal framework."})
            else:
                signals.append({"pillar": "Fiscal", "type": "strength", "metric": "Debt-to-GDP",
                    "value": f"{r['Debt_GDP']:.1f}%",
                    "message": "Below 45% — low debt burden, strong fiscal flexibility signal.",
                    "action": "Emphasize low debt as fiscal space buffer in sovereign narrative."})

            if r['Int_Rev'] > 25:
                signals.append({"pillar": "Fiscal", "type": "weakness", "metric": "Interest/Revenue",
                    "value": f"{r['Int_Rev']:.1f}%",
                    "message": "Above 25% — interest payments crowd out productive spending; S&P views this as a severe affordability concern.",
                    "action": "Debt management strategy: lengthen maturities, reduce FX exposure, develop local bond market depth."})
            elif r['Int_Rev'] > 15:
                signals.append({"pillar": "Fiscal", "type": "moderate", "metric": "Interest/Revenue",
                    "value": f"{r['Int_Rev']:.1f}%",
                    "message": "Between 15%–25% — affordability concern emerging; watch revenue growth vs interest trajectory.",
                    "action": "Broaden tax base to improve revenue denominator; resist new high-cost borrowing."})
            else:
                signals.append({"pillar": "Fiscal", "type": "strength", "metric": "Interest/Revenue",
                    "value": f"{r['Int_Rev']:.1f}%",
                    "message": "Below 15% — debt service is affordable relative to revenues.",
                    "action": "Maintain this ratio by controlling new debt issuance costs."})

            # External signals
            if r['GEFN'] > 75:
                signals.append({"pillar": "External", "type": "weakness", "metric": "GEFN",
                    "value": f"{r['GEFN']:.1f}%",
                    "message": "Above 75% — elevated rollover and refinancing vulnerability; high sensitivity to global risk-off.",
                    "action": "Reduce reliance on short-term external financing; build reserve buffer above 6 months."})
            elif r['GEFN'] > 50:
                signals.append({"pillar": "External", "type": "moderate", "metric": "GEFN",
                    "value": f"{r['GEFN']:.1f}%",
                    "message": "Between 50%–75% — moderate external financing need; manageable but directionally important.",
                    "action": "Diversify investor base (domestic vs foreign); extend debt maturity profile."})
            else:
                signals.append({"pillar": "External", "type": "strength", "metric": "GEFN",
                    "value": f"{r['GEFN']:.1f}%",
                    "message": "Below 50% — low gross external financing need; resilient to global liquidity tightening.",
                    "action": "Highlight low GEFN as external resilience strength in rating discussions."})

            if r['NIIP_CAR'] < -100:
                signals.append({"pillar": "External", "type": "weakness", "metric": "NIIP",
                    "value": f"{r['NIIP_CAR']:.1f}% GDP",
                    "message": "Below -100% — large net debtor position; structural external vulnerability.",
                    "action": "Focus on current account improvement and FDI attraction over portfolio debt flows."})
            elif r['NIIP_CAR'] < -20:
                signals.append({"pillar": "External", "type": "moderate", "metric": "NIIP",
                    "value": f"{r['NIIP_CAR']:.1f}% GDP",
                    "message": "Between -20% and -100% — moderate net debtor; watch trajectory and composition.",
                    "action": "Shift external financing mix towards FDI (non-debt creating) vs portfolio inflows."})
            else:
                signals.append({"pillar": "External", "type": "strength", "metric": "NIIP",
                    "value": f"{r['NIIP_CAR']:.1f}% GDP",
                    "message": "Above -20% — low net debtor or net creditor position; strong external balance sheet.",
                    "action": "Maintain current account discipline to preserve IIP position."})

            if r['Reserves'] < 3:
                signals.append({"pillar": "External", "type": "weakness", "metric": "Reserves",
                    "value": f"{r['Reserves']:.1f} months",
                    "message": "Below 3 months — critically low reserve buffer; high vulnerability to sudden stop.",
                    "action": "Priority: rebuild reserves via current account adjustment or precautionary IMF facility."})
            elif r['Reserves'] < 6:
                signals.append({"pillar": "External", "type": "moderate", "metric": "Reserves",
                    "value": f"{r['Reserves']:.1f} months",
                    "message": "Between 3–6 months — below S&P's 6-month adequacy threshold.",
                    "action": "Communicate reserve accumulation strategy; highlight FX intervention capacity."})
            else:
                signals.append({"pillar": "External", "type": "strength", "metric": "Reserves",
                    "value": f"{r['Reserves']:.1f} months",
                    "message": "Above 6 months — adequate reserve buffer per S&P threshold.",
                    "action": "Highlight reserve adequacy relative to GEFN as external resilience anchor."})

            # Monetary signals
            if r['CPI'] > 10:
                signals.append({"pillar": "Monetary", "type": "weakness", "metric": "CPI Inflation",
                    "value": f"{r['CPI']:.1f}%",
                    "message": "Above 10% — monetary instability; S&P assigns score 5, highest monetary risk.",
                    "action": "Central bank credibility restoration is prerequisite; tightening cycle narrative critical."})
            elif r['CPI'] > 6:
                signals.append({"pillar": "Monetary", "type": "moderate", "metric": "CPI Inflation",
                    "value": f"{r['CPI']:.1f}%",
                    "message": "Between 6%–10% — elevated; monetary credibility under pressure.",
                    "action": "Communicate clear disinflation path and central bank independence to rating agencies."})
            elif r['CPI'] <= 3:
                signals.append({"pillar": "Monetary", "type": "strength", "metric": "CPI Inflation",
                    "value": f"{r['CPI']:.1f}%",
                    "message": "At or below 3% — price stability achieved; supports score 2 (best tier).",
                    "action": "Reinforce inflation targeting credibility narrative; highlight forward guidance effectiveness."})

            if r['Fin_Depth'] < 20:
                signals.append({"pillar": "Monetary", "type": "weakness", "metric": "Financial Depth",
                    "value": f"{r['Fin_Depth']:.1f}% GDP",
                    "message": "Below 20% — shallow financial system limits monetary policy transmission effectiveness.",
                    "action": "Capital market deepening agenda: domestic bond market, pension fund reform, banking sector expansion."})
            elif r['Fin_Depth'] < 40:
                signals.append({"pillar": "Monetary", "type": "moderate", "metric": "Financial Depth",
                    "value": f"{r['Fin_Depth']:.1f}% GDP",
                    "message": "Between 20%–40% — moderate depth; below S&P's 40% bonus threshold.",
                    "action": "Demonstrate financial deepening progress; highlight credit growth quality not just quantity."})
            else:
                signals.append({"pillar": "Monetary", "type": "strength", "metric": "Financial Depth",
                    "value": f"{r['Fin_Depth']:.1f}% GDP",
                    "message": "Above 40% — deep financial market; supports best-tier monetary score.",
                    "action": "Leverage financial depth as monetary flexibility argument in QO discussions."})

            # Institutional signals
            if pd.notna(r['WGI_Score']):
                if r['WGI_Score'] < 40:
                    signals.append({"pillar": "Institutional", "type": "weakness", "metric": "WGI Score",
                        "value": f"{r['WGI_Score']:.1f}",
                        "message": "Below 40 — weak governance profile; institutional score near bottom tier.",
                        "action": "Anti-corruption reforms, rule of law improvements, and transparency agenda are prerequisite for upgrade."})
                elif r['WGI_Score'] < 60:
                    signals.append({"pillar": "Institutional", "type": "moderate", "metric": "WGI Score",
                        "value": f"{r['WGI_Score']:.1f}",
                        "message": "Between 40–60 — mid-tier governance; significant room for improvement.",
                        "action": "Governance reform narrative: highlight specific WGI sub-indicator improvements (e.g. control of corruption, rule of law)."})
                else:
                    signals.append({"pillar": "Institutional", "type": "strength", "metric": "WGI Score",
                        "value": f"{r['WGI_Score']:.1f}",
                        "message": "Above 60 — strong governance; institutional anchor for investment grade.",
                        "action": "Maintain and communicate institutional quality as rating differentiator vs peers."})

            # ── QO NARRATIVE OPPORTUNITIES ────────────────────────────
            qo_opportunities = []

            if r['Debt_GDP'] <= 45 and r['Int_Rev'] <= 15:
                qo_opportunities.append({
                    "factor": "Fiscal Flexibility",
                    "argument": f"Low debt-to-GDP ({r['Debt_GDP']:.1f}%) combined with affordable interest burden ({r['Int_Rev']:.1f}% of revenue) demonstrates superior fiscal flexibility vs IG peers. Constitutional deficit ceiling provides a credible fiscal anchor.",
                    "impact": "+1 notch potential"
                })

            if r['Reserves'] > 6 and r['GEFN'] < 75:
                qo_opportunities.append({
                    "factor": "External Liquidity & IIP",
                    "argument": f"Reserve coverage of {r['Reserves']:.1f} months exceeds S&P's 6-month adequacy threshold while GEFN of {r['GEFN']:.1f}% remains manageable. This combination demonstrates resilience beyond what quantitative metrics alone capture.",
                    "impact": "+1 notch potential"
                })

            if r['CPI'] <= 3 and r['Fin_Depth'] > 40:
                qo_opportunities.append({
                    "factor": "Monetary Flexibility",
                    "argument": f"Inflation at {r['CPI']:.1f}% within target band combined with financial depth of {r['Fin_Depth']:.1f}% GDP demonstrates effective monetary transmission and central bank credibility beyond peer average.",
                    "impact": "+1 notch potential"
                })

            if pd.notna(r['WGI_Score']) and r['WGI_Score'] > 60:
                qo_opportunities.append({
                    "factor": "Political & Governance Risks",
                    "argument": f"WGI score of {r['WGI_Score']:.1f} signals governance quality above the median for this rating tier. Democratic stability and institutional continuity reduce political transition risk premium.",
                    "impact": "Mitigate negative QO"
                })

            if r['Growth'] > 4 and r['Debt_GDP'] < 60:
                qo_opportunities.append({
                    "factor": "Contingent Liabilities",
                    "argument": f"Strong GDP growth ({r['Growth']:.1f}%) combined with contained debt ({r['Debt_GDP']:.1f}% GDP) reduces the probability of contingent liability crystallization — SOE/GRE distress is less likely in a high-growth environment.",
                    "impact": "Mitigate negative QO"
                })

            # ============================================================
            # LAYOUT
            # ============================================================

            # ── Section 1: Pillar Overview ─────────────────────────────
            st.markdown("### 🗺️ Pillar Strength Overview")
            ov_cols = st.columns(5)
            pillar_icons = {"Institutional": "🏛", "Economic": "📈", "Fiscal": "💰", "External": "🌐", "Monetary": "🏦"}
            for col, (name, data) in zip(ov_cols, pillars.items()):
                score = data["score"]
                strength_pct = (1 - (score - 1) / 5) * 100
                if score <= 2.5:
                    bg, fg, label = "#d1fae5", "#065f46", "Strong"
                elif score <= 3.5:
                    bg, fg, label = "#fef3c7", "#78350f", "Moderate"
                else:
                    bg, fg, label = "#fee2e2", "#991b1b", "Weak"
                col.markdown(f"""
                <div style="text-align:center; padding:14px; border-radius:10px; background:{bg}; border: 1px solid {fg}22;">
                    <div style="font-size:22px;">{pillar_icons[name]}</div>
                    <div style="font-size:12px; color:#555; margin:4px 0;">{name}</div>
                    <div style="font-size:26px; font-weight:800; color:{fg};">{score:.1f}</div>
                    <div style="font-size:11px; color:{fg}; font-weight:600;">{label}</div>
                    <div style="font-size:11px; color:#888;">{strength_pct:.0f}% strength</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("")

            # ── Section 2: Strengths / Weaknesses / Moderate ──────────
            r1c1, r1c2, r1c3 = st.columns(3)

            with r1c1:
                st.markdown("### ✅ Strengths")
                strength_signals = [s for s in signals if s["type"] == "strength"]
                if strength_signals:
                    for sig in strength_signals:
                        st.markdown(f"""
                        <div style="padding:12px; border-radius:8px; background:#d1fae5; border-left:4px solid #10b981; margin-bottom:10px;">
                            <div style="font-size:12px; color:#065f46; font-weight:700;">{pillar_icons.get(sig['pillar'], '📌')} {sig['pillar']} · {sig['metric']}</div>
                            <div style="font-size:13px; color:#065f46; font-weight:600; margin:4px 0;">{sig['value']}</div>
                            <div style="font-size:12px; color:#047857;">{sig['message']}</div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.info("No strong pillars identified based on current data.")

            with r1c2:
                st.markdown("### ⚠️ Moderate Areas")
                moderate_signals = [s for s in signals if s["type"] == "moderate"]
                if moderate_signals:
                    for sig in moderate_signals:
                        st.markdown(f"""
                        <div style="padding:12px; border-radius:8px; background:#fef3c7; border-left:4px solid #f59e0b; margin-bottom:10px;">
                            <div style="font-size:12px; color:#78350f; font-weight:700;">{pillar_icons.get(sig['pillar'], '📌')} {sig['pillar']} · {sig['metric']}</div>
                            <div style="font-size:13px; color:#78350f; font-weight:600; margin:4px 0;">{sig['value']}</div>
                            <div style="font-size:12px; color:#92400e;">{sig['message']}</div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.success("No moderate-risk areas identified.")

            with r1c3:
                st.markdown("### ❌ Weaknesses")
                weakness_signals = [s for s in signals if s["type"] == "weakness"]
                if weakness_signals:
                    for sig in weakness_signals:
                        st.markdown(f"""
                        <div style="padding:12px; border-radius:8px; background:#fee2e2; border-left:4px solid #ef4444; margin-bottom:10px;">
                            <div style="font-size:12px; color:#991b1b; font-weight:700;">{pillar_icons.get(sig['pillar'], '📌')} {sig['pillar']} · {sig['metric']}</div>
                            <div style="font-size:13px; color:#991b1b; font-weight:600; margin:4px 0;">{sig['value']}</div>
                            <div style="font-size:12px; color:#b91c1c;">{sig['message']}</div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.success("No critical weaknesses identified.")

            st.divider()

            # ── Section 3: Action Items ────────────────────────────────
            st.markdown("### 🎯 Priority Action Items")
            st.caption("Ranked by pillar impact — addressing these moves the needle most on the S&P SRM score.")

            # Sort: weaknesses first, then moderate, then strengths to maintain
            priority_order = {"weakness": 0, "moderate": 1, "strength": 2}
            sorted_signals = sorted(signals, key=lambda x: priority_order[x["type"]])

            for i, sig in enumerate(sorted_signals, 1):
                if sig["type"] == "weakness":
                    icon, color = "🔴", "#fee2e2"
                elif sig["type"] == "moderate":
                    icon, color = "🟡", "#fef9c3"
                else:
                    icon, color = "🟢", "#f0fdf4"

                with st.expander(f"{icon} {i}. [{sig['pillar']}] {sig['metric']} — {sig['value']}"):
                    st.markdown(f"**Assessment:** {sig['message']}")
                    st.markdown(f"**Recommended Action:** {sig['action']}")

            st.divider()

            # ── Section 4: QO Narrative Opportunities ─────────────────
            st.markdown("### 🗣️ QO Narrative Opportunities")
            st.markdown("Arguments that can be made in the S&P rating committee dialogue to earn or protect QO points:")

            if qo_opportunities:
                for opp in qo_opportunities:
                    impact_color = "#10b981" if "+" in opp["impact"] else "#6366f1"
                    st.markdown(f"""
                    <div style="padding:16px; border-radius:10px; background:#f0f4ff; border-left:5px solid {impact_color}; margin-bottom:14px;">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <div style="font-size:14px; font-weight:700; color:#1e3a8a;">⚖️ {opp['factor']}</div>
                            <div style="font-size:12px; font-weight:700; color:{impact_color}; background:{impact_color}22; padding:3px 10px; border-radius:20px;">{opp['impact']}</div>
                        </div>
                        <div style="font-size:13px; color:#374151; margin-top:8px; line-height:1.6;">{opp['argument']}</div>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.warning("Based on current data, limited QO narrative opportunities are available. Focus on addressing weaknesses first to unlock positive QO arguments.")

            st.divider()

            # ── Section 5: Rating Trajectory Summary ──────────────────
            st.markdown("### 📈 Rating Trajectory Summary")
            traj_col1, traj_col2, traj_col3 = st.columns(3)

            upgrade_count  = len([s for s in signals if s["type"] == "strength"])
            moderate_count = len([s for s in signals if s["type"] == "moderate"])
            weakness_count = len([s for s in signals if s["type"] == "weakness"])

            with traj_col1:
                st.markdown(f"""
                <div style="text-align:center; padding:16px; border-radius:10px; background:#d1fae5;">
                    <div style="font-size:32px; font-weight:800; color:#065f46;">{upgrade_count}</div>
                    <div style="font-size:13px; color:#047857;">Strong Signals</div>
                    <div style="font-size:11px; color:#6ee7b7;">Support current or higher rating</div>
                </div>
                """, unsafe_allow_html=True)

            with traj_col2:
                st.markdown(f"""
                <div style="text-align:center; padding:16px; border-radius:10px; background:#fef3c7;">
                    <div style="font-size:32px; font-weight:800; color:#78350f;">{moderate_count}</div>
                    <div style="font-size:13px; color:#92400e;">Watch Areas</div>
                    <div style="font-size:11px; color:#fcd34d;">Monitor for deterioration</div>
                </div>
                """, unsafe_allow_html=True)

            with traj_col3:
                st.markdown(f"""
                <div style="text-align:center; padding:16px; border-radius:10px; background:#fee2e2;">
                    <div style="font-size:32px; font-weight:800; color:#991b1b;">{weakness_count}</div>
                    <div style="font-size:13px; color:#b91c1c;">Weak Signals</div>
                    <div style="font-size:11px; color:#fca5a5;">Downgrade pressure if unaddressed</div>
                </div>
                """, unsafe_allow_html=True)

            # Overall trajectory verdict
            st.markdown("")
            if weakness_count == 0 and upgrade_count >= 3:
                verdict_color, verdict_bg, verdict = "#065f46", "#d1fae5", "🚀 Upgrade Candidate — Strong fundamental profile with multiple positive signals. Build QO narrative."
            elif weakness_count >= 3:
                verdict_color, verdict_bg, verdict = "#991b1b", "#fee2e2", "⚠️ Downgrade Risk — Multiple structural weaknesses require urgent policy attention."
            elif weakness_count >= 1 and moderate_count >= 2:
                verdict_color, verdict_bg, verdict = "#78350f", "#fef3c7", "➡️ Stable with Caution — Rating defensible but vulnerable to shocks. Address weaknesses to prevent negative outlook."
            else:
                verdict_color, verdict_bg, verdict = "#1e3a8a", "#f0f4ff", "✅ Broadly Stable — Mixed profile; maintain strengths and work on moderate areas for upgrade potential."

            st.markdown(f"""
            <div style="padding:16px; border-radius:10px; background:{verdict_bg}; border-left:6px solid {verdict_color}; margin-top:8px;">
                <div style="font-size:15px; font-weight:700; color:{verdict_color};">{verdict}</div>
            </div>
            """, unsafe_allow_html=True)

        with tabs[3]:  # BRIEFING NOTES
            st.markdown("## 📄 Briefing Notes Generator")
            st.markdown(
                "Generate a professional briefing note covering National Portfolio, "
                "Peer Comparison, Recommendations, and QO targeting. "
                "Download as **PDF** or **PowerPoint**."
            )

            # ── Config ────────────────────────────────────────────────────────
            bn_col1, bn_col2 = st.columns(2)
            with bn_col1:
                st.markdown("**Countries included in peer section:**")
                # Reuse sel_nations from tab[1] if already set, else just target
                try:
                    briefing_nations = sel_nations if sel_nations else [target]
                except:
                    briefing_nations = [target]
                st.write(", ".join(briefing_nations))

            with bn_col2:
                st.markdown("**Trend metrics included:**")
                try:
                    briefing_metrics = selected_metrics if selected_metrics else []
                except:
                    briefing_metrics = []
                st.write(", ".join(briefing_metrics) if briefing_metrics else "None selected in Peer Comparison tab")

            st.info(
                "💡 To include trend charts and peer metrics, first visit "
                "**📉 Peer Comparison → Historical Trends** and select countries & indicators."
            )
            st.divider()

            # ── Build comp_list for briefing ──────────────────────────────────
            try:
                briefing_comp_list = comp_list   # from tabs[1]
            except NameError:
                # fallback: just target
                s_inst_b = max(1.0, min(6.0, 6 - (float(r['WGI_Score'])/20))) if pd.notna(r['WGI_Score']) else 3.5
                s_eco_b  = score_economic(r['GDP_PC'], r['Growth'], "Standard")
                s_fis_b  = score_fiscal(r['Debt_GDP'], r['Int_Rev'], r['Balance'], "Neutral")[0]
                s_ext_b  = score_external(r['GEFN'], r['NIIP_CAR'], 20, r['Reserves'])[0]
                s_mon_b  = score_monetary("Floating", "High", r['CPI'], r['Fin_Depth'])
                ie_b = (s_inst_b + s_eco_b) / 2
                fp_b = (s_fis_b + s_ext_b + s_mon_b) / 3
                srm_b = get_indicative_rating(ie_b, fp_b)
                qo_b = RATING_TO_NUM.get(r['Actual_Rating'], 8) - RATING_TO_NUM.get(srm_b, 8)
                briefing_comp_list = [{
                    "Country": target,
                    "🏛 Institutional": s_inst_b, "📈 Economic": s_eco_b,
                    "💰 Fiscal": s_fis_b, "🌐 External": s_ext_b, "🏦 Monetary": s_mon_b,
                    "_ie": ie_b, "_fp": fp_b, "_srm": srm_b, "_qo": qo_b, "_nr": r,
                }]

            # trend data
            try:
                briefing_trend_df = trend_df
            except NameError:
                briefing_trend_df = None

            # ── Download buttons ──────────────────────────────────────────────
            dl1, dl2 = st.columns(2)

            with dl1:
                st.markdown("### 📑 PDF Briefing Note")
                st.markdown("Professional multi-page report with tables, charts, and narrative.")
                if st.button("⚙️ Generate PDF", use_container_width=True):
                    with st.spinner("Generating PDF..."):
                        pdf_buf = generate_pdf(
                            target=target, r=r,
                            srm_rating=srm_rating, qo=qo,
                            s_inst=s_inst, s_eco=s_eco,
                            s_fis=s_fis, s_ext=s_ext, s_mon=s_mon,
                            prof_ie=prof_ie, prof_fp=prof_fp,
                            comp_list=briefing_comp_list,
                            sel_nations=[c["Country"] for c in briefing_comp_list],
                            trend_df=briefing_trend_df,
                            selected_metrics=briefing_metrics,
                            signals=signals,
                            qo_opportunities=qo_opportunities,
                        )
                    st.download_button(
                        label="⬇️ Download PDF",
                        data=pdf_buf,
                        file_name=f"Briefing_{target}_{datetime.now().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

            with dl2:
                st.markdown("### 📊 PowerPoint Deck")
                st.markdown("Editable slide deck with charts, tables, and QO narrative.")
                if st.button("⚙️ Generate PowerPoint", use_container_width=True):
                    with st.spinner("Generating PowerPoint..."):
                        pptx_buf = generate_pptx(
                            target=target, r=r,
                            srm_rating=srm_rating, qo=qo,
                            s_inst=s_inst, s_eco=s_eco,
                            s_fis=s_fis, s_ext=s_ext, s_mon=s_mon,
                            prof_ie=prof_ie, prof_fp=prof_fp,
                            comp_list=briefing_comp_list,
                            sel_nations=[c["Country"] for c in briefing_comp_list],
                            trend_df=briefing_trend_df,
                            selected_metrics=briefing_metrics,
                            signals=signals,
                            qo_opportunities=qo_opportunities,
                        )
                    st.download_button(
                        label="⬇️ Download PowerPoint",
                        data=pptx_buf,
                        file_name=f"Briefing_{target}_{datetime.now().strftime('%Y%m%d')}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )

        with tabs[4]: # METHODOLOGY SIMULATOR (GAUGE STYLE)
            st.subheader("Interactive Stress-Test Simulator")
            st.info("Gunakan input angka di bawah untuk mensimulasikan dampak perubahan indikator terhadap rating indikatif.")
            
            sc1, sc2, sc3 = st.columns(3)
            
            # Input Angka Presisi (Tanpa Slider untuk akurasi)
            sc1.markdown("### Economic & Institutional")
            sim_gdp = sc1.number_input("GDP Per Capita (USD)", value=float(r['GDP_PC']), step=100.0)
            sim_growth = sc1.number_input("Trend Real GDP Growth (%)", value=float(r['Growth']), step=0.1)
            sim_div = sc1.selectbox("Economic Diversification", ["High", "Standard", "Low"], index = 1)
            sim_wgi = sc1.number_input("WGI Score (Governance)", value=float(r['WGI_Score']) if pd.notna(r['WGI_Score']) else 50.0, step=1.0)

            sc2.markdown("### Fiscal & Debt")
            sim_bal = sc2.number_input("Fiscal Balance (% GDP)", value=float(r['Balance']), step=0.1)
            sim_debt = sc2.number_input("Debt-to-GDP (%)", value=float(r['Debt_GDP']), step=0.5)
            sim_int = sc2.number_input("Interest-to-Revenue (%)", value=float(r['Int_Rev']), step=0.1)
            sim_flex = sc2.selectbox("Revenue Flexibility", ["High", "Neutral", "Low"], index=1)

            sc3.markdown("### External & Monetary")
            sim_gefn = sc3.number_input("GEFN (% of CAR)", value=float(r['GEFN']), step=1.0)
            sim_niip = sc3.number_input("NIIP (% of GDP)", value=float(r['NIIP_CAR']), step=1.0)
            sim_cpi = sc3.number_input("Inflation Trends (%)", value=float(r['CPI']), step=0.1)
            sim_depth = sc3.number_input("Financial Depth (0-100)", value=int(r['Fin_Depth']), step=1)
            sim_regime = sc3.selectbox("FX Regime", ["Floating", "Fixed/Managed"])

            # Kalkulasi Ulang Hasil Simulasi
            p_inst_sim = max(1.0, min(6.0, 6 - (sim_wgi / 20)))
            p_eco_sim = score_economic(sim_gdp, sim_growth, sim_div)
            p_fis_sim, _, _ = score_fiscal(sim_debt, sim_int, sim_bal, sim_flex)
            p_ext_sim, _, _ = score_external(sim_gefn, sim_niip, 20, r['Reserves'])
            p_mon_sim = score_monetary(sim_regime, "High", sim_cpi, sim_depth)

            # Hasil Profil Gabungan Simulasi
            res_ie = (p_inst_sim + p_eco_sim) / 2
            res_fp = (p_fis_sim + p_ext_sim + p_mon_sim) / 3
            sim_rating = get_indicative_rating(res_ie, res_fp)
            
            # Visualisasi Hasil (Progress Bar sebagai Mock Gauge)
            rating_val = RATING_TO_NUM.get(sim_rating, 0)
            progress_val = rating_val / 16.0 # Skala 0 hingga AAA(16)
            
            st.write("---")
            st.markdown(f"<h2 style='text-align: center;'>Simulated Rating Output: <span style='color: #10B981;'>{sim_rating}</span></h2>", unsafe_allow_html=True)
            st.progress(progress_val)
            
            ic1, ic2 = st.columns(2)
            ic1.info(f"**Institutional & Economic Profile:** {res_ie:.2f}")
            ic2.info(f"**Flexibility & Performance Profile:** {res_fp:.2f}")
else:
    st.info("👋 Welcome! Please upload the S&P Macro dataset to begin. Reference WGI data will be loaded automatically from the cloud.")