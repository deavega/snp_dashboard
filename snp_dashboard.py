import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from briefing_generator import generate_pdf, generate_pptx
import requests
import io
import urllib.parse
from datetime import datetime, timezone, timedelta

def now_jakarta():
    """Returns current time in Jakarta timezone (UTC+7)."""
    return datetime.now(timezone.utc) + timedelta(hours=7)

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
    """Updated 2024 GDP thresholds per S&P Appendix D Oct 2024."""
    if gdp_pc > 48700: s = 1
    elif gdp_pc > 34600: s = 2
    elif gdp_pc > 20500: s = 3
    elif gdp_pc > 7000:  s = 4
    elif gdp_pc > 1400:  s = 5
    else: s = 6
    # Growth adjustment — tier-specific benchmark (Table 11)
    benchmark = 0.9 if s <= 2 else (2.0 if s == 3 else 2.3)
    if growth > benchmark * 1.5:    s = max(1, s - 1)
    elif growth < benchmark * 0.5:  s = min(6, s + 1)
    # Diversification
    if diversification == "High":   s = max(1, s - 1)
    if diversification == "Low":    s = min(6, s + 1)
    return s


def score_fiscal(debt_gdp, int_rev, balance, flex="Neutral"):
    """
    Performance from GG balance proxy (Table 5).
    Burden from Table 6 debt×interest matrix.
    """
    # Fiscal performance (balance as proxy for change in net debt)
    if balance >= 0:    perf = 1
    elif balance >= -3: perf = 2
    elif balance >= -5: perf = 4   # ← keep OLD -3/-5 boundary, skip perf=3
    else:               perf = 5

    # Debt burden — Table 6 matrix
    if int_rev <= 5:
        burd = 1 if debt_gdp < 30 else (2 if debt_gdp < 60 else (3 if debt_gdp < 80 else (4 if debt_gdp < 100 else 5)))
    elif int_rev <= 10:
        burd = 2 if debt_gdp < 30 else (3 if debt_gdp < 60 else (4 if debt_gdp < 80 else (5 if debt_gdp < 100 else 6)))
    elif int_rev <= 15:
        burd = 2 if debt_gdp < 30 else (3 if debt_gdp < 60 else (4 if debt_gdp < 80 else (5 if debt_gdp < 100 else 6)))
    else:
        burd = 4 if debt_gdp < 30 else (5 if debt_gdp < 60 else 6)

    final = (perf + burd) / 2
    if flex == "High": final = max(1, final - 0.5)
    if flex == "Low":  final = min(6, final + 0.5)
    return final, perf, burd


def score_external(gefn, niip, car_receipts=20, reserves_months=6.0,
                   currency="Standard"):
    """
    Aligned to S&P Table 4.
    GEFN thresholds: <50 / 50-75 / 75-100 / 100-150 / >150
    NIIP thresholds: >0 / 0 to -20 / -20 to -100 / -100 to -150 / <-150
    Reserves below 6 months → worsen liquidity by 1.
    Currency status → uplift for reserve/actively traded.
    """
    # ── Liquidity score from GEFN (Table 4 column headers) ───────────────────
    if gefn <= 50:    l_s = 1
    elif gefn <= 75:  l_s = 2
    elif gefn <= 100: l_s = 3
    elif gefn <= 150: l_s = 4
    else:             l_s = 5

    # ── Reserves adjustment ───────────────────────────────────────────────────
    # Reserves < 3 months = critically low → worsen by 1
    # Reserves < 6 months = below adequacy → no bonus, slight pressure
    # Reserves > 6 months = adequate → improve by 1
    if reserves_months < 3:   l_s = min(6, l_s + 1)
    elif reserves_months > 6: l_s = max(1, l_s - 1)

    # ── Debt position from NIIP (Table 4 row headers) ─────────────────────────
    if niip >= 0:      d_s = 1   # net creditor
    elif niip > -20:   d_s = 2   # slight net debtor
    elif niip > -100:  d_s = 3   # moderate net debtor
    elif niip > -150:  d_s = 4   # large net debtor
    else:              d_s = 5   # very large net debtor

    # ── Currency status uplift (para. 46-51) ─────────────────────────────────
    if currency == "Reserve":          bonus = -2
    elif currency == "ActivelyTraded": bonus = -1
    else:                              bonus = 0

    score = max(1, min(6, (l_s + d_s) / 2 + bonus))
    return score, l_s, d_s


def score_monetary(regime, credibility, inflation, depth):
    """
    Table 8A (FX regime, 40%) + Table 8B (credibility, 60%) per para. 111.
    Credibility thresholds aligned to Table 8B descriptions.
    """
    regime_scores = {
        "Reserve":          1,
        "Floating":         2,
        "ManagedFloat":     3,
        "Peg":              4,
        "CurrencyBoard":    5,
        "NoLocalCurrency":  6,
    }
    fx = regime_scores.get(regime, 2)

    # Credibility — Table 8B: depth threshold is 50% of GDP (combined credit+bonds)
    if inflation <= 5 and depth > 50:   cred_s = 2   # score 1 reserved for reserve currency
    elif inflation <= 8 and depth > 30: cred_s = 3
    elif inflation <= 10:               cred_s = 4
    elif inflation <= 20:               cred_s = 5
    else:                               cred_s = 6

    return round(fx * 0.4 + cred_s * 0.6, 1)

def score_supplemental(s_inst, s_ext, s_fis_burd, debt_gdp, gefn, 
                       liquid_assets_gdp=0, event_risk=False):
    """
    Computes S&P Supplemental Adjustment Factors per para. 125-128.
    Returns:
        adjustment (int): net notch adjustment (negative = downgrade, positive = upgrade)
        factors (list): list of dicts describing each factor triggered
    """
    adjustment = 0
    factors = []

    # ── NEGATIVE: Institutional cap ──────────────────────────────────────────
    # Para 126: Inst=6 → cannot be rated above BB+ (capped, not a notch adjustment)
    # Para 126: Inst=6 AND debt burden ≥ 5 → cannot be rated above B+
    # These are handled as caps later in get_indicative_rating_with_supplemental()

    # ── NEGATIVE: Extremely weak external liquidity ───────────────────────────
    # Para 126: GEFN substantially worse than table 4 worst level (>150%)
    if gefn > 200:
        adjustment -= 2
        factors.append({
            "type": "negative",
            "factor": "Extremely Weak External Liquidity",
            "detail": f"GEFN of {gefn:.1f}% is substantially above the 150% worst-tier threshold.",
            "impact": "–2 notches"
        })
    elif gefn > 150:
        adjustment -= 1
        factors.append({
            "type": "negative",
            "factor": "Very Weak External Liquidity",
            "detail": f"GEFN of {gefn:.1f}% exceeds the 150% worst-tier threshold.",
            "impact": "–1 notch"
        })

    # ── NEGATIVE: Extremely high fiscal debt burden ───────────────────────────
    # Para 126: debt burden score at worst AND deteriorating
    if s_fis_burd >= 6 and debt_gdp > 100:
        adjustment -= 1
        factors.append({
            "type": "negative",
            "factor": "Extremely High Fiscal Debt Burden",
            "detail": f"Debt burden score {s_fis_burd:.0f}/6 with debt-to-GDP of {debt_gdp:.1f}% — "
                      f"significantly worse than worst-tier benchmark.",
            "impact": "–1 notch"
        })

    # ── NEGATIVE: Event risk ──────────────────────────────────────────────────
    # Para 126: imminent/rapidly rising political risk, war, etc.
    if event_risk:
        adjustment -= 1
        factors.append({
            "type": "negative",
            "factor": "Event Risk",
            "detail": "Imminent or rapidly rising political/security risk identified.",
            "impact": "–1 notch"
        })

    # ── POSITIVE: Very large liquid financial government assets ───────────────
    # Para 128: net asset position AND liquid assets > 100% GDP → +1 notch uplift
    if liquid_assets_gdp > 100 and debt_gdp < 0:
        adjustment += 1
        factors.append({
            "type": "positive",
            "factor": "Very Large Liquid Financial Government Assets",
            "detail": f"Government in net asset position with liquid assets of "
                      f"{liquid_assets_gdp:.1f}% GDP (>100% threshold).",
            "impact": "+1 notch"
        })

    return adjustment, factors


def apply_supplemental_caps(srm_rating, s_inst, s_fis_burd, adjustment):
    """
    Applies hard rating caps from para. 126 AFTER notch adjustments.
    Returns the final capped rating and cap description if applied.
    """
    # Apply notch adjustment first
    base_num = RATING_TO_NUM.get(srm_rating, 8)
    adjusted_num = max(0, min(16, base_num + adjustment))

    # Reverse lookup
    NUM_TO_RATING = {v: k for k, v in RATING_TO_NUM.items()}
    adjusted_rating = NUM_TO_RATING.get(adjusted_num, srm_rating)

    cap_applied = None

    # Cap 1: Institutional score = 6 → cannot exceed BB+ (num=6)
    if s_inst >= 6:
        if adjusted_num > 6:
            adjusted_num = 6
            adjusted_rating = "BB+"
            cap_applied = "Institutional score of 6 caps rating at BB+"

    # Cap 2: Institutional score = 6 AND debt burden score ≥ 5 → cannot exceed B+ (num=3)
    if s_inst >= 6 and s_fis_burd >= 5:
        if adjusted_num > 3:
            adjusted_num = 3
            adjusted_rating = "B+"
            cap_applied = "Institutional score 6 + debt burden ≥ 5 caps rating at B+"

    return adjusted_rating, cap_applied

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

        def f_c(df, kw):
            # Try latest estimate year (e.g. 2025e) first, fall back to 2024
            candidates = [c for c in df.columns if kw in c]
            # Prefer estimate years (ending with 'e')
            est_cols = [c for c in candidates if c.endswith('e')]
            if est_cols:
                # Sort and take the latest estimate year
                return sorted(est_cols)[-1]
            # Fall back to 2024 if no estimate year found
            hist_cols = [c for c in candidates if "2024" in c]
            if hist_cols:
                return hist_cols[0]
            # Last resort: take any matching column
            return candidates[0] if candidates else None

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
PASSCODE = st.secrets.get("APP_PASSCODE", "Dspp123#")

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
# ── Background image header ───────────────────────────────────────────────────
import base64

def get_base64_image(image_path):
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

HEADER_IMG_PATH = os.path.join(os.path.dirname(__file__), "assets", "background.png")

if os.path.exists(HEADER_IMG_PATH):
    img_b64 = get_base64_image(HEADER_IMG_PATH)
    st.markdown(f"""
    <div style="
        background-image: url('data:image/png;base64,{img_b64}');
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        border-radius: 12px;
        padding: 48px 24px;
        margin-bottom: 16px;
        position: relative;
    ">
        <!-- Dark overlay for text readability -->
        <div style="
            position: absolute;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(15, 30, 80, 0.55);
            border-radius: 12px;
        "></div>
        <!-- Text on top of overlay -->
        <div style="position: relative; z-index: 1; text-align: center;">
            <h1 style="color: #FFFFFF; font-size: 2.4rem; font-weight: 800;
                       margin: 0; text-shadow: 0 2px 8px rgba(0,0,0,0.4);">
                Sovereign Rating Monitoring
            </h1>
            <h4 style="color: #CADCFC; font-size: 1.1rem; font-weight: 400;
                       margin: 8px 0 0 0; text-shadow: 0 1px 4px rgba(0,0,0,0.3);">
                Based on S&amp;P Global Methodology
            </h4>
        </div>
    </div>
    <div style="text-align:right; font-size:10px; color:#94A3B8; margin-top:3px; margin-bottom:8px;">
        Photo by <a href="https://unsplash.com/id/@spensersembrat" target="_blank" 
        style="color:#94A3B8;">Spenser Sembrat</a> on 
        <a href="https://unsplash.com" target="_blank" style="color:#94A3B8;">Unsplash</a>
    </div>
    """, unsafe_allow_html=True)
else:
    # Fallback if image not found

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

# ── Methodology Reference ─────────────────────────────────────────────────────
st.markdown("""
<div style="background:#EFF6FF; border-left:4px solid #1E3A8A; border-radius:6px;
            padding:10px 16px; margin-bottom:16px; font-size:13px; color:#1E3A8A;">
    📖 <b>Methodology Reference:</b> This dashboard implements 
    <a href="https://www.spglobal.com/ratings/en/regulatory/article/-/view/sourceId/10221157" 
       target="_blank" style="color:#1E3A8A; font-weight:600;">
       S&P Global Ratings — Sovereign Rating Methodology
    </a>
    (Dec 2017, updated Oct 2024). All pillar scores, matrix lookups, and supplemental 
    adjustment factors follow this criteria document. Scores are model-derived and do not 
    constitute official S&P ratings.
</div>
""", unsafe_allow_html=True)

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

        # ── Load trend data once — available to ALL tabs ──────────────────────
        with st.spinner("Extracting trend data..."):
            if hasattr(f_macro, 'getvalue'):
                xls_bytes = f_macro.getvalue()   # Streamlit UploadedFile
            else:
                f_macro.seek(0)
                xls_bytes = f_macro.read()       # BytesIO from default file
            if hasattr(f_macro, 'seek'):
                f_macro.seek(0)                  # reset for any subsequent reads
            trend_df = extract_trend_data(xls_bytes)

        with tabs[0]: # NATION ANALYSIS & OVERLAY
            target = st.selectbox("Select Sovereign Target", df['Country'].unique(), 
                                   index=list(df['Country']).index('Indonesia') if 'Indonesia' in df['Country'].values else 0)
            r = df[df['Country'] == target].iloc[0]
            
            # Kalkulasi Skor pilar berdasarkan logika terkalibrasi di Chunk 1
            s_inst = max(1.0, min(6.0, 6 - (float(r['WGI_Score']) / 20))) if pd.notna(r['WGI_Score']) else 3.5
            s_eco = score_economic(r['GDP_PC'], r['Growth'], "Standard")
            s_fis, s_fis_perf, s_fis_burd = score_fiscal(r['Debt_GDP'], r['Int_Rev'], r['Balance'], "Neutral")
            s_ext, s_liq, s_debt_pos = score_external(r['GEFN'], r['NIIP_CAR'], 20, r['Reserves'])
            s_mon = score_monetary("Floating", "High", r['CPI'], r['Fin_Depth'])

            # Calculate metrix scores and initial SRM rating
            prof_ie = (s_inst + s_eco) / 2
            prof_fp = (s_fis + s_ext + s_mon) / 3
            srm_rating = get_indicative_rating(prof_ie, prof_fp)

            # ── Supplemental Adjustment Factors ──────────────────────────────
            supp_adj, supp_factors = score_supplemental(
                s_inst=s_inst,
                s_ext=s_ext,
                s_fis_burd=s_fis_burd,
                debt_gdp=r['Debt_GDP'],
                gefn=r['GEFN'],
                liquid_assets_gdp=0,
                event_risk=False,
            )
            final_indicative, cap_note = apply_supplemental_caps(
                srm_rating, s_inst, s_fis_burd, supp_adj)

            # ── Residual ±1 notch (what was previously called QO) ────────────
            actual_val   = RATING_TO_NUM.get(str(r['Actual_Rating']).replace('*','').strip(), 8)
            final_val    = RATING_TO_NUM.get(final_indicative, 8)
            srm_val      = RATING_TO_NUM.get(srm_rating, 8)
            residual_qo  = actual_val - final_val
            qo           = residual_qo  # keep alias so downstream code doesn't break

            

            # ── KPI metrics ──────────────────────────────────────────────────
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("SRM Matrix Output", srm_rating)
            m2.metric("After Supplemental",
                      final_indicative,
                      delta=f"{int(supp_adj):+} notch" if supp_adj != 0 else "No adj.")
            m3.metric("Official S&P Rating", r['Actual_Rating'])
            m4.metric("Residual ±1 Notch",
                      f"{int(residual_qo):+}",
                      delta=int(residual_qo), delta_color="normal")
            m5.metric("Supplemental Factors",
                      f"{len(supp_factors)} triggered",
                      delta="⚠️ Cap applied" if cap_note else "No cap")
            
            # ── SUPPLEMENTAL ADJUSTMENT FACTORS DISPLAY ───────────────────────
            st.markdown("#### ⚡ Supplemental Adjustment Factors")
            st.caption(
                "Applied **after** the SRM matrix per S&P methodology para. 125–128. "
                "Represents extreme conditions and hard rating caps."
            )

            if not supp_factors and not cap_note:
                st.success(
                    "✅ No supplemental adjustment factors triggered. "
                    "Final indicative rating equals SRM matrix output."
                )
            else:
                for f in supp_factors:
                    color  = "#fee2e2" if f["type"] == "negative" else "#d1fae5"
                    border = "#ef4444" if f["type"] == "negative" else "#10b981"
                    icon   = "📉" if f["type"] == "negative" else "📈"
                    st.markdown(f"""
                    <div style="padding:12px;border-radius:8px;background:{color};
                                border-left:5px solid {border};margin-bottom:8px;">
                        <b>{icon} {f['factor']}</b>
                        &nbsp;<span style="font-size:12px;font-weight:bold;
                        color:{'#991b1b' if f['type']=='negative' else '#065f46'};">
                        [{f['impact']}]</span><br>
                        <span style="font-size:13px;">{f['detail']}</span>
                    </div>
                    """, unsafe_allow_html=True)

                if cap_note:
                    st.markdown(f"""
                    <div style="padding:12px;border-radius:8px;background:#fef3c7;
                                border-left:5px solid #f59e0b;margin-bottom:8px;">
                        <b>🚧 Hard Rating Cap Applied</b><br>
                        <span style="font-size:13px;">{cap_note}</span>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown(f"""
            <div style="padding:12px;background:#f0f4ff;border-radius:8px;
                        text-align:center;margin-top:4px;margin-bottom:8px;">
                <b>SRM Matrix:</b> {srm_rating}
                &nbsp;→&nbsp;
                <b>After Supplemental ({int(supp_adj):+}):</b> {final_indicative}
                &nbsp;→&nbsp;
                <b>Official:</b> {r['Actual_Rating']}
                &nbsp;|&nbsp;
                <b>Residual ±1 notch:</b> {int(residual_qo):+}
            </div>
            """, unsafe_allow_html=True)

            # ── Methodology Plain Language Explainer ──────────────────────────
            with st.expander("📖 How does S&P rate a sovereign?", expanded=False):
                st.markdown(f"""
                #### How S&P Arrives at a Sovereign Rating

                S&P rates a country's ability and willingness to repay its debt to commercial creditors.
                The process has **three layers**, applied in sequence:

                ---

                **Layer 1 — The Sovereign Rating Model (SRM Matrix)**

                S&P scores five pillars on a scale of 1 (strongest) to 6 (weakest):

                | Pillar | What it measures | Weight |
                |--------|-----------------|--------|
                | 🏛 Institutional | Quality of governance, rule of law, policy predictability | 25% |
                | 📈 Economic | GDP per capita, growth prospects, economic diversity | 25% |
                | 💰 Fiscal | Deficit, debt burden, interest costs vs revenues | 16.7% |
                | 🌐 External | Ability to earn foreign currency, reserve adequacy, net debt position | 16.7% |
                | 🏦 Monetary | Central bank credibility, exchange rate flexibility, inflation | 16.7% |

                The five scores combine into two profiles:
                - **IE Profile** = (Institutional + Economic) ÷ 2 → currently **{prof_ie:.2f}**
                - **FP Profile** = (Fiscal + External + Monetary) ÷ 3 → currently **{prof_fp:.2f}**

                These two numbers are looked up in a **6×11 matrix** to produce the indicative rating → **{srm_rating}**
                """)

                # ── SRM Matrix image ──────────────────────────────────────────
                import pathlib
                # Try multiple possible paths
                _possible_paths = [
                    pathlib.Path(__file__).parent / "assets" / "srm_matrix.png",
                    pathlib.Path("assets/srm_matrix.png"),
                    pathlib.Path("./assets/srm_matrix.png"),
                ]
                _matrix_img = None
                for _p in _possible_paths:
                    if _p.exists():
                        _matrix_img = str(_p)
                        break

                if _matrix_img:
                    col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
                    with col_img2:
                        st.image(_matrix_img,
                                 caption="S&P Sovereign Rating Model (SRM) — Indicative Rating Matrix © S&P Global Ratings 2017",
                                 use_container_width=True)
                else:
                    # Fallback: load from base64 if file not found via path
                    st.caption(f"_Matrix image not found. Searched: {[str(p) for p in _possible_paths]}_")

                st.markdown(f"""

                ---

                **Layer 2 — Supplemental Adjustment Factors (para. 125–128)**

                After the matrix, S&P checks for *extreme* conditions that can override the model:
                - Extremely high external financing needs → downgrade
                - Extremely high debt burden → downgrade
                - Institutional score = 6 → hard cap at BB+
                - Very large liquid government assets (>100% GDP) → upgrade

                {'✅ No supplemental factors triggered for ' + target + ' — SRM output carries through unchanged.' if not supp_factors and not cap_note else
                 '⚠️ Supplemental factor triggered: ' + (', '.join([f["factor"] for f in supp_factors]) if supp_factors else cap_note or '')}

                Post-supplemental indicative rating → **{final_indicative}**

                ---

                **Layer 3 — Residual ±1 Notch Adjustment (para. 15)**

                The rating committee can nudge the rating **up or down by one notch** based on qualitative judgment:
                - Is the country a sustained outperformer vs peers? → +1 notch
                - Are there hidden risks not captured by the numbers? → −1 notch
                - Are there transitional factors (e.g. new resource discovery, reform momentum)? → ±1 notch

                For **{target}**: residual adjustment = **{int(qo):+} notch** → Official rating **{r['Actual_Rating']}**

                ---

                **In simple terms:**
                > Think of it like a school report card. The SRM matrix is the exam score.
                > Supplemental factors are the teacher saying *"this student has a serious discipline issue"* or
                > *"this student has exceptional talent"*. The residual adjustment is the principal's final note
                > on the report — based on overall impression, trajectory, and things the exam didn't fully capture.

                📖 Full methodology: [S&P Sovereign Rating Criteria](https://www.spglobal.com/ratings/en/regulatory/article/-/view/sourceId/10221157)
                """)    

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

                    qo_label = ("Upgrade" if residual_qo > 0 else ("Penalty" if residual_qo < 0 else "Neutral"))
                    qo_color = "#10b981" if qo > 0 else ("#ef4444" if qo < 0 else "#6b7280")

                    st.markdown("### 🗣️ Residual ±1 Notch Adjustment Opportunities")
                    st.markdown(
                "These arguments target the **residual one-notch adjustment** (para. 15) "
                "that S&P applies *after* the SRM matrix and supplemental factors. "
                "This covers transitional dynamics, ESG factors, and over/underperformance "
                "not captured elsewhere — **not** the supplemental adjustment factors above."
            )

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
                <div style="text-align:center; padding:14px; border-radius:8px; border: 1px solid #ddd;">
                    <div style="font-size:15px; color:#444; font-weight:600;">{name}</div>
                    <div style="font-size:30px; font-weight:bold; color:{color};">{score:.1f}</div>
                    <div style="font-size:13px; color:#777;">/ 6.0 · {profile}</div>
                    <div style="font-size:13px; color:{color}; font-weight:600;">{strength:.0f}% strength</div>
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

                if trend_df is None or trend_df.empty:
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
                            # Default indicators aligned to key S&P monitoring metrics
                            PREFERRED_DEFAULTS = [
                                'GDP per capita (000s $)',
                                'GG Revenues/GDP (%)',
                                'GG interest expenditure/revenues (%)',
                                'Gross GG debt/GDP (%)',
                                'Real GDP growth (%)',
                                'Usable reserves/CAPs (months)',
                            ]

                            # Only include defaults that are available in current pillar filter
                            smart_defaults = [m for m in PREFERRED_DEFAULTS if m in filtered_metrics]

                            # Fall back to first 2 if none of the preferred exist in current filter
                            if not smart_defaults:
                                smart_defaults = filtered_metrics[:2]

                            selected_metrics = st.multiselect(
                                "Select Indicators to Plot",
                                filtered_metrics,
                                default=smart_defaults,
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

                    if not selected_metrics:
                        st.info("Select at least one indicator above to see the summary table.")
                    else:
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
                                    sub2['_s']        = sub2['Year'].apply(yk)
                                    sub2['_is_hist']  = sub2['Year'].apply(
                                        lambda y: not str(y).strip().endswith('f'))

                                    # Priority: last historical/estimate (e) year
                                    # Only fall back to forecast (f) if nothing else exists
                                    hist_est = sub2[sub2['_is_hist']].sort_values('_s')
                                    if not hist_est.empty:
                                        latest = hist_est.iloc[-1]
                                    else:
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
            st.markdown("<h2 style='color:#1E3A8A;'>Sovereign Rating Recommendation</h2>", unsafe_allow_html=True)
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

            # ── Executive Summary Paragraph ───────────────────────────────────
            upgrade_count  = len([s for s in signals if s["type"] == "strength"])
            moderate_count = len([s for s in signals if s["type"] == "moderate"])
            weakness_count = len([s for s in signals if s["type"] == "weakness"])

            # Trajectory verdict
            if weakness_count == 0 and upgrade_count >= 3:
                trajectory = "upgrade candidate"
                traj_color = "#065f46"
                traj_bg    = "#d1fae5"
            elif weakness_count >= 3:
                trajectory = "facing downgrade pressure"
                traj_color = "#991b1b"
                traj_bg    = "#fee2e2"
            elif weakness_count >= 1 and moderate_count >= 2:
                trajectory = "broadly stable with caution"
                traj_color = "#78350f"
                traj_bg    = "#fef3c7"
            else:
                trajectory = "broadly stable"
                traj_color = "#1e3a8a"
                traj_bg    = "#f0f4ff"

            # Key strength and weakness labels for the summary
            top_strengths = [s['metric'] for s in signals if s['type'] == 'strength'][:3]
            top_weaknesses = [s['metric'] for s in signals if s['type'] == 'weakness'][:3]

            strength_text = (
                f"Key strengths include {', '.join(top_strengths)}."
                if top_strengths else "No material strengths identified."
            )
            weakness_text = (
                f"Key areas of concern are {', '.join(top_weaknesses)}."
                if top_weaknesses else "No material weaknesses identified."
            )

            qo_label_exec = "upgrade" if qo > 0 else ("penalty" if qo < 0 else "neutral")
            qo_text_exec  = (
                f"S&P's rating committee applied a <b>+{int(qo)}-notch residual adjustment</b> "
                f"above the post-supplemental indicative rating of <b>{final_indicative}</b>, "
                f"recognising positive factors beyond what the quantitative model captures."
                if qo > 0 else
                f"S&P's rating committee applied a <b>{int(qo)}-notch residual adjustment</b> "
                f"below the post-supplemental indicative rating of <b>{final_indicative}</b>, "
                f"reflecting hidden risks not fully captured by the quantitative model."
                if qo < 0 else
                f"The official rating of <b>{r['Actual_Rating']}</b> aligns with the "
                f"post-supplemental indicative rating of <b>{final_indicative}</b> — "
                f"no residual committee adjustment was applied."
            )

            supp_text_exec = ""
            if supp_factors:
                supp_names = ", ".join([f['factor'] for f in supp_factors])
                supp_text_exec = (
                    f" Supplemental adjustment factors were triggered: <b>{supp_names}</b>, "
                    f"resulting in a {int(supp_adj):+}-notch adjustment from the SRM output of "
                    f"<b>{srm_rating}</b>."
                )
            elif cap_note:
                supp_text_exec = f" A hard rating cap was applied: {cap_note}."
            else:
                supp_text_exec = (
                    f" No supplemental adjustment factors were triggered — "
                    f"the SRM matrix output of <b>{srm_rating}</b> carried through "
                    f"to the post-supplemental stage unchanged."
                )

            st.markdown(f"""
            <div style="padding:16px 20px; border-radius:10px; background:{traj_bg};
                        border-left:5px solid {traj_color}; margin-bottom:20px;">
                <div style="font-size:16px; color:{traj_color}; font-weight:700;
                            margin-bottom:6px; text-transform:uppercase; letter-spacing:0.5px;">
                    Rating Trajectory: {trajectory.title()}
                </div>
                <p style="font-size:15px; color:#1e293b; line-height:1.8; margin:0;">
                    <b>{target}</b> holds an official S&P rating of <b>{r['Actual_Rating']}</b>,
                    derived from an SRM matrix output of <b>{srm_rating}</b>
                    (IE Profile: {prof_ie:.2f}, FP Profile: {prof_fp:.2f}).
                    {supp_text_exec}
                    {qo_text_exec}
                    Across {len(signals)} indicator-level signals assessed,
                    <b>{upgrade_count} are strong</b>, <b>{moderate_count} are moderate</b>,
                    and <b>{weakness_count} require attention</b>.
                    {strength_text} {weakness_text}
                    {'The overall profile supports the current rating level with potential for upgrade if key weaknesses are addressed.' if trajectory == 'broadly stable' else
                     'The profile warrants close monitoring — deterioration in watch areas could trigger a negative outlook.' if trajectory == 'broadly stable with caution' else
                     'The strong fundamental profile positions this sovereign for a potential upgrade in the near to medium term.' if trajectory == 'upgrade candidate' else
                     'Sustained weakness across multiple pillars creates meaningful downgrade risk if policy correction is not forthcoming.'}
                </p>
            </div>
            """, unsafe_allow_html=True)

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

            # ── QO NARRATIVE OPPORTUNITIES ────────────────────────────────────
            
            st.markdown("### 🗣️ Qualitative Overlay (QO) — Narrative Targeting")
            st.markdown(
                "Arguments that can be made in the S&P rating committee dialogue to earn or protect "
                "residual ±1 notch adjustment points. Peer comparison uses **BBB-tier average (2025e)** "
                "as the benchmark."
            )

            # ── Build peer comparison data ────────────────────────────────────
            # Determine target's rating tier for peer group
            actual_rating_clean = str(r['Actual_Rating']).replace('*','').strip()
            target_rating_num   = RATING_TO_NUM.get(actual_rating_clean, 8)

            # BBB tier = BBB-, BBB, BBB+ (nums 7,8,9)
            # Expand to include ±1 notch from actual rating for broader peer set
            peer_rating_nums = [
                RATING_TO_NUM.get('BBB-', 7),
                RATING_TO_NUM.get('BBB',  8),
                RATING_TO_NUM.get('BBB+', 9),
            ]
            # Build a fast country→rating lookup dict to avoid repeated DataFrame filtering
            rating_lookup = {
                row['Country']: str(row['Actual_Rating']).replace('*','').strip()
                for _, row in df.iterrows()
                if pd.notna(row['Country']) and str(row['Country']).strip() not in ('nan','None','')
            }

            peer_countries = [
                c for c in rating_lookup
                if RATING_TO_NUM.get(rating_lookup.get(c, ''), -1) in peer_rating_nums
                and c != target
            ]

            # Extract 2025e values from trend_df for peer comparison
            PEER_METRICS = {
                'Real GDP Growth (%)':      ('growth',   True,   "Real GDP growth (%)"),
                'Fiscal Balance (% GDP)':   ('balance',  True,   "GG balance/GDP (%)"),
                'Debt-to-GDP (%)':          ('debt',     False,  "Net GG debt/GDP (%)"),
                'Interest/Revenue (%)':     ('int_rev',  False,  "GG interest expenditure/revenues (%)"),
                'GEFN (% CAR)':             ('gefn',     False,  "Gross ext. fin. needs/(CAR + use. res.) (%)"),
                'Reserves (months)':        ('reserves', True,   "Usable reserves/CAPs (months)"),
                'CPI Inflation (%)':        ('cpi',      False,  "CPI growth (%)"),
                'Financial Depth (% GDP)':  ('findep',   True,   "Banks' claims on resident non-gov't sector/GDP"),
            }

            # Get latest estimate (prefer 2025e, fall back to last non-f year)
            def get_latest_est(country, raw_metric):
                """Get latest estimate value from trend_df for a country/metric."""
                try:
                    sub = trend_df[
                        (trend_df['Country'] == country) &
                        (trend_df['Metric'] == raw_metric)
                    ].copy()
                    if sub.empty: return np.nan
                    sub['_s'] = sub['Year'].apply(
                        lambda y: float(str(y).replace('e','').replace('f',''))
                        if not str(y).strip().endswith('f') else -1
                    )
                    sub = sub[sub['_s'] > 0]  # exclude forecasts
                    if sub.empty: return np.nan
                    return float(sub.sort_values('_s').iloc[-1]['Value'])
                except:
                    return np.nan

            # Compute peer averages
            peer_avgs = {}
            target_vals = {}
            for label, (key, higher_better, raw_metric) in PEER_METRICS.items():
                peer_vals = []
                for pc in peer_countries:
                    v = get_latest_est(pc, raw_metric)
                    if not np.isnan(v):
                        peer_vals.append(v)
                peer_avgs[label] = np.nanmean(peer_vals) if peer_vals else np.nan
                target_vals[label] = get_latest_est(target, raw_metric)

            # ── Peer comparison table ─────────────────────────────────────────
            st.markdown(f"#### 📊 Performance vs BBB Peer Average ({len(peer_countries)} peers, 2025e)")

            if peer_countries:
                comp_rows = []
                for label, (key, higher_better, raw_metric) in PEER_METRICS.items():
                    t_val  = target_vals.get(label, np.nan)
                    p_avg  = peer_avgs.get(label, np.nan)
                    if np.isnan(t_val) or np.isnan(p_avg):
                        continue
                    diff   = t_val - p_avg
                    # Is target better than peer?
                    is_better = (diff > 0) if higher_better else (diff < 0)
                    is_worse  = (diff < 0) if higher_better else (diff > 0)
                    signal = "🟢 Above peers" if is_better else ("🔴 Below peers" if is_worse else "⚪ In line")
                    comp_rows.append({
                        "Indicator":    label,
                        f"{target}":   f"{t_val:.1f}",
                        "BBB Peer Avg": f"{p_avg:.1f}",
                        "Difference":   f"{diff:+.1f}",
                        "vs Peers":     signal,
                        "_better":      is_better,
                        "_worse":       is_worse,
                        "_diff":        abs(diff),
                        "_higher":      higher_better,
                    })

                if comp_rows:
                    comp_df_display = pd.DataFrame(comp_rows)

                    GREEN_P  = "background-color:#d1fae5;color:#065f46;font-weight:600;"
                    RED_P    = "background-color:#fee2e2;color:#991b1b;font-weight:600;"
                    NEUTRAL  = ""

                    display_cols = ["Indicator", f"{target}", "BBB Peer Avg", "Difference", "vs Peers"]
                    display_only = comp_df_display[display_cols].reset_index(drop=True)

                    def style_peer_table(df):
                        # df only has display cols — look up _better/_worse from full df
                        styles = pd.DataFrame("", index=df.index, columns=df.columns)
                        for i in df.index:
                            is_better = comp_df_display.loc[
                                comp_df_display.index[i], "_better"]
                            is_worse  = comp_df_display.loc[
                                comp_df_display.index[i], "_worse"]
                            if is_better:
                                styles.loc[i, f"{target}"] = GREEN_P
                                styles.loc[i, "Difference"] = GREEN_P
                                styles.loc[i, "vs Peers"]   = GREEN_P
                            elif is_worse:
                                styles.loc[i, f"{target}"] = RED_P
                                styles.loc[i, "Difference"] = RED_P
                                styles.loc[i, "vs Peers"]   = RED_P
                        return styles

                    st.dataframe(
                        display_only.style.apply(style_peer_table, axis=None),
                        use_container_width=True,
                        hide_index=True,
                    )

                # ── QO argument cards ─────────────────────────────────────────
                st.markdown("#### 💬 Data-Driven QO Arguments")
                st.caption(
                    "Arguments ranked by strength — those where the target country "
                    "materially outperforms peers are the most compelling in a rating committee dialogue."
                )

                # Sort by strength of outperformance
                strengths_for_qo = sorted(
                    [r2 for r2 in comp_rows if r2["_better"]],
                    key=lambda x: x["_diff"], reverse=True
                )
                weaknesses_for_qo = [r2 for r2 in comp_rows if r2["_worse"]]

                # Map indicator to QO factor
                QO_FACTOR_MAP = {
                    'Real GDP Growth (%)':     'Fiscal Flexibility & Economic Resilience',
                    'Fiscal Balance (% GDP)':  'Fiscal Flexibility',
                    'Debt-to-GDP (%)':         'Fiscal Flexibility',
                    'Interest/Revenue (%)':    'Fiscal Flexibility',
                    'GEFN (% CAR)':            'External Liquidity & IIP',
                    'Reserves (months)':       'External Liquidity & IIP',
                    'CPI Inflation (%)':       'Monetary Flexibility',
                    'Financial Depth (% GDP)': 'Monetary Flexibility',
                }

                QO_NARRATIVE_TEMPLATE = {
                    'Real GDP Growth (%)': (
                        "Growth of {t_val:.1f}% vs BBB peer average of {p_avg:.1f}% — "
                        "outperformance of {diff:+.1f}pp. S&P para. 15 explicitly recognises "
                        "sustained over-performance vs similarly rated peers as a basis for "
                        "positive residual adjustment. Highlight multi-year trend growth "
                        "trajectory, not just the current-year estimate."
                    ),
                    'Fiscal Balance (% GDP)': (
                        "Fiscal balance of {t_val:.1f}% vs peer average of {p_avg:.1f}% — "
                        "{diff:+.1f}pp advantage. A tighter deficit demonstrates greater "
                        "fiscal discipline than the BBB cohort and supports the argument for "
                        "fiscal flexibility under the QO Fiscal Flexibility factor."
                    ),
                    'Debt-to-GDP (%)': (
                        "Net debt of {t_val:.1f}% of GDP vs peer average of {p_avg:.1f}% — "
                        "{diff:+.1f}pp lower. Below-peer debt burden provides greater shock "
                        "absorption capacity and supports a positive Fiscal Flexibility argument. "
                        "Emphasise declining debt trajectory in medium-term fiscal framework."
                    ),
                    'Interest/Revenue (%)': (
                        "Interest-to-revenue of {t_val:.1f}% vs peer average of {p_avg:.1f}% — "
                        "{diff:+.1f}pp advantage. Lower debt servicing cost relative to revenues "
                        "signals stronger fiscal headroom than peers — a compelling Fiscal "
                        "Flexibility argument in the QO dialogue."
                    ),
                    'GEFN (% CAR)': (
                        "GEFN of {t_val:.1f}% vs peer average of {p_avg:.1f}% — "
                        "{diff:+.1f}pp lower external rollover need. Demonstrates superior "
                        "external liquidity management relative to the BBB cohort. Use as "
                        "an External Liquidity & IIP positive argument."
                    ),
                    'Reserves (months)': (
                        "Reserve coverage of {t_val:.1f} months vs peer average of {p_avg:.1f} months. "
                        "Above-peer reserve buffer reduces vulnerability to sudden stop scenarios "
                        "and supports the External Liquidity & IIP QO factor."
                    ),
                    'CPI Inflation (%)': (
                        "Inflation of {t_val:.1f}% vs peer average of {p_avg:.1f}% — "
                        "{diff:+.1f}pp lower. Price stability closer to trading partner levels "
                        "supports monetary credibility arguments under the Monetary Flexibility "
                        "QO factor. Emphasise central bank independence and track record."
                    ),
                    'Financial Depth (% GDP)': (
                        "Financial depth of {t_val:.1f}% of GDP vs peer average of {p_avg:.1f}% — "
                        "{diff:+.1f}pp advantage. Deeper financial system enhances monetary "
                        "transmission effectiveness, supporting the Monetary Flexibility "
                        "QO factor argument."
                    ),
                }

                if strengths_for_qo:
                    for i, row2 in enumerate(strengths_for_qo, 1):
                        label    = row2["Indicator"]
                        t_val    = float(row2[f"{target}"])
                        p_avg    = float(row2["BBB Peer Avg"])
                        diff     = t_val - p_avg
                        factor   = QO_FACTOR_MAP.get(label, "General Assessment")
                        template = QO_NARRATIVE_TEMPLATE.get(label, "")
                        argument = template.format(
                            t_val=t_val, p_avg=p_avg, diff=diff) if template else ""

                        # Strength of argument
                        pct_diff = abs(diff) / abs(p_avg) * 100 if p_avg != 0 else 0
                        if pct_diff >= 30:
                            strength_label = "🔥 Very Strong Argument"
                            strength_color = "#065f46"
                            strength_bg    = "#d1fae5"
                        elif pct_diff >= 15:
                            strength_label = "💪 Strong Argument"
                            strength_color = "#1e3a8a"
                            strength_bg    = "#dbeafe"
                        else:
                            strength_label = "📌 Supporting Argument"
                            strength_color = "#6366f1"
                            strength_bg    = "#ede9fe"

                        st.markdown(f"""
                        <div style="padding:16px; border-radius:10px; background:{strength_bg};
                                    border-left:5px solid {strength_color}; margin-bottom:14px;">
                            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                                <div style="font-size:16px; font-weight:700; color:{strength_color};">
                                    #{i} ⚖️ {factor}
                                </div>
                                <div style="display:flex; gap:8px; align-items:center;">
                                    <span style="font-size:12px; font-weight:700; color:{strength_color};
                                                 background:white; padding:3px 10px; border-radius:20px;">
                                        {label}
                                    </span>
                                    <span style="font-size:12px; font-weight:700; color:white;
                                                 background:{strength_color}; padding:3px 10px; border-radius:20px;">
                                        {strength_label}
                                    </span>
                                </div>
                            </div>
                            <div style="display:flex; gap:24px; margin-bottom:8px;">
                                <span style="font-size:13px; color:#374151;">
                                    <b>{target}:</b> {t_val:.1f}
                                </span>
                                <span style="font-size:13px; color:#374151;">
                                    <b>BBB Peer Avg:</b> {p_avg:.1f}
                                </span>
                                <span style="font-size:13px; font-weight:700; color:{strength_color};">
                                    Outperformance: {diff:+.1f} ({pct_diff:.0f}% vs peers)
                                </span>
                            </div>
                            <div style="font-size:15px; color:#1e293b; line-height:1.7;">
                                {argument}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                    # ── Overall QO recommendation ─────────────────────────────
                    top_factors = list(dict.fromkeys(
                        [QO_FACTOR_MAP.get(r2["Indicator"], "") for r2 in strengths_for_qo[:3]]
                    ))
                    st.markdown(f"""
                    <div style="padding:16px; border-radius:10px; background:#1E3A8A;
                                margin-top:8px; margin-bottom:8px;">
                        <div style="font-size:15px; font-weight:700; color:white; margin-bottom:8px;">
                            🎯 Overall QO Strategy for {target}
                        </div>
                        <div style="font-size:14px; color:#CADCFC; line-height:1.8;">
                            Based on peer comparison, the strongest QO arguments centre on
                            <b style="color:white;">{', '.join(top_factors)}</b>.
                            In the rating committee dialogue, lead with the
                            <b style="color:white;">{strengths_for_qo[0]['Indicator']}</b> advantage
                            ({float(strengths_for_qo[0][target]):.1f} vs peer avg
                            {float(strengths_for_qo[0]['BBB Peer Avg']):.1f}),
                            which represents the largest relative outperformance vs the BBB cohort
                            at {float(strengths_for_qo[0]['_diff']) / abs(float(strengths_for_qo[0]['BBB Peer Avg'])) * 100:.0f}%.
                            {'Pair this with the weakness mitigation narrative below to demonstrate a balanced credit profile.' if weaknesses_for_qo else 'The absence of material weaknesses vs peers further strengthens the case for a positive residual adjustment.'}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                else:
                    st.warning(
                        f"No indicators where {target} materially outperforms BBB peers. "
                        "Focus on addressing weaknesses before building a QO upgrade narrative."
                    )

                # ── Areas to mitigate ─────────────────────────────────────────
                if weaknesses_for_qo:
                    with st.expander(f"⚠️ Areas where {target} underperforms peers ({len(weaknesses_for_qo)} indicators) — address to protect QO"):
                        for row2 in weaknesses_for_qo:
                            label  = row2["Indicator"]
                            t_val  = float(row2[f"{target}"])
                            p_avg  = float(row2["BBB Peer Avg"])
                            diff   = t_val - p_avg
                            st.markdown(f"""
                            <div style="padding:10px; border-radius:8px; background:#fff5f5;
                                        border-left:4px solid #ef4444; margin-bottom:8px;">
                                <b style="color:#991b1b;">{label}</b>
                                &nbsp;|&nbsp;
                                {target}: <b>{t_val:.1f}</b>
                                &nbsp;vs BBB Peer Avg: <b>{p_avg:.1f}</b>
                                &nbsp;(<span style="color:#991b1b; font-weight:600;">{diff:+.1f}</span>)
                                <br>
                                <span style="font-size:13px; color:#7f1d1d;">
                                    Underperformance vs peers may be flagged by S&P as a negative
                                    residual factor — address proactively in rating dialogue.
                                </span>
                            </div>
                            """, unsafe_allow_html=True)
            else:
                # No peer trend data — fall back to original static opportunities
                if qo_opportunities:
                    for opp in qo_opportunities:
                        impact_color = "#10b981" if "+" in opp["impact"] else "#6366f1"
                        st.markdown(f"""
                        <div style="padding:16px; border-radius:10px; background:#f0f4ff;
                                    border-left:5px solid {impact_color}; margin-bottom:14px;">
                            <div style="font-size:16px; font-weight:700; color:#1e3a8a;">
                                ⚖️ {opp['factor']}
                                <span style="font-size:13px; font-weight:700; color:{impact_color};
                                             background:{impact_color}22; padding:3px 10px;
                                             border-radius:20px; margin-left:8px;">{opp['impact']}</span>
                            </div>
                            <div style="font-size:15px; color:#374151; margin-top:8px; line-height:1.7;">
                                {opp['argument']}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.warning(
                        "No QO narrative opportunities identified. "
                        "Visit the Peer Comparison tab to load peer data for enhanced analysis."
                    )

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
                "Recommendations, and QO targeting. Peer Comparison section is included "
                "only if you have selected peers in the **📉 Peer Comparison** tab."
            )

            # ── Resolve peer data (graceful fallback to target-only) ──────────
            try:
                briefing_comp_list = comp_list if (comp_list and len(comp_list) > 0) else None
            except NameError:
                briefing_comp_list = None

            # Always build at least a single-country entry for the target
            target_entry = {
                "Country": target,
                "🏛 Institutional": s_inst,
                "📈 Economic": s_eco,
                "💰 Fiscal": s_fis,
                "🌐 External": s_ext,
                "🏦 Monetary": s_mon,
                "_ie": prof_ie, "_fp": prof_fp,
                "_srm": srm_rating, "_qo": qo, "_nr": r,
            }

            if briefing_comp_list is None:
                briefing_comp_list = [target_entry]

            # Ensure target is always first entry
            countries_in_list = [c["Country"] for c in briefing_comp_list]
            if target not in countries_in_list:
                briefing_comp_list = [target_entry] + briefing_comp_list

            briefing_nations = [c["Country"] for c in briefing_comp_list]
            has_peers = len(briefing_comp_list) > 1

            # ── Resolve trend data ────────────────────────────────────────────
            try:
                briefing_trend_df = trend_df if (trend_df is not None and not trend_df.empty) else None
            except NameError:
                briefing_trend_df = None

            try:
                briefing_metrics = selected_metrics if (selected_metrics and len(selected_metrics) > 0) else []
            except NameError:
                briefing_metrics = []

            # ── Info summary ──────────────────────────────────────────────────
            col_i1, col_i2, col_i3 = st.columns(3)
            with col_i1:
                st.markdown("**Target Sovereign**")
                st.write(f"📌 {target} ({r['Actual_Rating']})")
            with col_i2:
                st.markdown("**Peers Included**")
                if has_peers:
                    st.write(", ".join(briefing_nations[1:]))
                else:
                    st.caption("None — visit Peer Comparison tab to add peers")
            with col_i3:
                st.markdown("**Trend Indicators**")
                st.write(", ".join(briefing_metrics) if briefing_metrics else "None selected")

            st.info(
                "💡 **To enrich the report:** visit **📉 Peer Comparison** to select peer countries "
                "and indicators before generating. The report works without them too."
            )
            st.divider()


            # ── Session state: persist PDF bytes and file.io link per target ──
            if st.session_state.get("_briefing_target") != target:
                st.session_state["_briefing_target"] = target
                st.session_state["_pdf_bytes"] = None
                st.session_state["_wa_file_url"] = None

            # ── Download buttons ──────────────────────────────────────────────
            dl1, dl2 = st.columns(2)

            with dl1:
                st.markdown("### 📑 PDF Briefing Note")
                st.markdown("Multi-page report with tables, charts, and narrative.")
                if st.button("⚙️ Generate PDF", use_container_width=True):
                    with st.spinner("Generating PDF..."):
                        pdf_buf = generate_pdf(
                            target=target, r=r,
                            srm_rating=srm_rating, qo=qo,
                            s_inst=s_inst, s_eco=s_eco,
                            s_fis=s_fis, s_ext=s_ext, s_mon=s_mon,
                            prof_ie=prof_ie, prof_fp=prof_fp,
                            comp_list=briefing_comp_list,
                            sel_nations=briefing_nations,
                            trend_df= trend_df,
                            selected_metrics=briefing_metrics,
                            signals=signals,
                            qo_opportunities=qo_opportunities,
                            supp_adj=supp_adj,
                            supp_factors=supp_factors,
                            final_indicative=final_indicative,
                            cap_note=cap_note,
                            df_master=df,
                        )
                    st.session_state["_pdf_bytes"] = pdf_buf.getvalue()
                    st.session_state["_wa_file_url"] = None  # reset old link

                if st.session_state["_pdf_bytes"] is not None:
                    st.download_button(
                        label="⬇️ Download PDF",
                        data=st.session_state["_pdf_bytes"],
                        file_name=f"Briefing_{target}_{now_jakarta().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                    if st.button("📤 Upload & get shareable link", use_container_width=True,
                                 help="Uploads the PDF to a temporary public link so you can share it via WhatsApp below."):
                        with st.spinner("Uploading PDF…"):
                            _fname = f"Briefing_{target}_{now_jakarta().strftime('%Y%m%d')}.pdf"
                            _pdf  = st.session_state["_pdf_bytes"]
                            _upload_url = None

                            # Primary: tmpfiles.org (reliable, no auth, expires 1 h)
                            try:
                                r1 = requests.post(
                                    "https://tmpfiles.org/api/v1/upload",
                                    files={"file": (_fname, _pdf, "application/pdf")},
                                    timeout=30,
                                )
                                if r1.ok and r1.text:
                                    j1 = r1.json()
                                    if j1.get("status") == "success":
                                        # convert view URL → direct download URL
                                        _upload_url = j1["data"]["url"].replace(
                                            "tmpfiles.org/", "tmpfiles.org/dl/"
                                        )
                            except Exception:
                                pass

                            # Fallback: file.io (expires 1 day)
                            if not _upload_url:
                                try:
                                    r2 = requests.post(
                                        "https://file.io",
                                        files={"file": (_fname, _pdf, "application/pdf")},
                                        data={"expires": "1d"},
                                        timeout=30,
                                    )
                                    if r2.ok and r2.text:
                                        j2 = r2.json()
                                        if j2.get("success"):
                                            _upload_url = j2["link"]
                                except Exception:
                                    pass

                            if _upload_url:
                                st.session_state["_wa_file_url"] = _upload_url
                                st.success("✅ PDF link auto-inserted into the WhatsApp message below — just hit Send!")
                            else:
                                st.error("Both upload services failed. Check your internet connection, or download the PDF and attach it manually in WhatsApp.")

                    if st.session_state["_wa_file_url"]:
                        st.info("📎 PDF link already included in the WhatsApp message below. Scroll down and tap **Open WhatsApp & Send**.")

            with dl2:
                st.markdown("### 📊 PowerPoint Deck")
                st.markdown("Editable slide deck ready for presentations.")
                if st.button("⚙️ Generate PowerPoint", use_container_width=True):
                    with st.spinner("Generating PowerPoint..."):
                        pptx_buf = generate_pptx(
                         target=target, r=r,
                            srm_rating=srm_rating, qo=qo,
                            s_inst=s_inst, s_eco=s_eco,
                            s_fis=s_fis, s_ext=s_ext, s_mon=s_mon,
                            prof_ie=prof_ie, prof_fp=prof_fp,
                            comp_list=briefing_comp_list,
                            sel_nations=briefing_nations,
                            trend_df= trend_df,
                            selected_metrics=briefing_metrics,
                            signals=signals,
                            qo_opportunities=qo_opportunities,
                            supp_adj=supp_adj,
                            supp_factors=supp_factors,
                            final_indicative=final_indicative,
                            cap_note=cap_note,
                            df_master=df,
                    )
                    st.download_button(
                        label="⬇️ Download PowerPoint",
                        data=pptx_buf,
                        file_name=f"Briefing_{target}_{now_jakarta().strftime('%Y%m%d')}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )

            # ── WhatsApp Share ────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### 📲 Share via WhatsApp")

            wa_col1, wa_col2 = st.columns([2, 1])
            with wa_col1:
                wa_phone = st.text_input(
                    "Recipient phone number (optional)",
                    placeholder="+628123456789  — leave blank to open WhatsApp chat picker",
                    help="Include country code, e.g. +6281234567890. Leave blank to choose the contact inside WhatsApp.",
                )
            with wa_col2:
                wa_lang = st.selectbox("Message language", ["English", "Indonesian"], index=0)

            _file_url = st.session_state.get("_wa_file_url")

            def _build_wa_message(lang, file_url=None):
                date_str = now_jakarta().strftime("%d %b %Y")
                qo_sign  = f"+{int(qo)}" if qo > 0 else str(int(qo))
                adj_sign = f"+{int(supp_adj)}" if supp_adj > 0 else str(int(supp_adj))
                link_line = f"\n📎 *PDF Report:* {file_url}" if file_url else ""
                if lang == "Indonesian":
                    msg = (
                        f"*📊 Briefing Note Sovereign Rating — {target}*\n"
                        f"_{date_str}_\n\n"
                        f"*Rating Aktual S&P:* {r['Actual_Rating']}\n"
                        f"*Output SRM:* {srm_rating}\n"
                        f"*Setelah Supplemental ({adj_sign}):* {final_indicative}\n"
                        f"*Penyesuaian QO:* {qo_sign} notch\n\n"
                        f"*Skor Pilar:*\n"
                        f"  • Institusional: {s_inst:.1f}\n"
                        f"  • Ekonomi: {s_eco:.1f}\n"
                        f"  • Fiskal: {s_fis:.1f}\n"
                        f"  • Eksternal: {s_ext:.1f}\n"
                        f"  • Moneter: {s_mon:.1f}\n\n"
                        f"  IE Profile: {prof_ie:.2f} | FP Profile: {prof_fp:.2f}"
                        f"{link_line}"
                    )
                else:
                    msg = (
                        f"*📊 Sovereign Rating Briefing Note — {target}*\n"
                        f"_{date_str}_\n\n"
                        f"*S&P Actual Rating:* {r['Actual_Rating']}\n"
                        f"*SRM Matrix Output:* {srm_rating}\n"
                        f"*Post-Supplemental ({adj_sign}):* {final_indicative}\n"
                        f"*QO Adjustment:* {qo_sign} notch\n\n"
                        f"*Pillar Scores:*\n"
                        f"  • Institutional: {s_inst:.1f}\n"
                        f"  • Economic: {s_eco:.1f}\n"
                        f"  • Fiscal: {s_fis:.1f}\n"
                        f"  • External: {s_ext:.1f}\n"
                        f"  • Monetary: {s_mon:.1f}\n\n"
                        f"  IE Profile: {prof_ie:.2f} | FP Profile: {prof_fp:.2f}"
                        f"{link_line}"
                    )
                return msg

            wa_message = _build_wa_message(wa_lang, _file_url)

            if not _file_url:
                st.info("📎 To attach the PDF: **① Generate PDF** → **② Upload & get shareable link** — the link will be inserted into the message automatically.")

            st.text_area("Message preview", value=wa_message, height=250, disabled=True)

            if _file_url:
                st.caption("✅ PDF link is already in the message. Tap the button below to open WhatsApp and send.")

            encoded_msg = urllib.parse.quote(wa_message)
            phone_clean = wa_phone.strip().replace(" ", "").replace("-", "")
            if phone_clean:
                wa_url = f"https://wa.me/{phone_clean.lstrip('+')}?text={encoded_msg}"
            else:
                wa_url = f"https://wa.me/?text={encoded_msg}"

            st.link_button(
                "📲 Open WhatsApp & Send",
                url=wa_url,
                use_container_width=True,
            )
            st.caption("Opens WhatsApp Web or the app with the message pre-filled. The recipient taps the PDF link inside the message to download it.")

        # ── Compute methodology-aligned defaults from trend_df ────────────────
        def get_avg(country, metric, trend_df, yr_start_type='e', n_years=3):
            """Average from estimate year + n forecast years."""
            try:
                sub = trend_df[
                    (trend_df['Country'] == country) &
                    (trend_df['Metric'] == metric)
                ].copy()
                if sub.empty: return np.nan
                sub['_n'] = sub['Year'].apply(
                    lambda y: float(str(y).replace('e','').replace('f',''))
                    if str(y).strip().endswith(('e','f')) or str(y).strip().isdigit()
                    else 0)
                sub['_type'] = sub['Year'].apply(
                    lambda y: 'e' if str(y).strip().endswith('e')
                    else ('f' if str(y).strip().endswith('f') else 'h'))
                est = sub[sub['_type']=='e']
                if est.empty: return np.nan
                anchor = est['_n'].max()
                window = sub[sub['_n'].between(anchor, anchor + n_years - 1)]
                return round(window['Value'].mean(), 2) if not window.empty else np.nan
            except:
                return np.nan

        def get_10yr_trend_growth(country, metric, trend_df):
            """
            10-year weighted average per S&P para. 36:
            6 historical + 1 estimate + 3 forecast.
            Latest year, estimate, and forecasts weighted 100%;
            earlier years weighted lower.
            """
            try:
                sub = trend_df[
                    (trend_df['Country'] == country) &
                    (trend_df['Metric'] == metric)
                ].copy()
                if sub.empty: return np.nan
                sub['_n'] = sub['Year'].apply(
                    lambda y: float(str(y).replace('e','').replace('f','')))
                sub['_type'] = sub['Year'].apply(
                    lambda y: 'e' if str(y).strip().endswith('e')
                    else ('f' if str(y).strip().endswith('f') else 'h'))
                sub = sub.sort_values('_n').tail(10)

                # S&P: latest hist, estimate, forecasts = weight 1.0; earlier = weight 0.5
                est_yr = sub[sub['_type']=='e']['_n'].max() if not sub[sub['_type']=='e'].empty else sub['_n'].max()
                sub['_w'] = sub['_n'].apply(lambda n: 1.0 if n >= est_yr - 1 else 0.5)
                wavg = (sub['Value'] * sub['_w']).sum() / sub['_w'].sum()
                return round(wavg, 2)
            except:
                return np.nan

        def get_cycle_avg(country, metric, trend_df, n=5):
            """Multi-year cycle average for inflation — use 5-year average."""
            try:
                sub = trend_df[
                    (trend_df['Country'] == country) &
                    (trend_df['Metric'] == metric)
                ].copy()
                if sub.empty: return np.nan
                sub['_n'] = sub['Year'].apply(
                    lambda y: float(str(y).replace('e','').replace('f','')))
                sub['_type'] = sub['Year'].apply(
                    lambda y: 'e' if str(y).strip().endswith('e')
                    else ('f' if str(y).strip().endswith('f') else 'h'))
                # Use historical + estimate only (not forecast) for inflation cycle
                hist_est = sub[sub['_type'].isin(['h','e'])].sort_values('_n').tail(n)
                return round(hist_est['Value'].mean(), 2) if not hist_est.empty else np.nan
            except:
                return np.nan

        # ── Apply per S&P methodology ─────────────────────────────────────────
        _td = trend_df if trend_df is not None and not trend_df.empty else None

        # GDP per capita — current year estimate only (already in r['GDP_PC'])
        sim_gdp_default = float(r['GDP_PC'])

        # Real GDP growth — 10-year weighted trend (para. 36)
        _g10 = get_10yr_trend_growth(target, 'Real GDP growth (%)', _td) if _td is not None else np.nan
        sim_growth_default = _g10 if not np.isnan(_g10) else float(r['Growth'])

        # Fiscal balance — average current estimate + 2 forecast years (para. 73)
        _bal = get_avg(target, 'GG balance/GDP (%)', _td, n_years=3) if _td is not None else np.nan
        sim_bal_default = _bal if not np.isnan(_bal) else float(r['Balance'])

        # Net debt/GDP — current year estimate (burden assessment uses current level)
        sim_debt_default = float(r['Debt_GDP'])

        # Interest/Revenue — current year estimate
        sim_int_default = float(r['Int_Rev'])

        # GEFN — average current estimate + 2 forecast years (para. 54)
        _gefn = get_avg(target, 'Gross ext. fin. needs/(CAR + use. res.) (%)', _td, n_years=3) if _td is not None else np.nan
        sim_gefn_default = _gefn if not np.isnan(_gefn) else float(r['GEFN'])

        # NIIP — current year estimate
        sim_niip_default = float(r['NIIP_CAR'])

        # Inflation — 5-year cycle average (para. 115-117)
        _cpi = get_cycle_avg(target, 'CPI growth (%)', _td, n=5) if _td is not None else np.nan
        sim_cpi_default = _cpi if not np.isnan(_cpi) else float(r['CPI'])

        # Financial depth — current year estimate
        sim_depth_default = int(r['Fin_Depth'])

        # Reserves — current year estimate
        sim_reserves_default = float(r['Reserves'])

        with tabs[4]: # METHODOLOGY SIMULATOR (GAUGE STYLE)
            st.subheader("Interactive Stress-Test Simulator")
            st.info("Gunakan input angka di bawah untuk mensimulasikan dampak perubahan indikator terhadap rating indikatif.")
            st.caption(
                "📐 **Methodology-aligned defaults:** GDP growth = 10-year weighted trend | "
                "Fiscal balance & GEFN = 3-year average (current estimate + 2 forecast) | "
                "Inflation = 5-year cycle average | Other metrics = current year estimate. "
                "All values can be manually adjusted for stress-testing."
            )
            
            sc1, sc2, sc3 = st.columns(3)
            
            # Input Angka Presisi (Tanpa Slider untuk akurasi)
            sc1.markdown("### Economic & Institutional")
            sim_gdp = sc1.number_input("GDP Per Capita (USD)", value=float(r['GDP_PC']), step=100.0)
            sim_growth = sc1.number_input("Trend Real GDP Growth (%)",
                value=sim_growth_default, step=0.1,
                help="S&P uses 10-year weighted average (6 hist + 1 est + 3 fcst)")
            sim_div = sc1.selectbox("Economic Diversification", ["High", "Standard", "Low"], index = 1)
            sim_wgi = sc1.number_input("WGI Score (Governance)", value=float(r['WGI_Score']) if pd.notna(r['WGI_Score']) else 50.0, step=1.0)

            sc1.markdown("### Supplemental Factors")
            sim_event_risk     = sc1.toggle("Event Risk (imminent political/security risk)", value=False)
            sim_liquid_assets  = sc1.number_input("Liquid Govt Assets (% GDP)", value=0.0, step=1.0,
                                                   help="Values >100% with net asset position → +1 notch uplift")
            sc1.markdown("### Transitional Factors")
            sim_resource_discovery = sc1.toggle(
                "Significant Resource Discovery",
                value=False,
                help="Major gas/oil/mineral discovery expected to materially improve fiscal "
                     "and external outlook within 3–5 years. Supports +1 residual notch argument."
            )
            sim_resource_magnitude = sc1.selectbox(
                "Discovery Magnitude",
                ["Large (>2% GDP revenue impact)", "Moderate (1–2% GDP)", "Small (<1% GDP)"],
                index=0,
                disabled=not sim_resource_discovery,
            ) if sim_resource_discovery else None

            sc2.markdown("### Fiscal & Debt")
            sim_bal  = sc2.number_input("Fiscal Balance (% GDP)",
                value=sim_bal_default, step=0.1,
                help="S&P uses average of current estimate + 2–3 forecast years")


            sim_debt = sc2.number_input("Debt-to-GDP (%)",
                value=sim_debt_default, step=0.5,
                help="S&P uses current year estimate")
            sim_int  = sc2.number_input("Interest-to-Revenue (%)",
                value=sim_int_default, step=0.1,
                help="S&P uses current year estimate")
            sim_flex = sc2.selectbox("Revenue Flexibility", ["High", "Neutral", "Low"], index=1)

            sc3.markdown("### External & Monetary")
            sim_gefn = sc3.number_input("GEFN (% of CAR)",
                value=sim_gefn_default, step=1.0,
                help="S&P uses average of current estimate + 2–3 forecast years")
            sim_niip = sc3.number_input("NIIP (% of GDP)",
                value=sim_niip_default, step=1.0,
                help="S&P uses current year estimate + trend direction")
            sim_cpi  = sc3.number_input("Inflation Trends (%)",
                value=sim_cpi_default, step=0.1,
                help="S&P uses multi-year cycle average (~5 years hist + estimate)")
            sim_depth = sc3.number_input("Financial Depth (0-100)",
                value=sim_depth_default, step=1,
                help="S&P uses current year estimate")
            sim_regime = sc3.selectbox("FX Regime", ["Floating", "Fixed/Managed"])

            # ── Recalculate pillar scores ─────────────────────────────────────
            p_inst_sim = max(1.0, min(6.0, 6 - (sim_wgi / 20)))
            p_eco_sim  = score_economic(sim_gdp, sim_growth, sim_div)

            # ── Resource discovery adjustment (para. 15 transitional factor) ──
            resource_eco_adj  = 0
            resource_ext_adj  = 0
            resource_fis_adj  = 0
            resource_narrative = ""

            if sim_resource_discovery:
                if sim_resource_magnitude and "Large" in sim_resource_magnitude:
                    resource_eco_adj = -0.5   # better economic score
                    resource_ext_adj = -0.5   # better external (future exports)
                    resource_fis_adj = -0.5   # better fiscal (future revenues)
                    resource_narrative = (
                        "Large resource discovery — significant improvement to medium-term "
                        "fiscal, external, and economic outlook. Supports para. 15 transitional "
                        "positive argument for +1 residual notch in rating committee dialogue."
                    )
                elif sim_resource_magnitude and "Moderate" in sim_resource_magnitude:
                    resource_eco_adj = -0.3
                    resource_ext_adj = -0.3
                    resource_fis_adj = 0
                    resource_narrative = (
                        "Moderate resource discovery — gradual improvement expected. "
                        "Supports transitional positive narrative but insufficient alone "
                        "for a full +1 notch residual adjustment without corroborating factors."
                    )
                else:
                    resource_eco_adj = -0.2
                    resource_narrative = (
                        "Small resource discovery — limited macroeconomic impact. "
                        "May support a transitional positive argument if combined with "
                        "other improving factors."
                    )

            # Apply adjustments (capped at 1.0 floor)
            p_eco_sim = max(1.0, p_eco_sim + resource_eco_adj)

            p_fis_sim, p_fis_perf_sim, p_fis_burd_sim = score_fiscal(
                sim_debt, sim_int, sim_bal, sim_flex)
            p_ext_sim, _, _ = score_external(sim_gefn, sim_niip, 20, r['Reserves'])

            # Apply resource discovery adjustments to fiscal and external
            if sim_resource_discovery:
                p_fis_sim = max(1.0, p_fis_sim + resource_fis_adj)
                p_ext_sim = max(1.0, p_ext_sim + resource_ext_adj)
            p_mon_sim  = score_monetary(sim_regime, "High", sim_cpi, sim_depth)

            # ── IE / FP profiles & SRM matrix ────────────────────────────────
            res_ie     = (p_inst_sim + p_eco_sim) / 2
            res_fp     = (p_fis_sim + p_ext_sim + p_mon_sim) / 3
            sim_rating = get_indicative_rating(res_ie, res_fp)

            # ── Supplemental Adjustment Factors ──────────────────────────────
            sim_supp_adj, sim_supp_factors = score_supplemental(
                s_inst          = p_inst_sim,
                s_ext           = p_ext_sim,
                s_fis_burd      = p_fis_burd_sim,
                debt_gdp        = sim_debt,
                gefn            = sim_gefn,
                liquid_assets_gdp = sim_liquid_assets,
                event_risk      = sim_event_risk,
            )
            sim_final, sim_cap = apply_supplemental_caps(
                sim_rating, p_inst_sim, p_fis_burd_sim, sim_supp_adj)

            # ── Progress bar (based on final post-supplemental rating) ────────
            rating_val   = RATING_TO_NUM.get(sim_final, 0)
            progress_val = rating_val / 16.0

            st.write("---")

            # ── Rating derivation flow ────────────────────────────────────────
            st.markdown(f"""
            <div style="text-align:center; margin-bottom:12px;">
                <span style="font-size:14px; color:#64748B;">SRM Matrix</span><br>
                <span style="font-size:28px; font-weight:800; color:#6366f1;">{sim_rating}</span>
            </div>
            """, unsafe_allow_html=True)

            if sim_supp_adj != 0 or sim_cap:
                st.markdown(f"""
                <div style="text-align:center; margin-bottom:12px;">
                    <span style="font-size:13px; color:#64748B;">
                        After Supplemental Adj. ({int(sim_supp_adj):+} notch)
                        {'&nbsp;|&nbsp;🚧 ' + sim_cap if sim_cap else ''}
                    </span><br>
                    <span style="font-size:36px; font-weight:800; color:#10B981;">{sim_final}</span>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="text-align:center; margin-bottom:12px;">
                    <span style="font-size:13px; color:#64748B;">
                        No supplemental adjustments triggered
                    </span><br>
                    <span style="font-size:36px; font-weight:800; color:#10B981;">{sim_final}</span>
                </div>
                """, unsafe_allow_html=True)

            st.progress(progress_val)

            # ── Profiles ──────────────────────────────────────────────────────
            ic1, ic2 = st.columns(2)
            ic1.info(f"**Institutional & Economic Profile:** {res_ie:.2f}")
            ic2.info(f"**Flexibility & Performance Profile:** {res_fp:.2f}")

            # ── Resource discovery transitional factor ────────────────────────
            if sim_resource_discovery and resource_narrative:
                st.markdown(f"""
                <div style="padding:12px;border-radius:8px;background:#d1fae5;
                            border-left:4px solid #10b981;margin-bottom:8px;">
                    <b>⛽ Significant Resource Discovery — Para. 15 Transitional Factor</b><br>
                    <span style="font-size:12px;">
                        {resource_narrative}<br>
                        <b>Score adjustments applied:</b>
                        Economic {resource_eco_adj:+.1f} |
                        Fiscal {resource_fis_adj:+.1f} |
                        External {resource_ext_adj:+.1f}
                    </span>
                </div>
                """, unsafe_allow_html=True)

                if "Large" in (sim_resource_magnitude or ""):
                    st.info(
                        "💡 **QO Narrative Opportunity:** In the S&P rating committee dialogue, "
                        "argue that the discovery represents a transitional positive dynamic "
                        "(para. 15) — improved fiscal revenue stream, stronger export outlook, "
                        "and enhanced debt-bearing capacity — not yet captured in current-year "
                        "quantitative metrics but material to medium-term creditworthiness."
                    )

            # ── Show triggered supplemental factors ───────────────────────────
            if sim_supp_factors:
                st.markdown("##### ⚡ Supplemental Factors Triggered")
                for f in sim_supp_factors:
                    color  = "#fee2e2" if f["type"] == "negative" else "#d1fae5"
                    border = "#ef4444" if f["type"] == "negative" else "#10b981"
                    icon   = "📉"      if f["type"] == "negative" else "📈"
                    st.markdown(f"""
                    <div style="padding:10px;border-radius:8px;background:{color};
                                border-left:4px solid {border};margin-bottom:6px;">
                        <b>{icon} {f['factor']}</b>
                        &nbsp;<span style="font-size:12px;font-weight:bold;
                        color:{'#991b1b' if f['type']=='negative' else '#065f46'};">
                        [{f['impact']}]</span><br>
                        <span style="font-size:12px;">{f['detail']}</span>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.success("✅ No supplemental adjustment factors triggered.")

            # ── Stress-Test PDF Report ────────────────────────────────────────
            st.divider()
            st.markdown("### 📑 Export Stress-Test Report")
            st.markdown(
                "Generate a PDF briefing note using the **simulated inputs above** "
                "instead of the actual baseline data."
            )

            # Show a diff summary so user knows what changed vs baseline
            deltas = []
            if abs(p_inst_sim - s_inst) >= 0.1: deltas.append(f"Institutional {s_inst:.1f}→{p_inst_sim:.1f}")
            if abs(p_eco_sim  - s_eco)  >= 0.1: deltas.append(f"Economic {s_eco:.1f}→{p_eco_sim:.1f}")
            if abs(p_fis_sim  - s_fis)  >= 0.1: deltas.append(f"Fiscal {s_fis:.1f}→{p_fis_sim:.1f}")
            if abs(p_ext_sim  - s_ext)  >= 0.1: deltas.append(f"External {s_ext:.1f}→{p_ext_sim:.1f}")
            if abs(p_mon_sim  - s_mon)  >= 0.1: deltas.append(f"Monetary {s_mon:.1f}→{p_mon_sim:.1f}")
            if deltas:
                st.caption(f"📊 Score changes vs baseline: {' | '.join(deltas)}")
            else:
                st.caption("📊 No pillar score changes vs baseline — adjust inputs above to create a stress scenario.")

            if st.button("⚙️ Generate Stress-Test PDF", use_container_width=True):
                with st.spinner("Generating stress-test PDF..."):
                    _sim_qo = RATING_TO_NUM.get(r['Actual_Rating'], 8) - RATING_TO_NUM.get(sim_final, 8)
                    _stress_buf = generate_pdf(
                        target=target, r=r,
                        srm_rating=sim_rating, qo=_sim_qo,
                        s_inst=p_inst_sim, s_eco=p_eco_sim,
                        s_fis=p_fis_sim, s_ext=p_ext_sim, s_mon=p_mon_sim,
                        prof_ie=res_ie, prof_fp=res_fp,
                        comp_list=briefing_comp_list,
                        sel_nations=briefing_nations,
                        trend_df=trend_df,
                        selected_metrics=briefing_metrics,
                        signals=signals,
                        qo_opportunities=qo_opportunities,
                        supp_adj=sim_supp_adj,
                        supp_factors=sim_supp_factors,
                        final_indicative=sim_final,
                        cap_note=sim_cap or "",
                        df_master=df,
                    )
                st.session_state["_stress_pdf_bytes"] = _stress_buf.getvalue()
                st.session_state["_stress_pdf_scenario"] = (
                    f"{target} | SRM: {sim_rating}"
                    + (f" → {sim_final} (after supp.)" if sim_supp_adj != 0 else "")
                )
                st.success("✅ Stress-test PDF ready — click below to download.")

            if st.session_state.get("_stress_pdf_bytes"):
                st.download_button(
                    label="⬇️ Download Stress-Test PDF",
                    data=st.session_state["_stress_pdf_bytes"],
                    file_name=f"StressTest_{target}_{now_jakarta().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
                st.caption(
                    f"📄 Scenario captured: {st.session_state.get('_stress_pdf_scenario', '')}. "
                    "Regenerate if you change inputs above."
                )

else:
    st.info("👋 Welcome! Please upload the S&P Macro dataset to begin. Reference WGI data will be loaded automatically from the cloud.")