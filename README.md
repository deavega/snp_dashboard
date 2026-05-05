# 🏛️ Sovereign Rating Monitoring Dashboard

<div align="center">

![Sovereign Rating](https://img.shields.io/badge/Methodology-S%26P%20Global-003087?style=for-the-badge&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PHBhdGggZD0iTTEyIDJMMiA3bDEwIDUgMTAtNS0xMC01ek0yIDE3bDEwIDUgMTAtNS0xMC01LTEwIDV6TTIgMTJsMTAgNSAxMC01LTEwLTUtMTAgNXoiLz48L3N2Zz4=)
![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-Live-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)
![Status](https://img.shields.io/badge/Status-Active-22c55e?style=for-the-badge)

**A professional sovereign credit rating analytics platform built on S&P Global methodology**

[🚀 **Launch Dashboard**](https://snp-dashboard.streamlit.app/) · [📖 Methodology](#methodology) · [✨ Features](#features) · [🛠️ Setup](#setup)

---

</div>

## 📌 Overview

This dashboard replicates and extends S&P Global's **Sovereign Rating Model (SRM)** to provide transparent, data-driven sovereign credit analysis. It covers **142+ countries** across five rating pillars — giving analysts, researchers, and policymakers a powerful tool to understand, compare, and stress-test sovereign creditworthiness.

> 🔐 **Live App:** [https://snp-dashboard.streamlit.app/](https://snp-dashboard.streamlit.app/)  
> Access is passcode-protected. Contact the administrator for credentials.

---

## ✨ Features

### 📊 National Portfolio
Drill into any sovereign's complete rating profile:
- **SRM Matrix Lookup** — calculates the Institutional & Economic (IE) and Flexibility & Performance (FP) profiles to derive an indicative rating
- **Qualitative Overlay (QO)** — reverse-engineers the notch difference between the SRM output and the published rating
- **Five-pillar scorecard** — Institutional, Economic, Fiscal, External, Monetary
- **In-depth pillar tabs** — each pillar shows raw metrics, S&P threshold logic, and signal assessment
- **Full QO decomposition** — covers all five qualitative adjustment factors with country-specific narrative

### 📉 Peer Comparison
Side-by-side benchmarking across up to 7 countries:
- Filter by **rating group** (AAA, AA, A, BBB, BB, B)
- **Pillar score chart** — grouped bar chart with heatmap
- **Metrics snapshot** — color-coded table (🟢 Strong / 🟡 Moderate / 🔴 Weak) based on S&P thresholds
- **Historical & Forecast Trends** — interactive line charts for all indicators across the full 10-year window (historical + estimates + forecasts), with toggle and shaded projection area

### 💡 Recommendation
Automated advisory for any selected sovereign:
- **Strengths / Watch Areas / Weaknesses** — metric-level signal detection
- **Priority action items** — ranked by urgency with specific policy actions
- **QO narrative opportunities** — data-driven arguments for rating committee dialogue
- **Rating trajectory verdict** — Upgrade Candidate / Stable / Downgrade Risk

### 📄 Briefing Notes Generator
One-click professional report generation:
- **PDF** — multi-page briefing with tables, charts, and narrative
- **PowerPoint** — editable slide deck ready for presentations
- Covers: National Portfolio → Peer Comparison → Recommendations → QO In-Depth

### 🧪 Methodology Simulator
Interactive stress-tester:
- Adjust any indicator using precise number inputs
- Instantly see the impact on the indicative SRM rating
- Covers all five pillars with custom FX regime and diversification inputs

---

## 🏗️ Methodology

This dashboard implements the **S&P Global Sovereign Rating Methodology** using a two-profile matrix system:

```
IE Profile = (Institutional Score + Economic Score) / 2        → 25% + 25% weight
FP Profile = (Fiscal Score + External Score + Monetary Score) / 3  → 16.7% each
```

| Profile | Pillars | Weight |
|---------|---------|--------|
| Institutional & Economic (IE) | WGI Governance + GDP/capita + Growth | 50% |
| Flexibility & Performance (FP) | Fiscal + External + Monetary | 50% |

The indicative rating is derived from a **6×11 SRM matrix** (IE profile × FP profile), with the Qualitative Overlay (QO) applied on top across five factors: Contingent Liabilities, Monetary Flexibility, External Liquidity & IIP, Fiscal Flexibility, and Political & Governance Risks.

---

## 📁 Project Structure

```
snp_dashboard/
├── snp_dashboard.py          # Main Streamlit application
├── briefing_generator.py     # PDF & PPTX report generator
├── requirements.txt          # Python dependencies
├── data/
│   └── snp_data.xlsx         # Default S&P macro dataset (private)
└── README.md
```

---

## 🛠️ Setup

### Prerequisites
- Python 3.9+
- A copy of the S&P macro dataset (`.xlsx`)

### Local Installation

```bash
# 1. Clone the repository
git clone https://github.com/yourusername/snp_dashboard.git
cd snp_dashboard

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the dashboard
streamlit run snp_dashboard.py
```

### Requirements

```txt
streamlit
pandas
numpy
plotly
requests
gdown
openpyxl
matplotlib
reportlab
python-pptx
```

### Streamlit Cloud Deployment

1. Push the repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) and connect the repo
3. Add secrets in **App Settings → Secrets**:

```toml
APP_PASSCODE = "your_passcode_here"
```

4. Place your dataset at `data/snp_data.xlsx` in the repo

---

## 📊 Data Sources

| Source | Description | Access |
|--------|-------------|--------|
| **S&P Global Macro Dataset** | Primary indicator data for 142+ sovereigns across 6 sheets | Uploaded by user or default file |
| **World Governance Indicators (WGI)** | Governance scores for institutional pillar scoring | Auto-fetched from Google Drive |

---

## 🗂️ Dashboard Sheets Used

| Sheet | Key Indicators |
|-------|---------------|
| Economic Data | GDP per capita, Real GDP growth, LT FC Rating |
| General Government Data | Fiscal balance, Net debt/GDP, Interest/Revenue |
| Monetary Data | CPI inflation, Financial depth (credit/GDP) |
| Balance-Of-Payments Data | GEFN, Reserves (months), Current account |
| External Balance Sheet | Net external liabilities/CARs (NIIP proxy) |
| Central Gov Debt & Borrowing | Debt structure and borrowing profile |

---

## 🎨 Screenshots

| Tab | Description |
|-----|-------------|
| 📊 National Portfolio | Full pillar breakdown with QO analysis |
| 📉 Peer Comparison | Multi-country benchmarking with trend charts |
| 💡 Recommendation | Automated strengths, weaknesses & QO narrative |
| 📄 Briefing Notes | One-click PDF/PPTX generation |
| 🧪 Simulator | Interactive what-if stress testing |

---

## ⚠️ Disclaimer

This dashboard is an **independent analytical tool** built for research and educational purposes. It is **not affiliated with, endorsed by, or a substitute for** S&P Global Ratings' official credit assessments. All ratings shown are model-derived indicative outputs and do not constitute investment advice.

---

## 📄 License

This project is for internal and research use only. The underlying S&P methodology is proprietary to S&P Global Ratings.

---

<div align="center">

Built with ❤️ using [Streamlit](https://streamlit.io) · [Plotly](https://plotly.com) · [ReportLab](https://www.reportlab.com) · [python-pptx](https://python-pptx.readthedocs.io)

**[🚀 Launch Dashboard →](https://snp-dashboard.streamlit.app/)**

</div>
