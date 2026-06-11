# 🏛️ Sovereign Rating Monitoring Dashboard

<div align="center">

![Methodology](https://img.shields.io/badge/Methodology-S%26P%20Global%20%28Oct%202024%29-003087?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-Live-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)
![Platform](https://img.shields.io/badge/Infrastructure-dashboardspp.cloud-1E3A8A?style=for-the-badge)

**A professional sovereign credit rating analytics platform aligned with S&P Global Methodology**

[🚀 **Launch Dashboard**](https://dashboardspp.cloud) · [📖 Methodology](#-methodology) · [✨ Features](#-features) · [🛠️ Setup & Secrets](#%EF%B8%8F-setup--secrets)

---

</div>

## 📌 Overview

This dashboard replicates and automates the **S&P Global Sovereign Rating Methodology** (including the core Sovereign Rating Model/SRM and Supplemental Adjustments updated per the **October 2024 Appendix D updates**). It provides dual-source data integration, qualitative evaluation anchoring, interactive stress-testing, and automated generation of policy briefing notes.

---

## ✨ Features

### 📊 National Portfolio
- **SRM Matrix Alignment**: Automatically calculates Institutional & Economic (IE) and Flexibility & Performance (FP) profiles to cross-reference S&P's 6×11 indicative rating matrix.
- **Supplemental Adjustment Engine**: Models extreme risk factors and hard rating caps (e.g., Institutional Score of 6 capping ratings at `BB+` or `B+` depending on debt burden; severe GEFN external liquidity pressure downgrades).
- **Residual ±1 Notch Capture**: Isolates and reverse-engineers the qualitative evaluation committee gap between model-derived outcomes and officially published ratings.

### 📉 Peer Comparison & Dynamic Trends
- **Pillar Heatmaps**: Side-by-side benchmarking of up to 7 countries utilizing a conditional gradient heatmap (Scale 1–6).
- **Dynamic Year & Suffix Detection**: Automatically parses historical and forecast indicators regardless of moving 10-year Excel structural windows (e.g., tracking `2025e` estimates, `2026f` forecasts, and historical actuals).
- **Visualized Projections**: Displays historical metrics in solid trends and projected forecasts inside custom shaded areas with dynamic vertical boundary indicators.

### 💡 Policy Recommendation & QO Strategy
- **Data-Driven Committee Narratives**: Ranks structural outperformance against peer group averages to formulate optimized defense strategies for credit rating dialogues.
- **Urgency-Ranked Policy Actions**: Triangulates macroeconomic weaknesses to auto-generate concrete fiscal and monetary adjustment suggestions.

### 📄 Briefing Notes & Mobile Distribution
- **Dual-Format Compilation**: Instant server-side rendering of comprehensive **PDF Briefings** (via ReportLab) and editable **PowerPoint Presentation Decks** (via python-pptx).
- **WhatsApp Cloud Sharing integration**: Automated public link generation via temporary file-hosting microservices (`tmpfiles.org` / `file.io`) with localized preset message formatting for immediate WhatsApp text forwarding.

### 📁 Counter Report Repository (GitHub-backed Storage)
- **Persistent Cloud Archival**: A production-grade Document Management System synced directly with a GitHub contents repository using a secure authenticated baseline JSON index (`index.json`).
- **Administrative CRUD Layer**: Full title editing, abstract summary management, file upload tracking, and record soft-deletions fully decoupled from local storage persistence.

---

## 📖 Methodology & Scoring Engine

The engine computes specific geometric and moving averages strictly tailored to rating criteria document definitions:
- **Economic Profile**: Evaluates current 2024 GDP per Capita brackets alongside a **10-year weighted trend growth average** ($6\text{ historical} + 1\text{ estimate} + 3\text{ forecast years}$).
- **Fiscal Profile**: Evaluates **3-year average General Government Balances** against a dynamic matrix intersection of current-year Gross/Net Debt and Interest-to-Revenue ratios.
- **External Profile**: Aggregates a **3-year Gross External Financing Needs (GEFN)** average combined with current Net International Investment Positions (NIIP).
- **Monetary Profile**: Processes a **5-year historical price stability cycle average** calibrated against local financial depth thresholds.

---

## 🗂️ Project Structure
