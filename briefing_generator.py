"""
briefing_generator.py
Generates PDF and PPTX briefing notes from sovereign rating dashboard data.
Import this module and call generate_pdf() or generate_pptx().
"""

import io
import textwrap

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd

# ── ReportLab ────────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, Image as RLImage, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

# ── python-pptx ──────────────────────────────────────────────────────────────
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from datetime import datetime, timezone, timedelta

# Jakarta timezone = UTC+7
def now_jakarta():
    return datetime.now(timezone.utc) + timedelta(hours=7)

# ─────────────────────────────────────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────────────────────────────────────
NAVY    = "#1E3A8A"
TEAL    = "#0D9488"
GREEN   = "#065f46"
YELLOW  = "#78350f"
RED     = "#991b1b"
LGREY   = "#F8FAFC"
MGREY   = "#64748B"
WHITE   = "#FFFFFF"

# ReportLab colour objects
RL_NAVY   = colors.HexColor(NAVY)
RL_TEAL   = colors.HexColor(TEAL)
RL_GREEN  = colors.HexColor("#d1fae5")
RL_YELLOW = colors.HexColor("#fef3c7")
RL_RED    = colors.HexColor("#fee2e2")
RL_LGREY  = colors.HexColor(LGREY)
RL_MGREY  = colors.HexColor(MGREY)
RL_WHITE  = colors.white

RATING_TO_NUM = {
    'AAA':16,'AA+':15,'AA':14,'AA-':13,'A+':12,'A':11,'A-':10,
    'BBB+':9,'BBB':8,'BBB-':7,'BB+':6,'BB':5,'BB-':4,'B+':3,'B':2,'B-':1,'CCC':0
}

# ─────────────────────────────────────────────────────────────────────────────
# CHART HELPERS (matplotlib → bytes)
# ─────────────────────────────────────────────────────────────────────────────

def _pillar_bar_chart(comp_list, width=7, height=3.2):
    """Grouped bar chart of pillar scores for selected countries."""
    pillars = ["🏛 Institutional","📈 Economic","💰 Fiscal","🌐 External","🏦 Monetary"]
    short   = ["Institutional","Economic","Fiscal","External","Monetary"]
    countries = [c["Country"] for c in comp_list]
    n = len(countries)
    x = np.arange(len(pillars))
    bar_w = 0.7 / max(n, 1)
    palette = ["#0D9488","#1E3A8A","#F59E0B","#EF4444","#8B5CF6","#10B981","#F97316"]

    fig, ax = plt.subplots(figsize=(width, height))
    for i, c in enumerate(comp_list):
        vals = [c[p] for p in pillars]
        offset = (i - n/2 + 0.5) * bar_w
        bars = ax.bar(x + offset, vals, bar_w * 0.9,
                      label=c["Country"], color=palette[i % len(palette)], alpha=0.85)

    ax.set_xticks(x)
    ax.set_xticklabels(short, fontsize=9)
    ax.set_ylim(0, 6.5)
    ax.set_ylabel("Score (1=Best, 6=Worst)", fontsize=8)
    ax.axhline(3, color="#CBD5E1", lw=0.8, ls="--")
    ax.legend(fontsize=8, loc="upper right", framealpha=0.7)
    ax.spines[["top","right"]].set_visible(False)
    ax.set_facecolor("#F8FAFC")
    fig.patch.set_facecolor("white")
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _trend_line_chart(trend_df, metric, sel_nations, width=5.5, height=2.8):
    """Single line chart for one metric, all selected countries."""
    palette = ["#0D9488","#1E3A8A","#F59E0B","#EF4444","#8B5CF6","#10B981","#F97316"]

    def year_to_num(y):
        try: return int(float(str(y).replace('e','').replace('f','')))
        except: return 9999
    def is_proj(y):
        s = str(y).strip()
        return s.endswith('e') or s.endswith('f')

    fig, ax = plt.subplots(figsize=(width, height))
    mdata = trend_df[trend_df['Metric'] == metric].copy()
    mdata['Year'] = mdata['Year'].astype(str)
    mdata['XNum'] = mdata['Year'].apply(year_to_num)
    mdata['IsProj'] = mdata['Year'].apply(is_proj)
    mdata = mdata.sort_values('XNum')

    if mdata.empty:
        plt.close(fig)
        return None

    # Shade forecast area
    proj_x = mdata[mdata['IsProj']]['XNum'].unique()
    if len(proj_x):
        ax.axvspan(proj_x.min() - 0.4, proj_x.max() + 0.4,
                   color="#EEF2FF", alpha=0.6, zorder=0)

    for i, country in enumerate(sel_nations):
        cdata = mdata[mdata['Country'] == country].dropna(subset=['Value'])
        if cdata.empty: continue
        color = palette[i % len(palette)]
        hist = cdata[~cdata['IsProj']]
        proj = cdata[cdata['IsProj']]
        lw = 2.0 if i == 0 else 1.4
        ls_proj = '-' if i == 0 else '--'
        if not hist.empty:
            ax.plot(hist['XNum'], hist['Value'], color=color, lw=lw,
                    marker='o', markersize=3, label=country)
        if not proj.empty:
            px = ([hist['XNum'].iloc[-1]] if not hist.empty else []) + proj['XNum'].tolist()
            py = ([hist['Value'].iloc[-1]] if not hist.empty else []) + proj['Value'].tolist()
            ax.plot(px, py, color=color, lw=lw, ls=ls_proj,
                    marker='o', markersize=3, markerfacecolor='none')

    # x-ticks: use Year labels
    all_years = mdata.drop_duplicates('XNum').sort_values('XNum')
    ax.set_xticks(all_years['XNum'].tolist())
    ax.set_xticklabels(all_years['Year'].tolist(), rotation=30, ha='right', fontsize=7)
    ax.set_title(metric, fontsize=8, fontweight='bold', pad=4)
    ax.set_ylabel("", fontsize=7)
    ax.legend(fontsize=7, loc='best', framealpha=0.6)
    ax.spines[["top","right"]].set_visible(False)
    ax.set_facecolor("#F8FAFC")
    fig.patch.set_facecolor("white")
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _rating_color(rating):
    num = RATING_TO_NUM.get(str(rating).replace("*","").strip(), -1)
    if num >= 10: return RL_GREEN, colors.HexColor(GREEN)
    if num >= 6:  return RL_YELLOW, colors.HexColor(YELLOW)
    return RL_RED, colors.HexColor(RED)


def _metric_color(metric, val):
    """Return (bg, fg) ReportLab colours based on S&P thresholds."""
    THRESHOLDS = {
        'GDP per Capita (USD)':    (True,  [20000, 7000]),
        'Real GDP Growth (%)':     (True,  [4.0, 1.0]),
        'Fiscal Balance (% GDP)':  (True,  [-3.0, -5.0]),
        'Debt-to-GDP (%)':         (False, [45, 60]),
        'Interest/Revenue (%)':    (False, [15, 25]),
        'GEFN (% CAR)':           (False, [50, 75]),
        'NIIP (% GDP)':           (True,  [0, -50]),
        'Reserves (months)':       (True,  [6.0, 3.0]),
        'CPI Inflation (%)':       (False, [3.0, 6.0]),
        'Financial Depth (% GDP)': (True,  [40, 20]),
    }
    if metric not in THRESHOLDS:
        return RL_WHITE, colors.black
    try: val = float(val)
    except: return RL_WHITE, colors.black
    higher, (g, y) = THRESHOLDS[metric]
    if higher:
        if val >= g: return RL_GREEN, colors.HexColor(GREEN)
        if val >= y: return RL_YELLOW, colors.HexColor(YELLOW)
        return RL_RED, colors.HexColor(RED)
    else:
        if val <= g: return RL_GREEN, colors.HexColor(GREEN)
        if val <= y: return RL_YELLOW, colors.HexColor(YELLOW)
        return RL_RED, colors.HexColor(RED)


# ─────────────────────────────────────────────────────────────────────────────
# PDF GENERATOR
# ─────────────────────────────────────────────────────────────────────────────


def _compute_peer_qo_data(target, trend_df, df_master, RATING_TO_NUM):
    """
    Compute BBB-tier peer averages (latest estimate year) and target values
    for QO narrative metrics. Returns (peer_avgs, target_vals, peer_countries, comp_rows).
    """
    if trend_df is None or trend_df.empty or df_master is None:
        return {}, {}, [], []

    PEER_METRICS = {
        'Real GDP Growth (%)':     (True,  'Real GDP growth (%)'),
        'Fiscal Balance (% GDP)':  (True,  'GG balance/GDP (%)'),
        'Debt-to-GDP (%)':         (False, 'Net GG debt/GDP (%)'),
        'Interest/Revenue (%)':    (False, 'GG interest expenditure/revenues (%)'),
        'GEFN (% CAR)':            (False, 'Gross ext. fin. needs/(CAR + use. res.) (%)'),
        'Reserves (months)':       (True,  'Usable reserves/CAPs (months)'),
        'CPI Inflation (%)':       (False, 'CPI growth (%)'),
        'Financial Depth (% GDP)': (True,  "Banks' claims on resident non-gov't sector/GDP"),
    }

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

    QO_TEMPLATES = {
        'Real GDP Growth (%)': (
            "Growth of {t:.1f}% vs BBB peer average {p:.1f}% ({d:+.1f}pp). "
            "S&P para.15 recognises sustained outperformance vs similarly rated peers "
            "as a basis for positive residual adjustment. Highlight multi-year trend growth."
        ),
        'Fiscal Balance (% GDP)': (
            "Fiscal balance {t:.1f}% vs peer avg {p:.1f}% ({d:+.1f}pp advantage). "
            "Tighter deficit than the BBB cohort demonstrates superior fiscal discipline — "
            "a strong Fiscal Flexibility QO argument."
        ),
        'Debt-to-GDP (%)': (
            "Net debt {t:.1f}% of GDP vs peer avg {p:.1f}% ({d:+.1f}pp lower). "
            "Below-peer debt burden provides greater shock absorption, supporting "
            "Fiscal Flexibility argument. Emphasise declining debt trajectory."
        ),
        'Interest/Revenue (%)': (
            "Interest/revenue {t:.1f}% vs peer avg {p:.1f}% ({d:+.1f}pp advantage). "
            "Lower debt servicing cost signals stronger fiscal headroom — "
            "compelling Fiscal Flexibility argument in QO dialogue."
        ),
        'GEFN (% CAR)': (
            "GEFN {t:.1f}% vs peer avg {p:.1f}% ({d:+.1f}pp lower rollover need). "
            "Superior external liquidity management vs BBB cohort — "
            "use as External Liquidity & IIP positive argument."
        ),
        'Reserves (months)': (
            "Reserves {t:.1f} months vs peer avg {p:.1f} months. "
            "Above-peer buffer reduces sudden stop vulnerability — "
            "supports External Liquidity & IIP QO factor."
        ),
        'CPI Inflation (%)': (
            "Inflation {t:.1f}% vs peer avg {p:.1f}% ({d:+.1f}pp lower). "
            "Closer price stability supports monetary credibility — "
            "Monetary Flexibility QO factor. Emphasise central bank independence."
        ),
        'Financial Depth (% GDP)': (
            "Financial depth {t:.1f}% vs peer avg {p:.1f}% ({d:+.1f}pp advantage). "
            "Deeper financial system enhances monetary transmission — "
            "supports Monetary Flexibility QO argument."
        ),
    }

    # BBB tier peer countries
    peer_rating_nums = [
        RATING_TO_NUM.get('BBB-', 7),
        RATING_TO_NUM.get('BBB',  8),
        RATING_TO_NUM.get('BBB+', 9),
    ]
    rating_lookup = {
        str(row['Country']).strip(): str(row['Actual_Rating']).replace('*','').strip()
        for _, row in df_master.iterrows()
        if pd.notna(row['Country']) and str(row['Country']).strip() not in ('nan','None','')
    }
    peer_countries = [
        c for c in rating_lookup
        if RATING_TO_NUM.get(rating_lookup.get(c,''), -1) in peer_rating_nums
        and c != target
    ]

    def get_latest_est(country, raw_metric):
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
            sub = sub[sub['_s'] > 0]
            if sub.empty: return np.nan
            return float(sub.sort_values('_s').iloc[-1]['Value'])
        except:
            return np.nan

    peer_avgs   = {}
    target_vals = {}
    comp_rows   = []

    for label, (higher_better, raw_metric) in PEER_METRICS.items():
        peer_vals = [get_latest_est(pc, raw_metric) for pc in peer_countries]
        peer_vals = [v for v in peer_vals if not np.isnan(v)]
        p_avg = np.nanmean(peer_vals) if peer_vals else np.nan
        t_val = get_latest_est(target, raw_metric)

        peer_avgs[label]   = p_avg
        target_vals[label] = t_val

        if np.isnan(t_val) or np.isnan(p_avg): continue
        diff      = t_val - p_avg
        is_better = (diff > 0) if higher_better else (diff < 0)
        is_worse  = (diff < 0) if higher_better else (diff > 0)
        pct_diff  = abs(diff) / abs(p_avg) * 100 if p_avg != 0 else 0
        template  = QO_TEMPLATES.get(label, '')
        argument  = template.format(t=t_val, p=p_avg, d=diff) if template else ''
        factor    = QO_FACTOR_MAP.get(label, 'General Assessment')

        comp_rows.append({
            'label':        label,
            't_val':        t_val,
            'p_avg':        p_avg,
            'diff':         diff,
            'is_better':    is_better,
            'is_worse':     is_worse,
            'pct_diff':     pct_diff,
            'factor':       factor,
            'argument':     argument,
            'higher_better':higher_better,
        })

    return peer_avgs, target_vals, peer_countries, comp_rows

def generate_pdf(target, r, srm_rating, qo,
                 s_inst, s_eco, s_fis, s_ext, s_mon,
                 prof_ie, prof_fp,
                 comp_list, sel_nations, trend_df, selected_metrics,
                 signals, qo_opportunities,
                 supp_adj=0, supp_factors=None, final_indicative=None, cap_note=None,
                 df_master=None,
                 stress_test=False, baseline_scores=None):
    """
    Returns a bytes buffer containing the PDF briefing note.
    When stress_test=True, baseline_scores should be a dict with keys:
    s_inst, s_eco, s_fis, s_ext, s_mon, srm_rating, final_indicative.
    """
    supp_factors     = supp_factors or []
    final_indicative = final_indicative or srm_rating
    buf = io.BytesIO()
    _title = f"Stress-Test Scenario — {target}" if stress_test else f"Sovereign Rating Briefing — {target}"
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=2*cm, bottomMargin=2*cm,
        title=_title
    )

    # ── Styles ────────────────────────────────────────────────────────────────
    styles = getSampleStyleSheet()
    S = lambda name, **kw: ParagraphStyle(name, **kw)

    st_h1 = S("H1", fontSize=18, textColor=RL_NAVY, fontName="Helvetica-Bold",
               spaceAfter=6, spaceBefore=12)
    st_h2 = S("H2", fontSize=13, textColor=RL_TEAL, fontName="Helvetica-Bold",
               spaceAfter=4, spaceBefore=10)
    st_h3 = S("H3", fontSize=10, textColor=RL_NAVY, fontName="Helvetica-Bold",
               spaceAfter=3, spaceBefore=6)
    st_body = S("Body", fontSize=9, leading=13, spaceAfter=4, alignment=TA_JUSTIFY)
    st_small = S("Small", fontSize=8, textColor=RL_MGREY, leading=11)
    st_center = S("Center", fontSize=9, alignment=TA_CENTER, leading=12)
    st_caption = S("Caption", fontSize=7.5, textColor=RL_MGREY, alignment=TA_CENTER)

    story = []

    def hr(): return HRFlowable(width="100%", thickness=0.5,
                                color=colors.HexColor("#CBD5E1"), spaceAfter=4)
    def vsp(h=6): return Spacer(1, h)

    # ── COVER / HEADER ────────────────────────────────────────────────────────
    if stress_test:
        story.append(Paragraph("⚠️ Stress-Test Scenario Report", st_h1))
        story.append(Paragraph(f"<b>{target}</b> | Simulated Rating Analysis", st_h2))
        story.append(Paragraph(
            f"Generated: {now_jakarta().strftime('%d %B %Y, %H:%M')} &nbsp;|&nbsp; "
            f"This report reflects <b>hypothetical inputs</b>, not the baseline actual data.",
            st_small))
        story.append(hr())

        # Stress-test disclaimer banner
        disclaimer_data = [[
            Paragraph(
                "⚠️  <b>STRESS-TEST SCENARIO — NOT AN OFFICIAL RATING</b><br/>"
                "<font size='8'>The pillar scores, profiles, and indicative ratings in this report "
                "are derived from user-defined hypothetical inputs. They do not represent "
                "S&amp;P's actual assessment of the sovereign.</font>",
                S("Disc", fontSize=9, textColor=colors.HexColor("#7c2d12"),
                  fontName="Helvetica", leading=13)
            )
        ]]
        disclaimer_tbl = Table(disclaimer_data, colWidths=[17*cm])
        disclaimer_tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0,0), (-1,-1), colors.HexColor("#fff7ed")),
            ('BOX',           (0,0), (-1,-1), 1.2, colors.HexColor("#f97316")),
            ('LEFTPADDING',   (0,0), (-1,-1), 10),
            ('RIGHTPADDING',  (0,0), (-1,-1), 10),
            ('TOPPADDING',    (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(disclaimer_tbl)
        story.append(vsp(10))

        # Baseline vs scenario comparison table (if baseline provided)
        if baseline_scores:
            story.append(Paragraph("Baseline vs. Stress-Test Scenario", st_h3))
            bs = baseline_scores
            def _delta(base, sim):
                d = sim - base
                arrow = "▲" if d < 0 else ("▼" if d > 0 else "–")
                color = "#065f46" if d < 0 else ("#991b1b" if d > 0 else "#64748B")
                sign  = f"{d:+.1f}" if d != 0 else "–"
                return Paragraph(
                    f"<font color='{color}'><b>{arrow} {sign}</b></font>",
                    S("Dlt", fontSize=9, alignment=TA_CENTER)
                )

            cmp_data = [
                ["Pillar", "Baseline", "Simulated", "Change"],
                ["🏛 Institutional", f"{bs['s_inst']:.1f}", f"{s_inst:.1f}", _delta(bs['s_inst'], s_inst)],
                ["📈 Economic",      f"{bs['s_eco']:.1f}",  f"{s_eco:.1f}",  _delta(bs['s_eco'],  s_eco)],
                ["💰 Fiscal",        f"{bs['s_fis']:.1f}",  f"{s_fis:.1f}",  _delta(bs['s_fis'],  s_fis)],
                ["🌐 External",      f"{bs['s_ext']:.1f}",  f"{s_ext:.1f}",  _delta(bs['s_ext'],  s_ext)],
                ["🏦 Monetary",      f"{bs['s_mon']:.1f}",  f"{s_mon:.1f}",  _delta(bs['s_mon'],  s_mon)],
                ["SRM Rating",      bs['srm_rating'],       srm_rating,      ""],
                ["Final Rating",    bs['final_indicative'], final_indicative, ""],
            ]
            cmp_style = TableStyle([
                ('BACKGROUND',    (0,0), (-1,0), RL_NAVY),
                ('TEXTCOLOR',     (0,0), (-1,0), RL_WHITE),
                ('FONTNAME',      (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE',      (0,0), (-1,-1), 9),
                ('ALIGN',         (1,0), (-1,-1), 'CENTER'),
                ('BOX',           (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
                ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
                ('TOPPADDING',    (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ('BACKGROUND',    (0,6), (-1,7), colors.HexColor("#EEF2FF")),
                ('FONTNAME',      (0,6), (-1,7), 'Helvetica-Bold'),
            ])
            story.append(Table(cmp_data, colWidths=[5*cm, 3.5*cm, 3.5*cm, 3*cm], style=cmp_style))
            story.append(vsp(10))
    else:
        story.append(Paragraph(f"Sovereign Credit Rating Briefing", st_h1))
        story.append(Paragraph(f"<b>{target}</b> | S&amp;P Methodology Analysis", st_h2))
        story.append(Paragraph(
            f"Generated: {now_jakarta().strftime('%d %B %Y, %H:%M')} &nbsp;|&nbsp; "
            f"Based on S&amp;P Global Rating Criteria",
            st_small))
        story.append(hr())

    # ── SECTION 1 — NATIONAL PORTFOLIO SNAPSHOT ───────────────────────────────
    story.append(Paragraph("1. National Portfolio Snapshot", st_h2))

    actual_num = RATING_TO_NUM.get(str(r['Actual_Rating']).replace("*","").strip(), 8)
    srm_num    = RATING_TO_NUM.get(srm_rating, 8)

    def rating_badge(label, val, note=""):
        bg, fg = _rating_color(val)
        return [Paragraph(f"<b>{label}</b>", st_small),
                Paragraph(f"<font color='{val}'><b>{val}</b></font>", st_center) if False
                else Paragraph(f"<b>{val}</b>", S("RB", fontSize=16, alignment=TA_CENTER,
                                                    textColor=RL_NAVY, fontName="Helvetica-Bold")),
                Paragraph(note, st_caption)]

    # Summary KPI table
    qo_sign = f"{int(qo):+}"
    kpi_data = [
        ["S&P Actual Rating", "SRM Model Output", "Qualitative Overlay", "IE Profile", "FP Profile"],
        [str(r['Actual_Rating']), srm_rating, qo_sign,
         f"{prof_ie:.2f}", f"{prof_fp:.2f}"],
    ]
    kpi_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), RL_NAVY),
        ('TEXTCOLOR',  (0,0), (-1,0), RL_WHITE),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,0), 8),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME',   (0,1), (-1,1), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,1), (-1,1), 14),
        ('BACKGROUND', (0,1), (0,1), _rating_color(r['Actual_Rating'])[0]),
        ('BACKGROUND', (1,1), (1,1), _rating_color(srm_rating)[0]),
        ('ROWBACKGROUNDS', (0,1), (-1,1), [RL_LGREY]),
        ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('INNERGRID',  (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ])
    story.append(Table(kpi_data, colWidths=[3.2*cm]*5, style=kpi_style))
    story.append(vsp(8))

    # Pillar scorecard table
    story.append(Paragraph("Pillar Scorecard", st_h3))
    pillar_rows = [
        ["Pillar", "Score", "Profile", "Strength"],
        ["🏛 Institutional", f"{s_inst:.1f}", "IE", f"{(1-(s_inst-1)/5)*100:.0f}%"],
        ["📈 Economic",      f"{s_eco:.1f}",  "IE", f"{(1-(s_eco-1)/5)*100:.0f}%"],
        ["💰 Fiscal",        f"{s_fis:.1f}",  "FP", f"{(1-(s_fis-1)/5)*100:.0f}%"],
        ["🌐 External",      f"{s_ext:.1f}",  "FP", f"{(1-(s_ext-1)/5)*100:.0f}%"],
        ["🏦 Monetary",      f"{s_mon:.1f}",  "FP", f"{(1-(s_mon-1)/5)*100:.0f}%"],
    ]
    def pillar_bg(score):
        pct = (1-(float(score)-1)/5)*100
        return RL_GREEN if pct >= 60 else (RL_YELLOW if pct >= 40 else RL_RED)

    p_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), RL_NAVY),
        ('TEXTCOLOR',  (0,0), (-1,0), RL_WHITE),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 9),
        ('ALIGN',      (1,0), (-1,-1), 'CENTER'),
        ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('INNERGRID',  (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ])
    for i in range(1, 6):
        score = pillar_rows[i][1]
        p_style.add('BACKGROUND', (1,i), (1,i), pillar_bg(score))

    story.append(Table(pillar_rows,
                       colWidths=[5*cm, 2.5*cm, 2.5*cm, 3*cm],
                       style=p_style))
    story.append(vsp(6))

    # QO interpretation
    qo_label = "Upgrade" if qo > 0 else ("Penalty" if qo < 0 else "Neutral")
    qo_text  = (
        "Positive QO: S&P views qualitative factors as stronger than the SRM model implies — "
        "strong institutions, lower contingent liability risks, or superior fiscal flexibility."
        if qo > 0 else
        "Negative QO: S&P identified hidden risks not fully captured by the quantitative model — "
        "SOE/GRE liabilities, narrow fiscal space, external vulnerability, or governance concerns."
        if qo < 0 else
        "Neutral QO: The SRM model output aligns with S&P's full qualitative assessment."
    )
    story.append(Paragraph(f"<b>Qualitative Overlay ({qo_sign} notch — {qo_label}):</b> {qo_text}", st_body))
    story.append(hr())

    # ── SECTION 2 — PEER COMPARISON ───────────────────────────────────────────
    story.append(Paragraph("2. Peer Comparison", st_h2))

    if len(comp_list) > 1:
        # 2a. Pillar bar chart
        story.append(Paragraph("2a. Pillar Score Comparison", st_h3))
        chart_buf = _pillar_bar_chart(comp_list, width=14/2.54, height=7/2.54)
        story.append(RLImage(chart_buf, width=14*cm, height=7*cm))
        story.append(vsp(6))

        # 2b. Metrics snapshot table
        story.append(Paragraph("2b. Metrics Snapshot", st_h3))
        metrics_order = [
            ("Actual Rating", None),
            ("SRM (Model)", None),
            ("QO (Notch)", None),
            ("GDP per Capita (USD)", "GDP per Capita (USD)"),
            ("Real GDP Growth (%)", "Real GDP Growth (%)"),
            ("Fiscal Balance (% GDP)", "Fiscal Balance (% GDP)"),
            ("Debt-to-GDP (%)", "Debt-to-GDP (%)"),
            ("Interest/Revenue (%)", "Interest/Revenue (%)"),
            ("GEFN (% CAR)", "GEFN (% CAR)"),
            ("NIIP (% GDP)", "NIIP (% GDP)"),
            ("Reserves (months)", "Reserves (months)"),
            ("CPI Inflation (%)", "CPI Inflation (%)"),
            ("Financial Depth (% GDP)", "Financial Depth (% GDP)"),
        ]

        hdr = ["Indicator"] + [c["Country"] for c in comp_list]
        snap_data = [hdr]

        def fmt_metric(label, c):
            nr = c["_nr"]
            MAP = {
                "Actual Rating":          str(nr['Actual_Rating']),
                "SRM (Model)":            c["_srm"],
                "QO (Notch)":             f"{int(c['_qo']):+}",
                "GDP per Capita (USD)":   f"${nr['GDP_PC']:,.0f}",
                "Real GDP Growth (%)":    f"{nr['Growth']:.1f}%",
                "Fiscal Balance (% GDP)": f"{nr['Balance']:.1f}%",
                "Debt-to-GDP (%)":        f"{nr['Debt_GDP']:.1f}%",
                "Interest/Revenue (%)":   f"{nr['Int_Rev']:.1f}%",
                "GEFN (% CAR)":           f"{nr['GEFN']:.1f}%",
                "NIIP (% GDP)":           f"{nr['NIIP_CAR']:.1f}%",
                "Reserves (months)":      f"{nr['Reserves']:.1f} mo",
                "CPI Inflation (%)":      f"{nr['CPI']:.1f}%",
                "Financial Depth (% GDP)": f"{nr['Fin_Depth']:.1f}%",
            }
            return MAP.get(label, "N/A")

        for label, thresh_key in metrics_order:
            row = [label]
            for c in comp_list:
                row.append(fmt_metric(label, c))
            snap_data.append(row)

        n_cols = len(comp_list) + 1
        col_w  = [4.5*cm] + [13.5/len(comp_list)*cm] * len(comp_list)
        snap_style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), RL_NAVY),
            ('TEXTCOLOR',  (0,0), (-1,0), RL_WHITE),
            ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0), (-1,-1), 8),
            ('ALIGN',      (1,0), (-1,-1), 'CENTER'),
            ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ('INNERGRID',  (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [RL_WHITE, RL_LGREY]),
        ])
        # colour rating cells
        for ri, (label, thresh_key) in enumerate(metrics_order, 1):
            for ci, c in enumerate(comp_list, 1):
                nr = c["_nr"]
                if label in ("Actual Rating", "SRM (Model)"):
                    val = fmt_metric(label, c)
                    bg, fg = _rating_color(val)
                    snap_style.add('BACKGROUND', (ci, ri), (ci, ri), bg)
                elif thresh_key:
                    raw_map = {
                        "GDP per Capita (USD)":   nr['GDP_PC'],
                        "Real GDP Growth (%)":    nr['Growth'],
                        "Fiscal Balance (% GDP)": nr['Balance'],
                        "Debt-to-GDP (%)":        nr['Debt_GDP'],
                        "Interest/Revenue (%)":   nr['Int_Rev'],
                        "GEFN (% CAR)":           nr['GEFN'],
                        "NIIP (% GDP)":           nr['NIIP_CAR'],
                        "Reserves (months)":      nr['Reserves'],
                        "CPI Inflation (%)":      nr['CPI'],
                        "Financial Depth (% GDP)": nr['Fin_Depth'],
                    }
                    raw = raw_map.get(thresh_key)
                    if raw is not None:
                        bg, fg = _metric_color(thresh_key, raw)
                        snap_style.add('BACKGROUND', (ci, ri), (ci, ri), bg)

        story.append(Table(snap_data, colWidths=col_w, style=snap_style,
                           repeatRows=1))
        story.append(vsp(6))

        # 2c. Historical trends (2-column grid)
        if trend_df is not None and not trend_df.empty and selected_metrics:
            story.append(Paragraph("2c. Historical & Forecast Trends", st_h3))
            story.append(Paragraph(
                "Shaded area = forecast/estimate period. "
                "Dashed line = non-primary countries.",
                st_small))
            story.append(vsp(4))

            charts = []
            for metric in selected_metrics:
                cbuf = _trend_line_chart(trend_df, metric, sel_nations,
                                         width=6.5/2.54, height=4.5/2.54)
                if cbuf:
                    charts.append((metric, cbuf))

            # Arrange in 2-column pairs
            for i in range(0, len(charts), 2):
                pair = charts[i:i+2]
                row_imgs = []
                for _, cbuf in pair:
                    img = RLImage(cbuf, width=8*cm, height=5.5*cm)
                    row_imgs.append(img)
                if len(row_imgs) == 1:
                    row_imgs.append(Spacer(8*cm, 5.5*cm))
                t = Table([row_imgs], colWidths=[8.5*cm, 8.5*cm])
                t.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),
                                       ('LEFTPADDING',(0,0),(-1,-1),2),
                                       ('RIGHTPADDING',(0,0),(-1,-1),2)]))
                story.append(t)
                story.append(vsp(4))

    story.append(hr())

    # ── SECTION 3 — RECOMMENDATION ────────────────────────────────────────────
    story.append(Paragraph("3. Recommendation", st_h2))

    strength_sigs  = [s for s in signals if s["type"] == "strength"]
    moderate_sigs  = [s for s in signals if s["type"] == "moderate"]
    weakness_sigs  = [s for s in signals if s["type"] == "weakness"]

    def signal_table(sig_list, bg, title):
        if not sig_list: return
        story.append(Paragraph(title, st_h3))
        rows = [["Pillar", "Metric", "Value", "Assessment", "Action"]]
        for s in sig_list:
            rows.append([
                s["pillar"], s["metric"], s["value"],
                Paragraph(s["message"][:120], st_small),
                Paragraph(s["action"][:120], st_small),
            ])
        ts = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), RL_NAVY),
            ('TEXTCOLOR',  (0,0), (-1,0), RL_WHITE),
            ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [bg, RL_WHITE]),
            ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ('INNERGRID',  (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
            ('VALIGN',     (0,0), (-1,-1), 'TOP'),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ])
        story.append(Table(rows,
            colWidths=[2*cm, 2.5*cm, 1.8*cm, 5.5*cm, 5.5*cm],
            style=ts, repeatRows=1))
        story.append(vsp(6))

    signal_table(weakness_sigs,  RL_RED,    "❌ Key Weaknesses")
    signal_table(moderate_sigs,  RL_YELLOW, "⚠️ Watch Areas")
    signal_table(strength_sigs,  RL_GREEN,  "✅ Key Strengths")
    story.append(hr())

    # ── SECTION 4 — QO DEEP DIVE ──────────────────────────────────────────────
    story.append(PageBreak())
    story.append(Paragraph("4. In-Depth Qualitative Overlay (QO) Targeting", st_h2))
    story.append(Paragraph(
        "The Qualitative Overlay reflects analyst judgment applied after the SRM produces an "
        "indicative rating. It covers five factors: Contingent Liabilities, Monetary Flexibility, "
        "External Liquidity &amp; IIP, Fiscal Flexibility, and Political &amp; Governance Risks. "
        "Each can adjust the rating by up to ±1–2 notches. The net QO for "
        f"<b>{target}</b> is <b>{qo_sign} notch ({qo_label})</b>.",
        st_body))
    story.append(vsp(6))

    # ── Peer-enhanced QO narrative ───────────────────────────────────────────
    # Use full df_master for BBB peer group — falls back to comp_list if not provided
    _peer_df = df_master if df_master is not None and not df_master.empty else \
               pd.DataFrame([{
                   "Country": c["Country"],
                   "Actual_Rating": c["_nr"]["Actual_Rating"]
               } for c in comp_list]) if comp_list else pd.DataFrame()

    _, _, peer_countries_qo, comp_rows_qo = _compute_peer_qo_data(
        target, trend_df, _peer_df, RATING_TO_NUM
    )

    strengths_qo  = [r2 for r2 in comp_rows_qo if r2["is_better"]]
    weaknesses_qo = [r2 for r2 in comp_rows_qo if r2["is_worse"]]
    strengths_qo  = sorted(strengths_qo, key=lambda x: x["pct_diff"], reverse=True)

    st_cell9 = ParagraphStyle("c9", fontSize=9, leading=12)

    if comp_rows_qo:
        # ── Peer comparison table ─────────────────────────────────────────────
        story.append(Paragraph(
            f"Performance vs BBB Peer Average ({len(peer_countries_qo)} peers, latest estimate)",
            st_h3))

        tbl_hdr = ["Indicator", target, "BBB Peer Avg", "Difference", "Signal"]
        tbl_rows = [tbl_hdr]
        row_colors = []
        for r2 in comp_rows_qo:
            sig = "Above peers" if r2["is_better"] else ("Below peers" if r2["is_worse"] else "In line")
            tbl_rows.append([
                Paragraph(r2["label"], st_cell9),
                Paragraph(f"{r2['t_val']:.1f}", st_cell9),
                Paragraph(f"{r2['p_avg']:.1f}", st_cell9),
                Paragraph(f"{r2['diff']:+.1f}", st_cell9),
                Paragraph(sig, st_cell9),
            ])
            row_colors.append(r2["is_better"])

        peer_tbl_style = TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), RL_NAVY),
            ("TEXTCOLOR",     (0,0), (-1,0), RL_WHITE),
            ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,-1), 9),
            ("ALIGN",         (1,0), (-1,-1), "CENTER"),
            ("BOX",           (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ])
        for ri, r2 in enumerate(comp_rows_qo, 1):
            if r2["is_better"]:
                peer_tbl_style.add("BACKGROUND", (1,ri), (1,ri), RL_GREEN)
                peer_tbl_style.add("BACKGROUND", (3,ri), (3,ri), RL_GREEN)
                peer_tbl_style.add("BACKGROUND", (4,ri), (4,ri), RL_GREEN)
            elif r2["is_worse"]:
                peer_tbl_style.add("BACKGROUND", (1,ri), (1,ri), RL_RED)
                peer_tbl_style.add("BACKGROUND", (3,ri), (3,ri), RL_RED)
                peer_tbl_style.add("BACKGROUND", (4,ri), (4,ri), RL_RED)

        story.append(Table(tbl_rows,
            colWidths=[5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 5*cm],
            style=peer_tbl_style, repeatRows=1))
        story.append(vsp(8))

        # ── Ranked QO arguments ───────────────────────────────────────────────
        if strengths_qo:
            story.append(Paragraph("Ranked QO Arguments (Strongest First)", st_h3))
            for i, r2 in enumerate(strengths_qo[:5], 1):
                pct = r2["pct_diff"]
                if pct >= 30:   strength_label = "Very Strong"; bg_col = colors.HexColor("#d1fae5")
                elif pct >= 15: strength_label = "Strong";      bg_col = colors.HexColor("#dbeafe")
                else:           strength_label = "Supporting";  bg_col = colors.HexColor("#ede9fe")

                arg_block = [
                    [Paragraph(
                        f"<b>#{i} {r2['factor']}</b> | {r2['label']} | "
                        f"{target}: {r2['t_val']:.1f} vs BBB avg: {r2['p_avg']:.1f} "
                        f"({r2['diff']:+.1f}, {pct:.0f}% outperformance) | {strength_label}",
                        ParagraphStyle("arg_h", fontSize=9, fontName="Helvetica-Bold",
                                       textColor=colors.HexColor("#1e3a8a")))],
                    [Paragraph(r2["argument"],
                        ParagraphStyle("arg_b", fontSize=9, leading=13,
                                       textColor=colors.HexColor("#1e293b")))],
                ]
                arg_tbl = Table(arg_block, colWidths=[17.4*cm])
                arg_tbl.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0), (-1,0), bg_col),
                    ("BACKGROUND",    (0,1), (-1,1), colors.HexColor("#f8fafc")),
                    ("BOX",           (0,0), (-1,-1), 0.8, colors.HexColor("#1e3a8a")),
                    ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.HexColor("#e2e8f0")),
                    ("TOPPADDING",    (0,0), (-1,-1), 5),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                    ("LEFTPADDING",   (0,0), (-1,-1), 6),
                    ("RIGHTPADDING",  (0,0), (-1,-1), 6),
                ]))
                story.append(arg_tbl)
                story.append(vsp(4))

            # Overall QO strategy box
            top_factors = list(dict.fromkeys([r2["factor"] for r2 in strengths_qo[:3]]))
            strategy = (
                f"Lead with {strengths_qo[0]['label']} advantage "
                f"({strengths_qo[0]['t_val']:.1f} vs peer avg {strengths_qo[0]['p_avg']:.1f}, "
                f"{strengths_qo[0]['pct_diff']:.0f}% outperformance). "
                f"Key QO factors to target: {', '.join(top_factors)}. "
                + ("Address underperforming metrics to protect QO position." if weaknesses_qo
                   else "No material peer underperformance — strong overall QO case.")
            )
            strat_tbl = Table(
                [[Paragraph(f"<b>🎯 Overall QO Strategy:</b> {strategy}",
                    ParagraphStyle("strat", fontSize=9, leading=13,
                                   textColor=colors.white))]],
                colWidths=[17.4*cm])
            strat_tbl.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), RL_NAVY),
                ("TOPPADDING",    (0,0), (-1,-1), 8),
                ("BOTTOMPADDING", (0,0), (-1,-1), 8),
                ("LEFTPADDING",   (0,0), (-1,-1), 10),
                ("RIGHTPADDING",  (0,0), (-1,-1), 10),
                ("BOX",           (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ]))
            story.append(strat_tbl)
            story.append(vsp(8))

        # ── Underperforming metrics ───────────────────────────────────────────
        if weaknesses_qo:
            story.append(Paragraph(
                f"Areas Where {target} Underperforms BBB Peers (Address to Protect QO)",
                st_h3))
            wk_rows = [["Indicator", target, "BBB Avg", "Gap", "Risk"]]
            for r2 in weaknesses_qo:
                wk_rows.append([
                    Paragraph(r2["label"], st_cell9),
                    Paragraph(f"{r2['t_val']:.1f}", st_cell9),
                    Paragraph(f"{r2['p_avg']:.1f}", st_cell9),
                    Paragraph(f"{r2['diff']:+.1f}", st_cell9),
                    Paragraph("Negative residual risk", st_cell9),
                ])
            wk_style = TableStyle([
                ("BACKGROUND",    (0,0), (-1,0), colors.HexColor("#991b1b")),
                ("TEXTCOLOR",     (0,0), (-1,0), RL_WHITE),
                ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE",      (0,0), (-1,-1), 9),
                ("ROWBACKGROUNDS",(0,1), (-1,-1), [RL_RED, RL_WHITE]),
                ("BOX",           (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
                ("INNERGRID",     (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
                ("TOPPADDING",    (0,0), (-1,-1), 4),
                ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ])
            story.append(Table(wk_rows,
                colWidths=[5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 5*cm],
                style=wk_style, repeatRows=1))
            story.append(vsp(6))

    else:
        # Fallback to static opportunities if no trend data
        if qo_opportunities:
            story.append(Paragraph("QO Narrative Opportunities", st_h3))
            opp_rows = [["Factor", "Argument", "Potential Impact"]]
            for opp in qo_opportunities:
                opp_rows.append([
                    Paragraph(f"<b>{opp['factor']}</b>", st_small),
                    Paragraph(opp['argument'], st_small),
                    Paragraph(f"<b>{opp['impact']}</b>", st_small),
                ])
            opp_style = TableStyle([
                ("BACKGROUND", (0,0), (-1,0), RL_TEAL),
                ("TEXTCOLOR",  (0,0), (-1,0), RL_WHITE),
                ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE",   (0,0), (-1,-1), 8),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.HexColor("#EFF6FF"), RL_WHITE]),
                ("BOX",        (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
                ("INNERGRID",  (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
                ("VALIGN",     (0,0), (-1,-1), "TOP"),
                ("TOPPADDING", (0,0), (-1,-1), 4),
                ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ])
            story.append(Table(opp_rows,
                colWidths=[3.5*cm, 11*cm, 3*cm],
                style=opp_style, repeatRows=1))
            story.append(vsp(8))
        else:
            story.append(Paragraph(
                "No QO narrative opportunities identified. "
                "Visit the Peer Comparison tab to load peer data for enhanced analysis.",
                st_body))
            story.append(vsp(6))

    # ── Selected Peer QO Context ──────────────────────────────────────────────
    if len(comp_list) > 1:
        story.append(Paragraph("Selected Peer QO Comparison", st_h3))
        story.append(Paragraph(
            "The table below shows the QO (residual ±1 notch) for each manually "
            "selected peer country alongside their SRM output and profile scores.",
            st_body))
        story.append(vsp(4))

        peer_rows = [["Country", "Actual", "SRM", "Post-Suppl.", "Residual QO", "IE Prof.", "FP Prof."]]
        for c in comp_list:
            peer_rows.append([
                c["Country"],
                c["_nr"]['Actual_Rating'],
                c["_srm"],
                c["_srm"],   # post-supplemental not computed for peers — use SRM
                f"{int(c['_qo']):+}",
                f"{c['_ie']:.2f}",
                f"{c['_fp']:.2f}",
            ])

        peer_style = TableStyle([
            ('BACKGROUND',    (0,0), (-1,0), RL_NAVY),
            ('TEXTCOLOR',     (0,0), (-1,0), RL_WHITE),
            ('FONTNAME',      (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',      (0,0), (-1,-1), 9),
            ('ALIGN',         (1,0), (-1,-1), 'CENTER'),
            ('BOX',           (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
            ('ROWBACKGROUNDS',(0,1), (-1,-1), [RL_LGREY, RL_WHITE]),
            ('TOPPADDING',    (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ])

        # Colour rating and QO cells
        for ri, c in enumerate(comp_list, 1):
            # Actual rating colour
            rbg, _ = _rating_color(c["_nr"]['Actual_Rating'])
            peer_style.add('BACKGROUND', (1, ri), (1, ri), rbg)
            # SRM colour
            sbg, _ = _rating_color(c["_srm"])
            peer_style.add('BACKGROUND', (2, ri), (2, ri), sbg)
            # QO colour
            qo_val = int(c['_qo'])
            qo_bg = RL_GREEN if qo_val > 0 else (RL_RED if qo_val < 0 else RL_WHITE)
            peer_style.add('BACKGROUND', (4, ri), (4, ri), qo_bg)

        story.append(Table(peer_rows,
            colWidths=[3.5*cm, 2*cm, 2*cm, 2.5*cm, 2.5*cm, 2.2*cm, 2.2*cm],
            style=peer_style))
        story.append(vsp(6))

        # QO delta narrative
        target_qo   = int(comp_list[0]["_qo"]) if comp_list else 0
        peer_qos    = [(c["Country"], int(c["_qo"])) for c in comp_list[1:]]
        higher_peers = [f"{n} ({q:+})" for n, q in peer_qos if q > target_qo]
        lower_peers  = [f"{n} ({q:+})" for n, q in peer_qos if q < target_qo]
        equal_peers  = [f"{n} ({q:+})" for n, q in peer_qos if q == target_qo]

        if higher_peers:
            story.append(Paragraph(
                f"<b>Peers with higher residual QO:</b> {', '.join(higher_peers)}. "
                f"These peers receive more favourable qualitative adjustments than {target}, "
                "suggesting room to strengthen the narrative on institutional quality, "
                "contingent liability management, or fiscal flexibility.",
                st_body))
            story.append(vsp(4))

        if lower_peers:
            story.append(Paragraph(
                f"<b>Peers with lower residual QO:</b> {', '.join(lower_peers)}. "
                f"{target}'s relative QO position is stronger than these peers — "
                "highlight this comparative advantage in rating dialogue.",
                st_body))
            story.append(vsp(4))

        if equal_peers:
            story.append(Paragraph(
                f"<b>Peers with equal residual QO:</b> {', '.join(equal_peers)}. "
                "S&P applies similar qualitative judgment to these sovereigns.",
                st_body))
            story.append(vsp(4))

    # Footer
    story.append(vsp(16))
    story.append(hr())


    # Footer
    story.append(vsp(16))
    story.append(hr())
    story.append(Paragraph(
        f"<i>Confidential — Sovereign Rating Monitoring Dashboard | "
        f"Generated {now_jakarta().strftime('%d %B %Y')}</i>",
        st_small))

    doc.build(story)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# PPTX GENERATOR
# ─────────────────────────────────────────────────────────────────────────────

def generate_pptx(target, r, srm_rating, qo,
                  s_inst, s_eco, s_fis, s_ext, s_mon,
                  prof_ie, prof_fp,
                  comp_list, sel_nations, trend_df, selected_metrics,
                  signals, qo_opportunities,
                  supp_adj=0, supp_factors=None, final_indicative=None, cap_note=None,
                  df_master=None):
    supp_factors     = supp_factors or []
    final_indicative = final_indicative or srm_rating
    """Returns a bytes buffer containing the PPTX briefing."""

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    BLANK = 6   # blank layout index
    def blank_slide():
        return prs.slides.add_slide(prs.slide_layouts[BLANK])

    def rgb(hex_str):
        h = hex_str.lstrip("#")
        return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

    def add_rect(slide, x, y, w, h, fill_hex, line_hex=None, line_w=Pt(0)):
        from pptx.util import Inches
        shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill_hex)
        if line_hex:
            shape.line.color.rgb = rgb(line_hex)
            shape.line.width = line_w
        else:
            shape.line.fill.background()
        return shape

    def add_text(slide, text, x, y, w, h, size=12, bold=False,
                 color="#1E293B", align=PP_ALIGN.LEFT, wrap=True, italic=False):
        txb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        txb.word_wrap = wrap
        tf = txb.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = str(text)
        run.font.size  = Pt(size)
        run.font.bold  = bold
        run.font.italic = italic
        run.font.color.rgb = rgb(color)
        return txb

    def add_image_buf(slide, buf, x, y, w, h):
        buf.seek(0)
        slide.shapes.add_picture(buf, Inches(x), Inches(y), Inches(w), Inches(h))

    def rating_fill(rating):
        num = RATING_TO_NUM.get(str(rating).replace("*","").strip(), -1)
        if num >= 10: return "d1fae5"
        if num >= 6:  return "fef3c7"
        return "fee2e2"

    date_str = datetime.now().strftime("%d %B %Y")
    qo_sign  = f"{int(qo):+}"
    qo_label = "Upgrade" if qo > 0 else ("Penalty" if qo < 0 else "Neutral")

    # ── SLIDE 1: COVER ────────────────────────────────────────────────────────
    sl = blank_slide()
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb("1E3A8A")
    add_rect(sl, 0, 5.8, 13.33, 1.7, "0D9488")
    add_text(sl, "SOVEREIGN CREDIT RATING BRIEFING",
             0.6, 1.2, 12, 0.6, size=16, bold=True, color="CADCFC", align=PP_ALIGN.LEFT)
    add_text(sl, target,
             0.6, 2.0, 12, 1.2, size=48, bold=True, color="FFFFFF", align=PP_ALIGN.LEFT)
    add_text(sl, "S&P Global Methodology Analysis",
             0.6, 3.4, 10, 0.5, size=16, color="94A3B8", align=PP_ALIGN.LEFT)
    add_text(sl, f"Generated: {date_str}",
             0.6, 6.1, 8, 0.4, size=12, color="FFFFFF", align=PP_ALIGN.LEFT)
    # Rating badge
    add_rect(sl, 10, 1.8, 2.6, 2.2, rating_fill(r['Actual_Rating']), line_hex="CBD5E1", line_w=Pt(1))
    add_text(sl, "S&P Rating", 10, 1.9, 2.6, 0.35, size=9, color="64748B", align=PP_ALIGN.CENTER)
    add_text(sl, str(r['Actual_Rating']), 10, 2.2, 2.6, 0.9,
             size=40, bold=True, color="1E3A8A", align=PP_ALIGN.CENTER)
    add_text(sl, f"SRM: {srm_rating}  |  QO: {qo_sign}",
             10, 3.1, 2.6, 0.4, size=10, color="475569", align=PP_ALIGN.CENTER)

    # ── SLIDE 2: NATIONAL PORTFOLIO ───────────────────────────────────────────
    sl = blank_slide()
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb("F8FAFC")
    add_rect(sl, 0, 0, 13.33, 0.7, "1E3A8A")
    add_text(sl, f"1. National Portfolio — {target}",
             0.3, 0.1, 12, 0.5, size=14, bold=True, color="FFFFFF")

    # KPI cards row
    kpis = [
        ("Actual Rating", str(r['Actual_Rating']), rating_fill(r['Actual_Rating'])),
        ("SRM Output",    srm_rating,               rating_fill(srm_rating)),
        ("QO Notch",      qo_sign,                  "d1fae5" if qo > 0 else ("fee2e2" if qo < 0 else "F1F5F9")),
        ("IE Profile",    f"{prof_ie:.2f}",          "EFF6FF"),
        ("FP Profile",    f"{prof_fp:.2f}",          "EFF6FF"),
    ]
    for i, (label, val, bg) in enumerate(kpis):
        x = 0.3 + i * 2.55
        add_rect(sl, x, 0.85, 2.4, 1.1, bg, line_hex="CBD5E1", line_w=Pt(0.5))
        add_text(sl, label, x, 0.9, 2.4, 0.3, size=8, color="64748B", align=PP_ALIGN.CENTER)
        add_text(sl, val,   x, 1.15, 2.4, 0.65, size=26, bold=True,
                 color="1E3A8A", align=PP_ALIGN.CENTER)

    # Pillar scores
    pillar_data = [
        ("🏛 Inst.", s_inst), ("📈 Econ.", s_eco), ("💰 Fiscal", s_fis),
        ("🌐 Ext.", s_ext), ("🏦 Mon.", s_mon),
    ]
    add_text(sl, "Pillar Scores (1 = Best, 6 = Worst)",
             0.3, 2.1, 10, 0.3, size=10, bold=True, color="1E3A8A")
    for i, (name, score) in enumerate(pillar_data):
        x = 0.3 + i * 2.55
        pct = (1-(score-1)/5)
        col = "065f46" if pct >= 0.6 else ("78350f" if pct >= 0.4 else "991b1b")
        bg2 = "d1fae5" if pct >= 0.6 else ("fef3c7" if pct >= 0.4 else "fee2e2")
        add_rect(sl, x, 2.45, 2.4, 0.9, bg2, line_hex="CBD5E1", line_w=Pt(0.5))
        add_text(sl, name,           x, 2.5,  2.4, 0.3, size=8,  color="475569", align=PP_ALIGN.CENTER)
        add_text(sl, f"{score:.1f}", x, 2.75, 2.4, 0.45, size=22, bold=True,
                 color=col, align=PP_ALIGN.CENTER)

    # QO note
    qo_text_short = (
        f"QO {qo_sign} ({qo_label}): " + (
        "S&P views qualitative factors as STRONGER than SRM implies." if qo > 0 else
        "S&P identified hidden risks not captured by the quantitative model." if qo < 0 else
        "SRM model aligns with full qualitative assessment — no material adjustments."
        ))
    add_rect(sl, 0.3, 3.55, 12.7, 0.6,
             "d1fae5" if qo > 0 else ("fee2e2" if qo < 0 else "F1F5F9"),
             line_hex="CBD5E1", line_w=Pt(0.5))
    add_text(sl, qo_text_short, 0.4, 3.6, 12.5, 0.5, size=9,
             color="065f46" if qo > 0 else ("991b1b" if qo < 0 else "475569"))

    # ── SLIDE 3: PEER COMPARISON — PILLAR SCORES ──────────────────────────────
    if len(comp_list) > 1:
        sl = blank_slide()
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = rgb("F8FAFC")
        add_rect(sl, 0, 0, 13.33, 0.7, "1E3A8A")
        add_text(sl, "2a. Peer Comparison — Pillar Scores",
                 0.3, 0.1, 12, 0.5, size=14, bold=True, color="FFFFFF")

        chart_buf = _pillar_bar_chart(comp_list, width=16/2.54, height=10/2.54)
        add_image_buf(sl, chart_buf, 0.4, 0.85, 12.5, 5.8)

        # ── SLIDE 4: PEER COMPARISON — METRICS TABLE ──────────────────────────
        sl = blank_slide()
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = rgb("F8FAFC")
        add_rect(sl, 0, 0, 13.33, 0.7, "1E3A8A")
        add_text(sl, "2b. Peer Comparison — Metrics Snapshot",
                 0.3, 0.1, 12, 0.5, size=14, bold=True, color="FFFFFF")

        # Build table as image via matplotlib
        metrics_labels = [
            "Actual Rating", "SRM", "QO",
            "GDP/Capita", "Growth", "Fiscal Bal.", "Debt/GDP",
            "Int/Rev", "GEFN", "NIIP", "Reserves", "CPI", "Fin.Depth"
        ]
        def get_val(label, c):
            nr = c["_nr"]
            MAP = {
                "Actual Rating": str(nr['Actual_Rating']),
                "SRM": c["_srm"], "QO": f"{int(c['_qo']):+}",
                "GDP/Capita": f"${nr['GDP_PC']:,.0f}",
                "Growth": f"{nr['Growth']:.1f}%",
                "Fiscal Bal.": f"{nr['Balance']:.1f}%",
                "Debt/GDP": f"{nr['Debt_GDP']:.1f}%",
                "Int/Rev": f"{nr['Int_Rev']:.1f}%",
                "GEFN": f"{nr['GEFN']:.1f}%",
                "NIIP": f"{nr['NIIP_CAR']:.1f}%",
                "Reserves": f"{nr['Reserves']:.1f}mo",
                "CPI": f"{nr['CPI']:.1f}%",
                "Fin.Depth": f"{nr['Fin_Depth']:.1f}%",
            }
            return MAP.get(label, "N/A")

        fig, ax = plt.subplots(figsize=(13, 5.5))
        ax.axis('off')
        col_labels = [c["Country"] for c in comp_list]
        cell_data  = [[get_val(m, c) for c in comp_list] for m in metrics_labels]

        tbl = ax.table(
            cellText=cell_data,
            rowLabels=metrics_labels,
            colLabels=col_labels,
            loc='center', cellLoc='center'
        )
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(9)
        tbl.scale(1, 1.35)

        THRESH_MAP = {
            "Growth":     (True,  [4.0, 1.0]),
            "Fiscal Bal.":(True,  [-3.0, -5.0]),
            "Debt/GDP":   (False, [45, 60]),
            "Int/Rev":    (False, [15, 25]),
            "GEFN":       (False, [50, 75]),
            "NIIP":       (True,  [0, -50]),
            "Reserves":   (True,  [6.0, 3.0]),
            "CPI":        (False, [3.0, 6.0]),
            "Fin.Depth":  (True,  [40, 20]),
        }

        for ri, m in enumerate(metrics_labels):
            for ci, c in enumerate(comp_list):
                cell = tbl[ri+1, ci]
                nr = c["_nr"]
                if m in ("Actual Rating","SRM"):
                    val_str = get_val(m, c)
                    num = RATING_TO_NUM.get(val_str.replace("*","").strip(), -1)
                    if num >= 10: cell.set_facecolor("#d1fae5")
                    elif num >= 6: cell.set_facecolor("#fef3c7")
                    else: cell.set_facecolor("#fee2e2")
                elif m in THRESH_MAP:
                    raw_map = {
                        "Growth": nr['Growth'], "Fiscal Bal.": nr['Balance'],
                        "Debt/GDP": nr['Debt_GDP'], "Int/Rev": nr['Int_Rev'],
                        "GEFN": nr['GEFN'], "NIIP": nr['NIIP_CAR'],
                        "Reserves": nr['Reserves'], "CPI": nr['CPI'],
                        "Fin.Depth": nr['Fin_Depth'],
                    }
                    raw = raw_map.get(m)
                    if raw is not None:
                        higher, (g, y) = THRESH_MAP[m]
                        try:
                            v = float(raw)
                            if higher:
                                fc = "#d1fae5" if v >= g else ("#fef3c7" if v >= y else "#fee2e2")
                            else:
                                fc = "#d1fae5" if v <= g else ("#fef3c7" if v <= y else "#fee2e2")
                            cell.set_facecolor(fc)
                        except: pass

        last_row = len(metrics_labels)  # 0=header, 1..N=data rows
        for ci in range(len(comp_list)):
            tbl[last_row, ci].set_facecolor("#DBEAFE")
            tbl[last_row, ci].set_text_props(fontweight='bold')

        fig.tight_layout()
        tbuf = io.BytesIO()
        fig.savefig(tbuf, format="png", dpi=150, bbox_inches="tight")
        plt.close(fig)
        tbuf.seek(0)
        add_image_buf(sl, tbuf, 0.3, 0.8, 12.7, 6.2)

        # ── SLIDE 5+: TREND CHARTS (2 per slide) ──────────────────────────────
        if trend_df is not None and not trend_df.empty and selected_metrics:
            charts = []
            for metric in selected_metrics:
                cbuf = _trend_line_chart(trend_df, metric, sel_nations,
                                         width=6.5/2.54, height=5/2.54)
                if cbuf:
                    charts.append((metric, cbuf))

            for i in range(0, len(charts), 2):
                sl = blank_slide()
                sl.background.fill.solid()
                sl.background.fill.fore_color.rgb = rgb("F8FAFC")
                add_rect(sl, 0, 0, 13.33, 0.7, "1E3A8A")
                slide_num = i // 2 + 1
                add_text(sl, f"2c. Historical & Forecast Trends (Part {slide_num})",
                         0.3, 0.1, 12, 0.5, size=14, bold=True, color="FFFFFF")

                pair = charts[i:i+2]
                for j, (_, cbuf) in enumerate(pair):
                    x_pos = 0.3 + j * 6.5
                    add_image_buf(sl, cbuf, x_pos, 0.85, 6.3, 6.2)

    # ── SLIDE: RECOMMENDATION ─────────────────────────────────────────────────
    sl = blank_slide()
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb("F8FAFC")
    add_rect(sl, 0, 0, 13.33, 0.7, "1E3A8A")
    add_text(sl, "3. Recommendation — Strengths & Weaknesses",
             0.3, 0.1, 12, 0.5, size=14, bold=True, color="FFFFFF")

    weakness_sigs = [s for s in signals if s["type"] == "weakness"]
    strength_sigs = [s for s in signals if s["type"] == "strength"]

    # Left: weaknesses
    add_rect(sl, 0.2, 0.8, 6.2, 0.4, "fee2e2")
    add_text(sl, f"❌ Key Weaknesses ({len(weakness_sigs)})",
             0.3, 0.83, 6, 0.35, size=11, bold=True, color="991b1b")
    y_w = 1.3
    for s in weakness_sigs[:5]:
        add_rect(sl, 0.2, y_w, 6.2, 0.62, "fff5f5", line_hex="fca5a5", line_w=Pt(0.5))
        add_text(sl, f"[{s['pillar']}] {s['metric']} — {s['value']}",
                 0.3, y_w + 0.02, 6.0, 0.25, size=8, bold=True, color="991b1b")
        add_text(sl, s["action"][:90],
                 0.3, y_w + 0.27, 6.0, 0.3, size=7.5, color="374151")
        y_w += 0.68

    # Right: strengths
    add_rect(sl, 6.9, 0.8, 6.2, 0.4, "d1fae5")
    add_text(sl, f"✅ Key Strengths ({len(strength_sigs)})",
             7.0, 0.83, 6, 0.35, size=11, bold=True, color="065f46")
    y_s = 1.3
    for s in strength_sigs[:5]:
        add_rect(sl, 6.9, y_s, 6.2, 0.62, "f0fdf4", line_hex="6ee7b7", line_w=Pt(0.5))
        add_text(sl, f"[{s['pillar']}] {s['metric']} — {s['value']}",
                 7.0, y_s + 0.02, 6.0, 0.25, size=8, bold=True, color="065f46")
        add_text(sl, s["action"][:90],
                 7.0, y_s + 0.27, 6.0, 0.3, size=7.5, color="374151")
        y_s += 0.68

    # ── SLIDE: QO IN-DEPTH — Peer Comparison Table ───────────────────────────
    # Use full df_master for BBB peer group
    _peer_df2 = df_master if df_master is not None and not df_master.empty else \
                pd.DataFrame([{
                    "Country": c["Country"],
                    "Actual_Rating": c["_nr"]["Actual_Rating"]
                } for c in comp_list]) if comp_list else pd.DataFrame()

    _, _, peer_countries_pptx, comp_rows_pptx = _compute_peer_qo_data(
        target, trend_df, _peer_df2, RATING_TO_NUM)

    strengths_pptx  = sorted([r2 for r2 in comp_rows_pptx if r2["is_better"]],
                               key=lambda x: x["pct_diff"], reverse=True)
    weaknesses_pptx = [r2 for r2 in comp_rows_pptx if r2["is_worse"]]

    sl = blank_slide()
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb("F8FAFC")
    add_rect(sl, 0, 0, 13.33, 0.7, "0D9488")
    add_text(sl, "4a. QO Targeting — Performance vs BBB Peers (Latest Estimate)",
             0.3, 0.1, 12, 0.5, size=13, bold=True, color="FFFFFF")

    # Rating derivation banner
    add_rect(sl, 0.2, 0.78, 12.9, 0.42,
             "d1fae5" if qo > 0 else ("fee2e2" if qo < 0 else "EFF6FF"),
             line_hex="CBD5E1", line_w=Pt(0.5))
    add_text(sl,
             f"Residual: {qo_sign} notch ({qo_label})  |  "
             f"SRM: {srm_rating}  Post-Suppl: {final_indicative}  Official: {r['Actual_Rating']}  |  "
             f"BBB Peers: {len(peer_countries_pptx)}",
             0.4, 0.82, 12.5, 0.32, size=9, bold=True,
             color="065f46" if qo > 0 else ("991b1b" if qo < 0 else "1E3A8A"))

    # Peer comparison mini-table as image
    if comp_rows_pptx:
        import matplotlib.pyplot as _plt
        import matplotlib.patches as _mp
        fig, ax = _plt.subplots(figsize=(13, 4.5))
        ax.axis("off")
        col_labels = ["Indicator", target, "BBB Avg", "Diff", "Signal"]
        cell_data  = [
            [r2["label"],
             f"{r2['t_val']:.1f}",
             f"{r2['p_avg']:.1f}",
             f"{r2['diff']:+.1f}",
             "Above" if r2["is_better"] else ("Below" if r2["is_worse"] else "In line")]
            for r2 in comp_rows_pptx
        ]
        tbl = ax.table(cellText=cell_data, colLabels=col_labels,
                       loc="center", cellLoc="center")
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(9)
        tbl.scale(1, 1.3)
        for ri, r2 in enumerate(comp_rows_pptx):
            for ci in [1, 3, 4]:
                cell = tbl[ri+1, ci]
                if r2["is_better"]:   cell.set_facecolor("#d1fae5")
                elif r2["is_worse"]:  cell.set_facecolor("#fee2e2")
        fig.tight_layout()
        tbuf = io.BytesIO()
        fig.savefig(tbuf, format="png", dpi=150, bbox_inches="tight")
        _plt.close(fig)
        tbuf.seek(0)
        add_image_buf(sl, tbuf, 0.3, 1.3, 12.7, 4.0)

    # ── SLIDE: SELECTED PEER QO COMPARISON ───────────────────────────────────
    if len(comp_list) > 1:
        sl = blank_slide()
        sl.background.fill.solid()
        sl.background.fill.fore_color.rgb = rgb("F8FAFC")
        add_rect(sl, 0, 0, 13.33, 0.7, "0D9488")
        add_text(sl, "4c. Selected Peer QO Comparison",
                 0.3, 0.1, 12, 0.5, size=13, bold=True, color="FFFFFF")

        # Build table as matplotlib image
        import matplotlib.pyplot as _plt
        col_labels = ["Country", "Actual", "SRM", "Residual QO", "IE", "FP"]
        cell_data  = [
            [c["Country"],
             str(c["_nr"]['Actual_Rating']),
             c["_srm"],
             f"{int(c['_qo']):+}",
             f"{c['_ie']:.2f}",
             f"{c['_fp']:.2f}"]
            for c in comp_list
        ]

        fig, ax = _plt.subplots(figsize=(13, max(2.5, len(comp_list) * 0.6 + 1)))
        ax.axis("off")
        tbl = ax.table(cellText=cell_data, colLabels=col_labels,
                       loc="center", cellLoc="center")
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(10)
        tbl.scale(1, 1.5)

        # Colour cells
        for ri, c in enumerate(comp_list):
            qo_val = int(c['_qo'])
            # QO column (index 3)
            tbl[ri+1, 3].set_facecolor(
                "#d1fae5" if qo_val > 0 else ("#fee2e2" if qo_val < 0 else "#F1F5F9"))
            # Actual rating (index 1)
            num = RATING_TO_NUM.get(str(c["_nr"]['Actual_Rating']).replace('*','').strip(), -1)
            tbl[ri+1, 1].set_facecolor(
                "#d1fae5" if num >= 10 else ("#fef3c7" if num >= 6 else "#fee2e2"))

        fig.tight_layout()
        tbuf = io.BytesIO()
        fig.savefig(tbuf, format="png", dpi=150, bbox_inches="tight")
        _plt.close(fig)
        tbuf.seek(0)
        add_image_buf(sl, tbuf, 0.5, 0.85, 12.3, 3.5)

        # QO delta narrative text
        target_qo    = int(comp_list[0]["_qo"]) if comp_list else 0
        peer_qos     = [(c["Country"], int(c["_qo"])) for c in comp_list[1:]]
        higher_peers = [f"{n}({q:+})" for n, q in peer_qos if q > target_qo]
        lower_peers  = [f"{n}({q:+})" for n, q in peer_qos if q < target_qo]

        y_txt = 4.5
        if higher_peers:
            add_rect(sl, 0.3, y_txt, 12.7, 0.5, "fef3c7",
                     line_hex="CBD5E1", line_w=Pt(0.5))
            add_text(sl,
                     f"Peers with higher QO: {', '.join(higher_peers)} — "
                     f"room to strengthen {target}'s narrative.",
                     0.4, y_txt + 0.08, 12.5, 0.35, size=9, color="78350f")
            y_txt += 0.58

        if lower_peers:
            add_rect(sl, 0.3, y_txt, 12.7, 0.5, "d1fae5",
                     line_hex="CBD5E1", line_w=Pt(0.5))
            add_text(sl,
                     f"Peers with lower QO: {', '.join(lower_peers)} — "
                     f"{target} has a comparative QO advantage over these peers.",
                     0.4, y_txt + 0.08, 12.5, 0.35, size=9, color="065f46")
            y_txt += 0.58

        if not higher_peers and not lower_peers:
            add_rect(sl, 0.3, y_txt, 12.7, 0.5, "EFF6FF",
                     line_hex="CBD5E1", line_w=Pt(0.5))
            add_text(sl,
                     f"All selected peers have the same residual QO as {target}.",
                     0.4, y_txt + 0.08, 12.5, 0.35, size=9, color="1E3A8A")


    # ── SLIDE: QO RANKED ARGUMENTS ────────────────────────────────────────────
    sl = blank_slide()
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb("F8FAFC")
    add_rect(sl, 0, 0, 13.33, 0.7, "0D9488")
    add_text(sl, "4b. QO Targeting — Ranked Arguments & Strategy",
             0.3, 0.1, 12, 0.5, size=13, bold=True, color="FFFFFF")

    y_o = 0.82
    if strengths_pptx:
        add_text(sl, f"Top QO Arguments — {target} vs BBB Peers:",
                 0.3, y_o, 12.5, 0.3, size=10, bold=True, color="1E3A8A")
        y_o += 0.35

        for i, r2 in enumerate(strengths_pptx[:4], 1):
            pct = r2["pct_diff"]
            if pct >= 30:   bg_c = "d1fae5"; fg_c = "065f46"; lbl = "Very Strong"
            elif pct >= 15: bg_c = "dbeafe"; fg_c = "1e3a8a"; lbl = "Strong"
            else:           bg_c = "ede9fe"; fg_c = "6366f1"; lbl = "Supporting"

            add_rect(sl, 0.2, y_o, 12.9, 0.88, bg_c, line_hex=fg_c, line_w=Pt(0.5))
            add_text(sl,
                     f"#{i} [{lbl}] {r2['factor']} | {r2['label']}  "
                     f"{target}: {r2['t_val']:.1f} vs BBB avg: {r2['p_avg']:.1f} "
                     f"({r2['diff']:+.1f}, {pct:.0f}% outperformance)",
                     0.35, y_o + 0.04, 12.5, 0.3, size=9, bold=True, color=fg_c)
            arg_short = r2["argument"][:200] + ("…" if len(r2["argument"]) > 200 else "")
            add_text(sl, arg_short, 0.35, y_o + 0.36, 12.5, 0.46, size=8, color="374151")
            y_o += 0.96
            if y_o > 6.2: break

        # Strategy box
        top_factors = list(dict.fromkeys([r2["factor"] for r2 in strengths_pptx[:3]]))
        strategy = (
            f"Lead with {strengths_pptx[0]['label']} "
            f"({strengths_pptx[0]['t_val']:.1f} vs avg {strengths_pptx[0]['p_avg']:.1f}, "
            f"{strengths_pptx[0]['pct_diff']:.0f}% outperformance). "
            f"Target QO factors: {', '.join(top_factors)}."
        )
        # Calculate strategy box height based on text length
        strat_h = 0.65 if len(strategy) > 120 else 0.5
        if y_o + strat_h < 6.3:
            add_rect(sl, 0.2, y_o, 12.9, strat_h, "1E3A8A")
            add_text(sl, f"🎯 Strategy: {strategy}",
                     0.35, y_o + 0.08, 12.4, strat_h - 0.1, size=9, color="FFFFFF")
            y_o += strat_h + 0.15   # ← advance y_o past the strategy box

    # Weaknesses note — only show if there's enough space remaining
    if weaknesses_pptx and y_o < 6.2:
        wk_lines = [
            f"{r2['label']}: {r2['t_val']:.1f} vs BBB avg {r2['p_avg']:.1f} ({r2['diff']:+.1f})"
            for r2 in weaknesses_pptx[:3]
        ]
        wk_content = "  |  ".join(wk_lines)
        box_h = 0.75
        add_rect(sl, 0.3, y_o, 12.7, box_h, "fff5f5",
                 line_hex="fca5a5", line_w=Pt(0.5))
        add_text(sl,
                 f"⚠️ Underperforming vs BBB peers ({len(weaknesses_pptx)} indicators):",
                 0.4, y_o + 0.05, 12.5, 0.28, size=8, bold=True, color="991b1b")
        add_text(sl, wk_content,
                 0.4, y_o + 0.33, 12.5, 0.35, size=7.5, color="7f1d1d")

    # Footer slide
    sl = blank_slide()
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb("1E3A8A")
    add_text(sl, "Sovereign Rating Monitoring",
             1, 2.5, 11.33, 1, size=28, bold=True, color="FFFFFF", align=PP_ALIGN.CENTER)
    add_text(sl, f"Briefing Note — {target}",
             1, 3.6, 11.33, 0.6, size=16, color="94A3B8", align=PP_ALIGN.CENTER)
    add_text(sl, f"Generated {date_str}  |  Based on S&P Global Methodology",
             1, 4.4, 11.33, 0.4, size=11, color="64748B", align=PP_ALIGN.CENTER)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out