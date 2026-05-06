"""
briefing_generator.py
Generates PDF and PPTX briefing notes from sovereign rating dashboard data.
Import this module and call generate_pdf() or generate_pptx().
"""

import io
import textwrap
from datetime import datetime

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

def generate_pdf(target, r, srm_rating, qo,
                 s_inst, s_eco, s_fis, s_ext, s_mon,
                 prof_ie, prof_fp,
                 comp_list, sel_nations, trend_df, selected_metrics,
                 signals, qo_opportunities,
                 # ── ADD THESE ──
                 supp_adj=0, supp_factors=None, final_indicative=None, cap_note=None):
    supp_factors    = supp_factors or []
    final_indicative = final_indicative or srm_rating
    supp_adj_sign   = f"{int(supp_adj):+}" if supp_adj != 0 else "None"

    srm_to_final_diff  = RATING_TO_NUM.get(final_indicative, 8) - RATING_TO_NUM.get(srm_rating, 8)
    final_to_actual_diff = RATING_TO_NUM.get(str(r['Actual_Rating']).replace('*','').strip(), 8) - RATING_TO_NUM.get(final_indicative, 8)

    """
    Returns a bytes buffer containing the PDF briefing note.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=2*cm, bottomMargin=2*cm,
        title=f"Sovereign Rating Briefing — {target}"
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
    story.append(Paragraph(f"Sovereign Credit Rating Briefing", st_h1))
    story.append(Paragraph(f"<b>{target}</b> | S&amp;P Methodology Analysis", st_h2))
    story.append(Paragraph(
        f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')} &nbsp;|&nbsp; "
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
    qo_sign       = f"{int(qo):+}"
    supp_adj_sign = f"{int(supp_adj):+}" if supp_adj != 0 else "None"
    kpi_data = [
        ["Actual Rating", "SRM Matrix", "After Supplemental", "Residual ±1", "IE / FP Profile"],
        [str(r['Actual_Rating']), srm_rating, final_indicative,
         qo_sign, f"{prof_ie:.2f} / {prof_fp:.2f}"],
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
    # Supplemental factors narrative
    if supp_factors:
        story.append(Paragraph("<b>⚡ Supplemental Adjustment Factors (para. 125–128):</b>", st_h3))
        for f in supp_factors:
            prefix = "▼" if f["type"] == "negative" else "▲"
            story.append(Paragraph(
                f"{prefix} <b>{f['factor']}</b> [{f['impact']}]: {f['detail']}",
                st_body))
    if cap_note:
        story.append(Paragraph(
            f"🚧 <b>Hard Cap Applied:</b> {cap_note}", st_body))
    if not supp_factors and not cap_note:
        story.append(Paragraph(
            "No supplemental adjustment factors triggered. "
            "SRM matrix output equals post-supplemental indicative rating.",
            st_body))

    story.append(vsp(4))

    # Residual ±1 notch
    qo_label = "Upgrade" if qo > 0 else ("Penalty" if qo < 0 else "Neutral")
    qo_text  = (
        "Positive residual: S&P views transitional factors, ESG considerations, or "
        "over-performance vs peers as supporting a higher rating than the model implies."
        if qo > 0 else
        "Negative residual: S&P identified factors not fully captured by the model — "
        "transitional risks, underperformance vs peers, or governance concerns."
        if qo < 0 else
        "No residual adjustment — the post-supplemental indicative rating aligns with "
        "the official rating."
    )
    story.append(Paragraph(
        f"<b>Residual ±1 Notch Adjustment ({qo_sign} — {qo_label}):</b> {qo_text}",
        st_body))
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

    # ── Executive Summary ─────────────────────────────────────────────────────
    strength_sigs  = [s for s in signals if s["type"] == "strength"]
    moderate_sigs  = [s for s in signals if s["type"] == "moderate"]
    weakness_sigs  = [s for s in signals if s["type"] == "weakness"]

    upgrade_count  = len(strength_sigs)
    moderate_count = len(moderate_sigs)
    weakness_count = len(weakness_sigs)

    top_strengths  = [s['metric'] for s in strength_sigs][:3]
    top_weaknesses = [s['metric'] for s in weakness_sigs][:3]

    strength_text  = (f"Key strengths include {', '.join(top_strengths)}."
                      if top_strengths else "No material strengths identified.")
    weakness_text  = (f"Key areas of concern are {', '.join(top_weaknesses)}."
                      if top_weaknesses else "No material weaknesses identified.")

    if weakness_count == 0 and upgrade_count >= 3:
        trajectory     = "Upgrade Candidate"
        traj_bg        = RL_GREEN
    elif weakness_count >= 3:
        trajectory     = "Downgrade Risk"
        traj_bg        = RL_RED
    elif weakness_count >= 1 and moderate_count >= 2:
        trajectory     = "Stable with Caution"
        traj_bg        = RL_YELLOW
    else:
        trajectory     = "Broadly Stable"
        traj_bg        = colors.HexColor("#EFF6FF")

    if qo > 0:
        qo_exec = (f"S&P's committee applied a +{int(qo)}-notch residual adjustment above "
                   f"the post-supplemental indicative rating of {final_indicative}, "
                   f"recognising positive factors beyond the quantitative model.")
    elif qo < 0:
        qo_exec = (f"S&P's committee applied a {int(qo)}-notch residual adjustment below "
                   f"the post-supplemental indicative of {final_indicative}, "
                   f"reflecting hidden risks not captured by the model.")
    else:
        qo_exec = (f"The official rating {r['Actual_Rating']} aligns with the "
                   f"post-supplemental indicative {final_indicative} — no residual adjustment applied.")

    if supp_factors:
        supp_exec = (f"Supplemental factors triggered: "
                     f"{', '.join([f['factor'] for f in supp_factors])} "
                     f"({int(supp_adj):+} notch from SRM output of {srm_rating}).")
    elif cap_note:
        supp_exec = f"Hard cap applied: {cap_note}."
    else:
        supp_exec = (f"No supplemental factors triggered — SRM output {srm_rating} "
                     f"carried through unchanged.")

    if trajectory == "Broadly Stable":
        closing = ("The overall profile supports the current rating with potential "
                   "for upgrade if key weaknesses are addressed.")
    elif trajectory == "Stable with Caution":
        closing = ("The profile warrants close monitoring — deterioration in watch "
                   "areas could trigger a negative outlook.")
    elif trajectory == "Upgrade Candidate":
        closing = ("The strong fundamental profile positions this sovereign for a "
                   "potential upgrade in the near to medium term.")
    else:
        closing = ("Sustained weakness across multiple pillars creates meaningful "
                   "downgrade risk if policy correction is not forthcoming.")

    summary_text = (
        f"{target} holds an official S&P rating of {r['Actual_Rating']}, derived from "
        f"an SRM matrix output of {srm_rating} (IE: {prof_ie:.2f}, FP: {prof_fp:.2f}). "
        f"{supp_exec} {qo_exec} "
        f"Across {len(signals)} indicator signals: {upgrade_count} strong, "
        f"{moderate_count} moderate, {weakness_count} requiring attention. "
        f"{strength_text} {weakness_text} {closing}"
    )

    # Trajectory label box
    traj_table = Table(
        [[Paragraph(f"Rating Trajectory: {trajectory}", ParagraphStyle(
            "traj", fontSize=10, fontName="Helvetica-Bold",
            textColor=colors.HexColor("#065f46" if trajectory=="Upgrade Candidate"
                       else ("#991b1b" if trajectory=="Downgrade Risk"
                       else ("#78350f" if trajectory=="Stable with Caution"
                       else "#1e3a8a")))))]],
        colWidths=[17.4*cm]
    )
    traj_table.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), traj_bg),
        ('TOPPADDING',    (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING',   (0,0), (-1,-1), 10),
        ('BOX',           (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
    ]))
    story.append(traj_table)
    story.append(vsp(4))
    story.append(Paragraph(summary_text, st_body))
    story.append(vsp(8))


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

    # QO opportunity table
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
            ('BACKGROUND', (0,0), (-1,0), RL_TEAL),
            ('TEXTCOLOR',  (0,0), (-1,0), RL_WHITE),
            ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor("#EFF6FF"), RL_WHITE]),
            ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ('INNERGRID',  (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
            ('VALIGN',     (0,0), (-1,-1), 'TOP'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ])
        story.append(Table(opp_rows,
            colWidths=[3.5*cm, 11*cm, 3*cm],
            style=opp_style, repeatRows=1))
        story.append(vsp(8))
    else:
        story.append(Paragraph(
            "No strong QO narrative opportunities identified based on current data. "
            "Focus on addressing weaknesses first to unlock positive QO arguments.",
            st_body))
        story.append(vsp(6))

    # Peer comparison for QO
    if len(comp_list) > 1:
        story.append(Paragraph("QO Peer Context", st_h3))
        peer_rows = [["Country", "Actual", "SRM", "QO", "IE Prof.", "FP Prof."]]
        for c in comp_list:
            peer_rows.append([
                c["Country"], c["_nr"]['Actual_Rating'], c["_srm"],
                f"{int(c['_qo']):+}",
                f"{c['_ie']:.2f}", f"{c['_fp']:.2f}",
            ])
        peer_style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), RL_NAVY),
            ('TEXTCOLOR',  (0,0), (-1,0), RL_WHITE),
            ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0), (-1,-1), 9),
            ('ALIGN',      (1,0), (-1,-1), 'CENTER'),
            ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ('INNERGRID',  (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [RL_LGREY, RL_WHITE]),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ])
        for ri, c in enumerate(comp_list, 1):
            qo_val = int(c['_qo'])
            bg = RL_GREEN if qo_val > 0 else (RL_RED if qo_val < 0 else RL_WHITE)
            peer_style.add('BACKGROUND', (3, ri), (3, ri), bg)
            for ci_idx, col_label in [(1,'Actual'),(2,'SRM')]:
                rating_val = c["_nr"]['Actual_Rating'] if ci_idx == 1 else c["_srm"]
                rbg, _ = _rating_color(rating_val)
                peer_style.add('BACKGROUND', (ci_idx, ri), (ci_idx, ri), rbg)

        story.append(Table(peer_rows,
            colWidths=[4.5*cm, 2.2*cm, 2.2*cm, 1.8*cm, 2.3*cm, 2.3*cm],
            style=peer_style))
        story.append(vsp(6))

        # QO delta narrative
        target_qo = int(comp_list[0]["_qo"]) if comp_list else 0
        peer_qos = [(c["Country"], int(c["_qo"])) for c in comp_list[1:]]
        higher_peers = [f"{n} ({q:+})" for n, q in peer_qos if q > target_qo]
        lower_peers  = [f"{n} ({q:+})" for n, q in peer_qos if q < target_qo]

        if higher_peers:
            story.append(Paragraph(
                f"<b>Peers with higher QO:</b> {', '.join(higher_peers)}. "
                "These peers receive more favourable qualitative adjustments, suggesting "
                f"{target} has room to strengthen its narrative on institutional quality, "
                "contingent liability management, or fiscal flexibility.",
                st_body))
        if lower_peers:
            story.append(Paragraph(
                f"<b>Peers with lower QO:</b> {', '.join(lower_peers)}. "
                f"{target}'s relative QO position is stronger, which should be highlighted "
                "in rating dialogue as a comparative advantage.",
                st_body))

    # Footer
    story.append(vsp(16))
    story.append(hr())
    # Rating derivation flow summary box
    story.append(vsp(8))
    story.append(Paragraph("<b>Rating Derivation — Three Layer Flow</b>", st_h3))

    # Wrap long text in Paragraph so ReportLab word-wraps within column
    st_cell = ParagraphStyle("cell", fontSize=8, leading=11, wordWrap='CJK')
    st_head = ParagraphStyle("head", fontSize=8, leading=11,
                             fontName="Helvetica-Bold", textColor=colors.white)

    def cell(text):
        return Paragraph(str(text), st_cell)

    def head(text):
        return Paragraph(str(text), st_head)

    # Build basis text for step 2
    if cap_note:
        basis_2 = cap_note
    elif supp_factors:
        basis_2 = "; ".join([f['factor'] for f in supp_factors])
    else:
        basis_2 = "Para. 125-128 conditions not met — no extreme external, fiscal, or institutional triggers."

    # Build basis text for step 3
    if qo != 0:
        basis_3 = ("Para. 15: positive transitional dynamics, ESG factors, or sustained "
                   "outperformance vs peers." if qo > 0 else
                   "Para. 15: negative transitional dynamics, ESG concerns, or sustained "
                   "underperformance vs peers.")
    else:
        basis_3 = "Post-supplemental indicative rating equals official rating — no committee adjustment."

    flow_rows = [
        [head("Step"), head("Description"), head("Output"), head("Basis")],
        [cell("1"),
         cell("Five-pillar SRM matrix lookup"),
         cell(srm_rating),
         cell(f"IE Profile = {prof_ie:.2f}, FP Profile = {prof_fp:.2f}. "
              f"Matrix intersection determines indicative rating.")],
        [cell("2"),
         cell(f"Supplemental adjustment: {int(supp_adj):+} notch" if supp_adj != 0
              else ("Hard cap applied" if cap_note
                    else "No supplemental factors triggered")),
         cell(final_indicative),
         cell(basis_2)],
        [cell("3"),
         cell(f"Residual ±1 notch adjustment: {int(qo):+}" if qo != 0
              else "No residual adjustment"),
         cell(str(r['Actual_Rating'])),
         cell(basis_3)],
    ]

    flow_style = TableStyle([
        ('BACKGROUND',    (0,0), (-1,0), RL_NAVY),
        ('ALIGN',         (2,0), (2,-1), 'CENTER'),
        ('VALIGN',        (0,0), (-1,-1), 'TOP'),
        ('BOX',           (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('INNERGRID',     (0,0), (-1,-1), 0.3, colors.HexColor("#E2E8F0")),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [RL_WHITE, RL_LGREY, RL_WHITE]),
        ('BACKGROUND',    (2,1), (2,1), _rating_color(srm_rating)[0]),
        ('BACKGROUND',    (2,2), (2,2), _rating_color(final_indicative)[0]),
        ('BACKGROUND',    (2,3), (2,3), _rating_color(r['Actual_Rating'])[0]),
        ('TOPPADDING',    (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING',   (0,0), (-1,-1), 4),
        ('RIGHTPADDING',  (0,0), (-1,-1), 4),
    ])

    story.append(Table(flow_rows,
        colWidths=[1*cm, 5.5*cm, 2*cm, 9*cm],   # wider Basis column
        style=flow_style,
        repeatRows=1))

    story.append(vsp(4))
    story.append(Paragraph(
        (f"<i>Note: A residual adjustment of {int(qo):+} notch means S&P's rating committee "
         f"applied a para. 15 adjustment {'above' if qo>0 else 'below'} the post-supplemental "
         f"indicative rating of <b>{final_indicative}</b> to arrive at the official "
         f"<b>{r['Actual_Rating']}</b>. This is NOT a supplemental adjustment — it reflects "
         f"{'positive transitional dynamics, favorable ESG factors, or sustained outperformance vs peers.' if qo>0 else 'negative transitional dynamics, ESG concerns, or sustained underperformance vs peers.'}</i>")
        if qo != 0 else
        (f"<i>The official rating <b>{r['Actual_Rating']}</b> equals the post-supplemental "
         f"indicative rating — no residual adjustment was applied by S&P's rating committee.</i>"),
        st_small))
    
    story.append(Paragraph(
        f"<i>Confidential — Sovereign Rating Monitoring Dashboard | "
        f"Generated {datetime.now().strftime('%d %B %Y')} | "
        f"Methodology: S&P Global Ratings Sovereign Rating Criteria (Dec 2017, updated Oct 2024) — "
        f"spglobal.com/ratings/en/regulatory/article/-/view/sourceId/10221157</i>",
        st_small))

    story.append(Paragraph(
        f"<i>Confidential — Sovereign Rating Monitoring Dashboard | "
        f"Generated {datetime.now().strftime('%d %B %Y')}</i>",
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
                  # ── ADD THESE ──
                  supp_adj=0, supp_factors=None, final_indicative=None, cap_note=None):
    supp_factors    = supp_factors or []
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
    add_text(sl,
             f"SRM: {srm_rating}  →  {final_indicative}  |  Residual: {qo_sign}",
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
        ("Actual Rating",     str(r['Actual_Rating']), rating_fill(r['Actual_Rating'])),
        ("SRM Matrix",        srm_rating,               rating_fill(srm_rating)),
        ("Post-Supplemental", final_indicative,         rating_fill(final_indicative)),
        ("Residual ±1 Notch", qo_sign,
         "d1fae5" if qo > 0 else ("fee2e2" if qo < 0 else "F1F5F9")),
        ("IE / FP",           f"{prof_ie:.2f} / {prof_fp:.2f}", "EFF6FF"),
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
    # Supplemental + residual combined note
    if supp_factors:
        supp_text = f"⚡ Supplemental: {'; '.join([f['factor']+' ('+f['impact']+')' for f in supp_factors])}"
    elif cap_note:
        supp_text = f"🚧 Cap: {cap_note}"
    else:
        supp_text = "✅ No supplemental adjustments triggered"

    residual_text = (
        f"Residual {qo_sign} notch ({qo_label}): " + (
        "Positive factors beyond model." if qo > 0 else
        "Hidden risks beyond model." if qo < 0 else
        "Model aligned with official rating.")
    )

    add_rect(sl, 0.3, 3.55, 12.7, 0.35,
             "d1fae5" if not supp_factors and not cap_note else "fef3c7",
             line_hex="CBD5E1", line_w=Pt(0.5))
    add_text(sl, supp_text, 0.4, 3.58, 12.5, 0.28, size=8,
             color="065f46" if not supp_factors else "78350f")

    add_rect(sl, 0.3, 3.95, 12.7, 0.35,
             "d1fae5" if qo > 0 else ("fee2e2" if qo < 0 else "F1F5F9"),
             line_hex="CBD5E1", line_w=Pt(0.5))
    add_text(sl, residual_text, 0.4, 3.98, 12.5, 0.28, size=8,
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

    # ── SLIDE: RECOMMENDATION EXECUTIVE SUMMARY ───────────────────────────────
    sl = blank_slide()
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb("F8FAFC")
    add_rect(sl, 0, 0, 13.33, 0.7, "1E3A8A")
    add_text(sl, "3. Recommendation — Executive Summary",
             0.3, 0.1, 12, 0.5, size=14, bold=True, color="FFFFFF")

    # Recompute counts for pptx
    strength_sigs_p  = [s for s in signals if s["type"] == "strength"]
    moderate_sigs_p  = [s for s in signals if s["type"] == "moderate"]
    weakness_sigs_p  = [s for s in signals if s["type"] == "weakness"]
    upgrade_count_p  = len(strength_sigs_p)
    moderate_count_p = len(moderate_sigs_p)
    weakness_count_p = len(weakness_sigs_p)

    top_str_p = [s['metric'] for s in strength_sigs_p][:3]
    top_wk_p  = [s['metric'] for s in weakness_sigs_p][:3]

    if weakness_count_p == 0 and upgrade_count_p >= 3:
        traj_p = "Upgrade Candidate"; traj_col_p = "065f46"; traj_bg_p = "d1fae5"
    elif weakness_count_p >= 3:
        traj_p = "Downgrade Risk";    traj_col_p = "991b1b"; traj_bg_p = "fee2e2"
    elif weakness_count_p >= 1 and moderate_count_p >= 2:
        traj_p = "Stable with Caution"; traj_col_p = "78350f"; traj_bg_p = "fef3c7"
    else:
        traj_p = "Broadly Stable";   traj_col_p = "1e3a8a"; traj_bg_p = "EFF6FF"

    # Trajectory banner
    add_rect(sl, 0.3, 0.85, 12.7, 0.5, traj_bg_p, line_hex="CBD5E1", line_w=Pt(0.5))
    add_text(sl, f"Rating Trajectory: {traj_p}",
             0.4, 0.9, 12.5, 0.38, size=13, bold=True, color=traj_col_p)

    # Rating derivation row
    add_rect(sl, 0.3, 1.45, 12.7, 0.5, "F1F5F9", line_hex="CBD5E1", line_w=Pt(0.5))
    add_text(sl,
             f"SRM: {srm_rating}  →  Post-Supplemental: {final_indicative}  "
             f"→  Official: {r['Actual_Rating']}  |  "
             f"IE: {prof_ie:.2f}  FP: {prof_fp:.2f}  |  Residual: {int(qo):+} notch",
             0.4, 1.5, 12.5, 0.38, size=10, color="1e3a8a", bold=True)

    # Supplemental note
    if supp_factors:
        supp_p = (f"⚡ Supplemental: "
                  f"{', '.join([f['factor'] for f in supp_factors])} "
                  f"({int(supp_adj):+} notch)")
    elif cap_note:
        supp_p = f"🚧 Cap: {cap_note}"
    else:
        supp_p = f"✅ No supplemental factors — SRM {srm_rating} carried through unchanged"

    add_rect(sl, 0.3, 2.05, 12.7, 0.4,
             "fee2e2" if supp_factors else "d1fae5",
             line_hex="CBD5E1", line_w=Pt(0.5))
    add_text(sl, supp_p, 0.4, 2.1, 12.5, 0.3, size=9,
             color="991b1b" if supp_factors else "065f46")

    # Signal count cards
    for i, (label, count, bg, fg) in enumerate([
        (f"✅ Strong Signals",   upgrade_count_p,  "d1fae5", "065f46"),
        (f"⚠️ Watch Areas",      moderate_count_p, "fef3c7", "78350f"),
        (f"❌ Weak Signals",     weakness_count_p, "fee2e2", "991b1b"),
    ]):
        x = 0.3 + i * 4.3
        add_rect(sl, x, 2.6, 4.0, 0.9, bg, line_hex="CBD5E1", line_w=Pt(0.5))
        add_text(sl, str(count), x, 2.65, 4.0, 0.5,
                 size=28, bold=True, color=fg, align=PP_ALIGN.CENTER)
        add_text(sl, label, x, 3.15, 4.0, 0.28,
                 size=9, color=fg, align=PP_ALIGN.CENTER)

    # Key strengths
    add_text(sl, "Key Strengths:", 0.3, 3.65, 6.2, 0.3, size=10, bold=True, color="065f46")
    y_s = 3.98
    for s in strength_sigs_p[:4]:
        add_rect(sl, 0.3, y_s, 6.2, 0.45, "f0fdf4", line_hex="6ee7b7", line_w=Pt(0.3))
        add_text(sl, f"[{s['pillar']}] {s['metric']} — {s['value']}",
                 0.4, y_s + 0.05, 6.0, 0.35, size=8, color="065f46")
        y_s += 0.5

    # Key weaknesses
    add_text(sl, "Key Weaknesses:", 6.9, 3.65, 6.2, 0.3, size=10, bold=True, color="991b1b")
    y_w = 3.98
    for s in weakness_sigs_p[:4]:
        add_rect(sl, 6.9, y_w, 6.2, 0.45, "fff5f5", line_hex="fca5a5", line_w=Pt(0.3))
        add_text(sl, f"[{s['pillar']}] {s['metric']} — {s['value']}",
                 7.0, y_w + 0.05, 6.0, 0.35, size=8, color="991b1b")
        y_w += 0.5

    # ── SLIDE: QO IN-DEPTH ────────────────────────────────────────────────────
    sl = blank_slide()
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb("F8FAFC")
    add_rect(sl, 0, 0, 13.33, 0.7, "0D9488")
    add_text(sl, "4. Qualitative Overlay (QO) — In-Depth Targeting",
             0.3, 0.1, 12, 0.5, size=14, bold=True, color="FFFFFF")

    add_rect(sl, 0.2, 0.8, 12.9, 0.55,
             "d1fae5" if qo > 0 else ("fee2e2" if qo < 0 else "EFF6FF"),
             line_hex="CBD5E1", line_w=Pt(0.5))
    add_text(sl, f"Net QO: {qo_sign} notch ({qo_label})  |  "
                 f"SRM: {srm_rating}  →  Actual: {r['Actual_Rating']}",
             0.4, 0.85, 12.5, 0.4, size=12, bold=True,
             color="065f46" if qo > 0 else ("991b1b" if qo < 0 else "1E3A8A"))

    # Opportunities
    y_o = 1.5
    if qo_opportunities:
        add_text(sl, "Narrative Opportunities to Earn QO Points:",
                 0.3, y_o, 12.5, 0.3, size=10, bold=True, color="1E3A8A")
        y_o += 0.35
        for opp in qo_opportunities[:4]:
            impact_col = "065f46" if "+" in opp["impact"] else "6366f1"
            add_rect(sl, 0.2, y_o, 12.9, 0.75, "EFF6FF",
                     line_hex="BFDBFE", line_w=Pt(0.5))
            add_text(sl, f"⚖ {opp['factor']}  [{opp['impact']}]",
                     0.35, y_o + 0.04, 9, 0.28, size=9, bold=True, color=impact_col)
            arg_short = opp["argument"][:160] + ("…" if len(opp["argument"]) > 160 else "")
            add_text(sl, arg_short, 0.35, y_o + 0.3, 12.5, 0.38, size=8, color="374151")
            y_o += 0.82

    # Peer QO comparison
    if len(comp_list) > 1:
        add_text(sl, "QO Peer Comparison:",
                 0.3, y_o, 12, 0.28, size=10, bold=True, color="1E3A8A")
        y_o += 0.3
        for c in comp_list:
            cqo = int(c['_qo'])
            col = "065f46" if cqo > 0 else ("991b1b" if cqo < 0 else "475569")
            bg2 = "d1fae5" if cqo > 0 else ("fee2e2" if cqo < 0 else "F1F5F9")
            add_rect(sl, 0.2, y_o, 12.9, 0.38, bg2,
                     line_hex="CBD5E1", line_w=Pt(0.3))
            add_text(sl,
                     f"{c['Country']}  |  Actual: {c['_nr']['Actual_Rating']}  "
                     f"SRM: {c['_srm']}  QO: {cqo:+}  "
                     f"IE: {c['_ie']:.2f}  FP: {c['_fp']:.2f}",
                     0.35, y_o + 0.05, 12.5, 0.28, size=9, color=col)
            y_o += 0.42
            if y_o > 7.0: break

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
    add_text(sl,
             "Methodology: S&P Global Ratings — Sovereign Rating Methodology (Dec 2017, updated Oct 2024)",
             1, 5.0, 11.33, 0.3, size=9, color="64748B", align=PP_ALIGN.CENTER)
    add_text(sl,
             "spglobal.com/ratings/en/regulatory/article/-/view/sourceId/10221157",
             1, 5.3, 11.33, 0.3, size=9, color="1E3A8A", align=PP_ALIGN.CENTER)
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out