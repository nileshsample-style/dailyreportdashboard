import os
import json
import time
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from google import genai
from google.genai import types

st.set_page_config(
    page_title="Executive Dashboard Suite | AI PPT Generator",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Executive Dashboard Suite")
st.caption("Upload your metrics screenshots to generate pixel-accurate, C-Suite ready PowerPoint presentations with native charts, tables, badges, and stickers.")

# API Key handling
default_key = ""
if "GEMINI_API_KEY" in st.secrets:
    default_key = st.secrets["GEMINI_API_KEY"]
elif os.getenv("GEMINI_API_KEY"):
    default_key = os.getenv("GEMINI_API_KEY")

api_key = st.sidebar.text_input("Gemini API Key", type="password", value=default_key, help="Get key from [https://aistudio.google.com](https://aistudio.google.com)")

# ----------------------------------------------------------------------
# AI Extraction Engine
# ----------------------------------------------------------------------
PROMPT = """
You are an expert executive reporting analyst. Examine this dashboard screenshot and extract all data into strict JSON.
First determine which template family this belongs to:
1. "LEADS_OVERVIEW" (Used by Fellowship in Digital Health or HBF Leads: has dark indigo header banner with logo, Lead Status table, Qualification Breakdown table, Donut chart, Bar chart, Key Insights, and footer pills)
2. "ASSIMILATE" (Used by Assimilate.One: has header banner with icons, and vertical metric rows)
3. "YOUTUBE" (Used by YouTube Metrics: has date banner, subscribers added row, 4 KPI cards, spends banner at bottom)
4. "MEDVARSITY" (Used by Medvarsity Performance: has dark blue banner, 6 top KPI cards, 3 columns: Marketing Spend, Leads Assigned, Revenue Achieved, and 4 colored bottom insights)

Return JSON matching this exact structure:

If template_type is "LEADS_OVERVIEW":
{
    "template_type": "LEADS_OVERVIEW",
    "program_name": "FELLOWSHIP IN DIGITAL HEALTH / HBF LEADS OVERVIEW",
    "subtitle": "LEADS OVERVIEW",
    "total_leads": "43",
    "total_spends": "₹6,693",
    "cost_per_lead": "₹155.65",
    "report_date": "25th August 2026",
    "note": "Leads and Qualification data as of 25th August 2026.",
    "lead_status_summary": [
        {"status": "Lead", "count": 19, "share": "44.2%"},
        {"status": "Need Follow Up", "count": 3, "share": "7.0%"},
        {"status": "Fresh - Not Answered", "count": 21, "share": "48.8%"}
    ],
    "qualification_breakdown": [
        {"qualification": "Others", "count": 22, "share": "51.2%"},
        {"qualification": "BAMS/BHMS/...", "count": 15, "share": "34.9%"},
        {"qualification": "MBBS/MD/MS/DNB", "count": 5, "share": "11.6%"},
        {"qualification": "MBBS", "count": 1, "share": "2.3%"}
    ],
    "key_insights": [
        "Insight 1 text",
        "Insight 2 text",
        "Insight 3 text",
        "Insight 4 text"
    ]
}

If template_type is "ASSIMILATE":
{
    "template_type": "ASSIMILATE",
    "title": "Assimilate.One",
    "subtitle": "Assimilate Traffic",
    "metrics": [
        {"label": "Assimilate Traffic Total Spends", "value": "₹0"},
        {"label": "Total Users", "value": "251"},
        {"label": "New Users", "value": "191"},
        {"label": "Total No Of Assimilate Sessions Till 2nd'Sep'2026", "value": "0"},
        {"label": "Total No Of KOLs Onboarded Till 2nd'Sep'2026", "value": "0"},
        {"label": "Total No Of KOLs Onboarded Dates Confirmed", "value": "0"}
    ]
}

If template_type is "YOUTUBE":
{
    "template_type": "YOUTUBE",
    "date_title": "August'2026 31st",
    "gain_loss_text": "We gained 36,685 & lost 5,184",
    "subscribers_added": {
        "we_gained": "36,685",
        "we_lost": "5,184",
        "total_subscribers": "31,501"
    },
    "kpi_cards": [
        {"label": "Total Video Uploads", "value": "9"},
        {"label": "Total Views", "value": "1,94,082"},
        {"label": "Total Watch Hours", "value": "1,167"},
        {"label": "Overall Subscribers", "value": "5,69,187"}
    ],
    "spends": "63,355"
}

If template_type is "MEDVARSITY":
{
    "template_type": "MEDVARSITY",
    "title": "MEDVARSITY PERFORMANCE DASHBOARD",
    "period": "Data Till 11th September 2026",
    "top_kpis": [
        {"label": "Total Visits", "value": "13,862"},
        {"label": "Total Leads", "value": "805"},
        {"label": "New Leads", "value": "672"},
        {"label": "Reactivated", "value": "133"},
        {"label": "Total Deals", "value": "72"},
        {"label": "Total Revenue", "value": "58.29 L"}
    ],
    "marketing_spends": [
        {"label": "TOTAL MARKETING SPENDS (Till 10th Sep 2026)", "value": "2.5 L"},
        {"label": "Total Spends on Only Lead Gen (Google, Meta & WA)", "value": "2.51 L"},
        {"label": "Total Spent on Google", "value": "0"},
        {"label": "Total Spent on Meta", "value": "2.25 L"},
        {"label": "Total Spent on Whatsapp", "value": "26 K"},
        {"label": "Other Marketing Spends (YT, IG)", "value": "0"},
        {"label": "SPEND TO REVENUE %", "value": "4.32%"}
    ],
    "sales_assigned": [
        {"manager": "Bala", "new": "258", "reactivated": "38", "total": "296"},
        {"manager": "Muni", "new": "116", "reactivated": "16", "total": "132"},
        {"manager": "Dilip", "new": "79", "reactivated": "10", "total": "89"},
        {"manager": "Rahul", "new": "148", "reactivated": "52", "total": "200"},
        {"manager": "HBF", "new": "49", "reactivated": "10", "total": "49"},
        {"manager": "Post MBBS Apollo(Shwetha)", "new": "9", "reactivated": "6", "total": "15"},
        {"manager": "Sales Head", "new": "13", "reactivated": "1", "total": "14"},
        {"manager": "TOTAL", "new": "672", "reactivated": "133", "total": "805"}
    ],
    "revenue_team": [
        {"member": "Bala", "revenue": "27.83 L"},
        {"member": "Muni", "revenue": "13.51 L"},
        {"member": "Dilip", "revenue": "9.13 L"},
        {"member": "Rahul", "revenue": "7.81 L"}
    ],
    "insights": [
        {"title": "EXECUTIVE TAKEAWAY", "desc": "Strong monetisation with clear scope to improve lead-to-deal productivity.", "color": "dark_blue"},
        {"title": "REVENUE MOMENTUM IS STRONG", "desc": "₹58.29 L revenue delivered from ₹2.50 L marketing spend, keeping spend at only 4.32% of revenue.", "color": "green"},
        {"title": "CONVERSION IS THE KEY UPSIDE", "desc": "72 deals from 805 assigned leads gives a 8.9% lead-to-deal conversion.", "color": "purple"},
        {"title": "REVENUE & SALES PERFORMANCE NEED FOCUS", "desc": "Bala contributes ₹27.83 L (47.7% of revenue), while Bala + Muni together contribute 70.9%.", "color": "pink"}
    ]
}
Output raw JSON only.
"""

def extract_dashboard_data(img, key):
    client = genai.Client(api_key=key)
    candidate_models = ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-3.1-pro", "gemini-2.5-flash"]
    last_err = None
    for model_name in candidate_models:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=[PROMPT, img],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                text = response.text.strip()
                if text.startswith("```json"):
                    text = text[7:]
                elif text.startswith("```"):
                    text = text[3:]
                if text.endswith("```"):
                    text = text[:-3]
                return json.loads(text.strip())
            except Exception as e:
                last_err = e
                time.sleep(1.5)
                continue
    raise last_err

# ----------------------------------------------------------------------
# Visual Elements & Sticker Builders
# ----------------------------------------------------------------------
def draw_logo_badge(slide, left, top, badge_type="edu"):
    bg_circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, Inches(0.85), Inches(0.85))
    bg_circle.fill.solid()
    bg_circle.line.fill.background()
    
    if badge_type == "edu":
        bg_circle.fill.fore_color.rgb = RGBColor(255, 255, 255)
        cap_top = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, left + Inches(0.18), top + Inches(0.2), Inches(0.49), Inches(0.28))
        cap_top.fill.solid()
        cap_top.fill.fore_color.rgb = RGBColor(94, 23, 235)
        cap_top.line.fill.background()
        
        cap_base = slide.shapes.add_shape(MSO_SHAPE.CAN, left + Inches(0.28), top + Inches(0.42), Inches(0.29), Inches(0.2))
        cap_base.fill.solid()
        cap_base.fill.fore_color.rgb = RGBColor(94, 23, 235)
        cap_base.line.fill.background()
    elif badge_type == "hbf":
        bg_circle.fill.fore_color.rgb = RGBColor(255, 255, 255)
        tf = bg_circle.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = "HBF"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = RGBColor(27, 10, 89)
        p.alignment = PP_ALIGN.CENTER
    elif badge_type == "yt":
        bg_rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(0.95), Inches(0.68))
        bg_rect.fill.solid()
        bg_rect.fill.fore_color.rgb = RGBColor(255, 0, 0)
        bg_rect.line.fill.background()
        play_triangle = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, left + Inches(0.36), top + Inches(0.18), Inches(0.28), Inches(0.32))
        play_triangle.rotation = 90
        play_triangle.fill.solid()
        play_triangle.fill.fore_color.rgb = RGBColor(255, 255, 255)
        play_triangle.line.fill.background()

def draw_money_bag_sticker(slide, left, top):
    sack = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top + Inches(0.12), Inches(0.72), Inches(0.75))
    sack.fill.solid()
    sack.fill.fore_color.rgb = RGBColor(16, 185, 129)
    sack.line.fill.background()
    
    neck = slide.shapes.add_shape(MSO_SHAPE.TRAPEZOID, left + Inches(0.18), top, Inches(0.36), Inches(0.18))
    neck.rotation = 180
    neck.fill.solid()
    neck.fill.fore_color.rgb = RGBColor(5, 150, 105)
    neck.line.fill.background()
    
    coin = slide.shapes.add_shape(MSO_SHAPE.OVAL, left - Inches(0.2), top + Inches(0.35), Inches(0.35), Inches(0.45))
    coin.fill.solid()
    coin.fill.fore_color.rgb = RGBColor(245, 158, 11)
    coin.line.color.rgb = RGBColor(251, 191, 36)
    p = coin.text_frame.paragraphs[0]
    p.text = "₹"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER

# ----------------------------------------------------------------------
# Template 1 & 2: Leads Overview (Fellowship & HBF)
# ----------------------------------------------------------------------
def render_leads_overview(slide, data):
    BG_CANVAS = RGBColor(245, 247, 252)
    CARD_BG = RGBColor(255, 255, 255)
    CARD_BORDER = RGBColor(230, 235, 245)
    HEADER_BANNER = RGBColor(27, 10, 89)
    PURPLE_ACCENT = RGBColor(94, 23, 235)
    PINK_ACCENT = RGBColor(236, 32, 110)
    TEXT_DARK = RGBColor(30, 27, 75)
    TEXT_MUTED = RGBColor(100, 116, 139)

    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_CANVAS
    bg.line.fill.background()

    banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.35), Inches(12.333), Inches(1.15))
    banner.fill.solid()
    banner.fill.fore_color.rgb = HEADER_BANNER
    banner.line.fill.background()

    is_hbf = "HBF" in str(data.get("program_name", "")).upper()
    draw_logo_badge(slide, Inches(0.7), Inches(0.5), badge_type="hbf" if is_hbf else "edu")

    tb = slide.shapes.add_textbox(Inches(1.75), Inches(0.42), Inches(6.3), Inches(1.0))
    tf = tb.text_frame
    p1 = tf.paragraphs[0]
    p1.text = str(data.get("program_name", "PROGRAM OVERVIEW")).upper()
    p1.font.size = Pt(19)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(255, 255, 255)

    p2 = tf.add_paragraph()
    p2.text = str(data.get("subtitle", "LEADS OVERVIEW")).upper()
    p2.font.size = Pt(11)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(196, 181, 253)

    kpi1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.2), Inches(0.48), Inches(2.1), Inches(0.88))
    kpi1.fill.solid()
    kpi1.fill.fore_color.rgb = RGBColor(255, 255, 255)
    kpi1.line.fill.background()
    ktf1 = kpi1.text_frame
    ktf1.margin_top = Inches(0.08)
    p = ktf1.paragraphs[0]
    p.text = "👥 Total Leads"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = TEXT_MUTED
    p.alignment = PP_ALIGN.CENTER
    p = ktf1.add_paragraph()
    p.text = str(data.get("total_leads", "0"))
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = PURPLE_ACCENT
    p.alignment = PP_ALIGN.CENTER

    kpi2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.5), Inches(0.48), Inches(2.1), Inches(0.88))
    kpi2.fill.solid()
    kpi2.fill.fore_color.rgb = RGBColor(255, 255, 255)
    kpi2.line.fill.background()
    ktf2 = kpi2.text_frame
    ktf2.margin_top = Inches(0.08)
    p = ktf2.paragraphs[0]
    p.text = "💰 Total Spends"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = TEXT_MUTED
    p.alignment = PP_ALIGN.CENTER
    p = ktf2.add_paragraph()
    p.text = str(data.get("total_spends", "₹0"))
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = PINK_ACCENT
    p.alignment = PP_ALIGN.CENTER

    card_y = Inches(1.65)
    card_h = Inches(2.3)
    card_w = Inches(6.06)

    c_l = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), card_y, card_w, card_h)
    c_l.fill.solid()
    c_l.fill.fore_color.rgb = CARD_BG
    c_l.line.color.rgb = CARD_BORDER

    t_lbl1 = slide.shapes.add_textbox(Inches(0.6), card_y + Inches(0.06), Inches(5.8), Inches(0.3))
    p = t_lbl1.text_frame.paragraphs[0]
    p.text = "LEAD STATUS SUMMARY"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = PURPLE_ACCENT
    p.alignment = PP_ALIGN.CENTER

    l_items = data.get("lead_status_summary", [])
    t1 = slide.shapes.add_table(len(l_items) + 2, 3, Inches(0.65), card_y + Inches(0.36), Inches(5.76), Inches(1.8)).table
    for i, h in enumerate(["Lead Status", "Count", "% of Total"]):
        c = t1.cell(0, i)
        c.text = h
        c.fill.solid()
        c.fill.fore_color.rgb = PURPLE_ACCENT
        p = c.text_frame.paragraphs[0]
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

    tot1 = 0
    for r, item in enumerate(l_items):
        t1.cell(r + 1, 0).text = f"👤  {str(item.get('status', ''))}"
        cnt = item.get("count", 0)
        tot1 += int(cnt) if str(cnt).isdigit() else 0
        t1.cell(r + 1, 1).text = str(cnt)
        t1.cell(r + 1, 2).text = str(item.get("share", ""))
        for col_i in range(3):
            cell = t1.cell(r + 1, col_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(8)
            p.font.color.rgb = TEXT_DARK
            if col_i > 0: p.alignment = PP_ALIGN.CENTER

    tr0, tr1, tr2 = t1.cell(len(l_items) + 1, 0), t1.cell(len(l_items) + 1, 1), t1.cell(len(l_items) + 1, 2)
    tr0.text, tr1.text, tr2.text = "TOTAL", str(data.get("total_leads", tot1)), "100%"
    for col_i, c in enumerate([tr0, tr1, tr2]):
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor(245, 243, 255)
        p = c.text_frame.paragraphs[0]
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = PURPLE_ACCENT
        if col_i > 0: p.alignment = PP_ALIGN.CENTER

    c_r = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.77), card_y, card_w, card_h)
    c_r.fill.solid()
    c_r.fill.fore_color.rgb = CARD_BG
    c_r.line.color.rgb = CARD_BORDER

    t_lbl2 = slide.shapes.add_textbox(Inches(6.87), card_y + Inches(0.06), Inches(5.8), Inches(0.3))
    p = t_lbl2.text_frame.paragraphs[0]
    p.text = "QUALIFICATION BREAKDOWN"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = PINK_ACCENT
    p.alignment = PP_ALIGN.CENTER

    q_items = data.get("qualification_breakdown", [])
    t2 = slide.shapes.add_table(len(q_items) + 2, 3, Inches(6.92), card_y + Inches(0.36), Inches(5.76), Inches(1.8)).table
    for i, h in enumerate(["Qualification", "Count", "% of Total"]):
        c = t2.cell(0, i)
        c.text = h
        c.fill.solid()
        c.fill.fore_color.rgb = PINK_ACCENT
        p = c.text_frame.paragraphs[0]
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

    tot2 = 0
    for r, item in enumerate(q_items):
        t2.cell(r + 1, 0).text = f"🎓  {str(item.get('qualification', ''))}"
        cnt = item.get("count", 0)
        tot2 += int(cnt) if str(cnt).isdigit() else 0
        t2.cell(r + 1, 1).text = str(cnt)
        t2.cell(r + 1, 2).text = str(item.get("share", ""))
        for col_i in range(3):
            cell = t2.cell(r + 1, col_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(8)
            p.font.color.rgb = TEXT_DARK
            if col_i > 0: p.alignment = PP_ALIGN.CENTER

    tqr0, tqr1, tqr2 = t2.cell(len(q_items) + 1, 0), t2.cell(len(q_items) + 1, 1), t2.cell(len(q_items) + 1, 2)
    tqr0.text, tqr1.text, tqr2.text = "TOTAL", str(data.get("total_leads", tot2)), "100%"
    for col_i, c in enumerate([tqr0, tqr1, tqr2]):
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor(253, 242, 248)
        p = c.text_frame.paragraphs[0]
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = PINK_ACCENT
        if col_i > 0: p.alignment = PP_ALIGN.CENTER

    mid_y = Inches(4.05)
    mid_h = Inches(2.55)

    c_ins = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), mid_y, Inches(3.7), mid_h)
    c_ins.fill.solid()
    c_ins.fill.fore_color.rgb = CARD_BG
    c_ins.line.color.rgb = CARD_BORDER

    itb = slide.shapes.add_textbox(Inches(0.65), mid_y + Inches(0.08), Inches(3.4), Inches(0.3))
    ip = itb.text_frame.paragraphs[0]
    ip.text = "💡 KEY INSIGHTS"
    ip.font.size = Pt(10)
    ip.font.bold = True
    ip.font.color.rgb = PURPLE_ACCENT

    ins_tb = slide.shapes.add_textbox(Inches(0.65), mid_y + Inches(0.38), Inches(3.4), Inches(2.1))
    itf = ins_tb.text_frame
    itf.word_wrap = True
    for idx, insight in enumerate(data.get("key_insights", [])):
        p = itf.paragraphs[0] if idx == 0 else itf.add_paragraph()
        p.text = f"✔  {insight}"
        p.font.size = Pt(8)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(4)

    d_x = Inches(4.35)
    d_w = Inches(4.15)
    c_donut = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, d_x, mid_y, d_w, mid_h)
    c_donut.fill.solid()
    c_donut.fill.fore_color.rgb = CARD_BG
    c_donut.line.color.rgb = CARD_BORDER

    dtb = slide.shapes.add_textbox(d_x + Inches(0.1), mid_y + Inches(0.08), d_w - Inches(0.2), Inches(0.3))
    dp = dtb.text_frame.paragraphs[0]
    dp.text = "LEAD STATUS DISTRIBUTION"
    dp.font.size = Pt(10)
    dp.font.bold = True
    dp.font.color.rgb = PURPLE_ACCENT
    dp.alignment = PP_ALIGN.CENTER

    donut_data = CategoryChartData()
    donut_data.categories = [item.get("status", "") for item in l_items]
    donut_data.add_series("Leads", [float(str(item.get("share", "0")).replace("%", "")) for item in l_items])
    ch_d = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, d_x + Inches(0.15), mid_y + Inches(0.4), d_w - Inches(0.3), Inches(2.05), donut_data).chart
    ch_d.has_legend = True
    ch_d.legend.position = XL_LEGEND_POSITION.RIGHT
    ch_d.legend.font.size = Pt(7.5)

    b_x = Inches(8.65)
    b_w = Inches(4.15)
    c_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, b_x, mid_y, b_w, mid_h)
    c_bar.fill.solid()
    c_bar.fill.fore_color.rgb = CARD_BG
    c_bar.line.color.rgb = CARD_BORDER

    btb = slide.shapes.add_textbox(b_x + Inches(0.1), mid_y + Inches(0.08), b_w - Inches(0.2), Inches(0.3))
    bp = btb.text_frame.paragraphs[0]
    bp.text = "QUALIFICATION DISTRIBUTION"
    bp.font.size = Pt(10)
    bp.font.bold = True
    bp.font.color.rgb = PINK_ACCENT
    bp.alignment = PP_ALIGN.CENTER

    bar_data = CategoryChartData()
    bar_data.categories = [item.get("qualification", "")[:10] for item in q_items]
    bar_data.add_series("Share %", [float(str(item.get("share", "0")).replace("%", "")) for item in q_items])
    ch_b = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, b_x + Inches(0.15), mid_y + Inches(0.4), b_w - Inches(0.3), Inches(2.05), bar_data).chart
    ch_b.has_legend = False

    foot_y = Inches(6.72)
    foot_h = Inches(0.52)
    p_d = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), foot_y, Inches(3.0), foot_h)
    p_d.fill.solid()
    p_d.fill.fore_color.rgb = CARD_BG
    p_d.line.color.rgb = CARD_BORDER
    p = p_d.text_frame.paragraphs[0]
    p.text = f"📅 Report Date: {data.get('report_date', 'N/A')}"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK

    p_n = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.65), foot_y, Inches(6.0), foot_h)
    p_n.fill.solid()
    p_n.fill.fore_color.rgb = CARD_BG
    p_n.line.color.rgb = CARD_BORDER
    p = p_n.text_frame.paragraphs[0]
    p.text = f"⭐ Note: {data.get('note', '')}"
    p.font.size = Pt(8.5)
    p.font.color.rgb = TEXT_MUTED

    p_c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), foot_y, Inches(3.0), foot_h)
    p_c.fill.solid()
    p_c.fill.fore_color.rgb = CARD_BG
    p_c.line.color.rgb = CARD_BORDER
    p = p_c.text_frame.paragraphs[0]
    p.text = f"₹ Cost Per Lead: {data.get('cost_per_lead', 'N/A')}"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = PURPLE_ACCENT

# ----------------------------------------------------------------------
# Template 3: Assimilate.One
# ----------------------------------------------------------------------
def render_assimilate(slide, data):
    BG_CANVAS = RGBColor(255, 255, 255)
    HEADER_BANNER = RGBColor(90, 10, 50)
    PILL_BG = RGBColor(253, 242, 248)
    VALUE_COLOR = RGBColor(190, 24, 93)

    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_CANVAS
    bg.line.fill.background()

    banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(0.4), Inches(10.333), Inches(1.15))
    banner.fill.solid()
    banner.fill.fore_color.rgb = HEADER_BANNER
    banner.line.fill.background()

    stk = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.7), Inches(0.55), Inches(0.85), Inches(0.85))
    stk.fill.solid()
    stk.fill.fore_color.rgb = RGBColor(255, 255, 255)
    stk.line.fill.background()
    p = stk.text_frame.paragraphs[0]
    p.text = "📈"
    p.font.size = Pt(20)
    p.alignment = PP_ALIGN.CENTER

    tb = slide.shapes.add_textbox(Inches(3.0), Inches(0.45), Inches(7.333), Inches(1.0))
    p = tb.text_frame.paragraphs[0]
    p.text = str(data.get("title", "Assimilate.One"))
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    p2 = tb.text_frame.add_paragraph()
    p2.text = str(data.get("subtitle", "Assimilate Traffic"))
    p2.font.size = Pt(12)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(254, 205, 211)
    p2.alignment = PP_ALIGN.CENTER

    metrics = data.get("metrics", [])
    start_y = Inches(1.8)
    row_h = Inches(0.78)
    gap = Inches(0.12)
    icon_list = ["₹", "👥", "👤+", "📅", "👤✔", "🗓"]
    for i, m in enumerate(metrics[:6]):
        y = start_y + i * (row_h + gap)
        row_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), y, Inches(10.333), row_h)
        row_shape.fill.solid()
        row_shape.fill.fore_color.rgb = PILL_BG
        row_shape.line.color.rgb = RGBColor(251, 207, 232)

        ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.7), y + Inches(0.14), Inches(0.5), Inches(0.5))
        ic.fill.solid()
        ic.fill.fore_color.rgb = RGBColor(244, 114, 182)
        ic.line.fill.background()
        p = ic.text_frame.paragraphs[0]
        p.text = icon_list[i % len(icon_list)]
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        tb_l = slide.shapes.add_textbox(Inches(2.35), y + Inches(0.15), Inches(6.8), Inches(0.5))
        p = tb_l.text_frame.paragraphs[0]
        p.text = str(m.get("label", ""))
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = RGBColor(30, 27, 75)

        tb_r = slide.shapes.add_textbox(Inches(9.2), y + Inches(0.1), Inches(2.2), Inches(0.5))
        p = tb_r.text_frame.paragraphs[0]
        p.text = str(m.get("value", ""))
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = VALUE_COLOR
        p.alignment = PP_ALIGN.RIGHT

# ----------------------------------------------------------------------
# Template 4: YouTube Metrics
# ----------------------------------------------------------------------
def render_youtube(slide, data):
    BG_CANVAS = RGBColor(248, 250, 252)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_CANVAS
    bg.line.fill.background()

    draw_logo_badge(slide, Inches(1.0), Inches(0.35), badge_type="yt")

    tb = slide.shapes.add_textbox(Inches(2.2), Inches(0.28), Inches(8.9), Inches(0.9))
    p = tb.text_frame.paragraphs[0]
    p.text = str(data.get("date_title", "August'2026 31st"))
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(15, 23, 42)
    p.alignment = PP_ALIGN.CENTER
    p2 = tb.text_frame.add_paragraph()
    p2.text = str(data.get("gain_loss_text", ""))
    p2.font.size = Pt(12)
    p2.font.color.rgb = RGBColor(100, 116, 139)
    p2.alignment = PP_ALIGN.CENTER

    sub_data = data.get("subscribers_added", {})
    r1_y = Inches(1.4)
    w_card = Inches(3.6)
    gap = Inches(0.26)
    r1_items = [
        ("We Gained", sub_data.get("we_gained", "0"), RGBColor(16, 185, 129)),
        ("We Lost", sub_data.get("we_lost", "0"), RGBColor(239, 68, 68)),
        ("Total Subscribers", sub_data.get("total_subscribers", "0"), RGBColor(99, 102, 241))
    ]
    for i, (lbl, val, col) in enumerate(r1_items):
        x = Inches(1.0) + i * (w_card + gap)
        c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, r1_y, w_card, Inches(1.3))
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c.line.color.rgb = RGBColor(226, 232, 240)
        p = c.text_frame.paragraphs[0]
        p.text = lbl
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(100, 116, 139)
        p2 = c.text_frame.add_paragraph()
        p2.text = str(val)
        p2.font.size = Pt(26)
        p2.font.bold = True
        p2.font.color.rgb = col

    kpis = data.get("kpi_cards", [])
    r2_y = Inches(3.0)
    w_kpi = Inches(2.66)
    gap_kpi = Inches(0.23)
    icon_map = ["▶️", "👁️", "⏱️", "👥"]
    for i, k in enumerate(kpis[:4]):
        x = Inches(1.0) + i * (w_kpi + gap_kpi)
        c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, r2_y, w_kpi, Inches(1.7))
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor(241, 245, 249)
        c.line.color.rgb = RGBColor(226, 232, 240)
        p = c.text_frame.paragraphs[0]
        p.text = f"{icon_map[i]} {str(k.get('label', ''))}"
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(71, 85, 105)
        p.alignment = PP_ALIGN.CENTER
        p2 = c.text_frame.add_paragraph()
        p2.text = str(k.get("value", ""))
        p2.font.size = Pt(24)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(14, 116, 144)
        p2.alignment = PP_ALIGN.CENTER

    r3_y = Inches(5.0)
    sp_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), r3_y, Inches(11.333), Inches(1.8))
    sp_card.fill.solid()
    sp_card.fill.fore_color.rgb = RGBColor(254, 249, 195)
    sp_card.line.color.rgb = RGBColor(254, 240, 138)
    
    draw_money_bag_sticker(slide, Inches(8.5), r3_y + Inches(0.35))
    
    p = sp_card.text_frame.paragraphs[0]
    p.text = "💰 Spends"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RGBColor(161, 98, 7)
    p2 = sp_card.text_frame.add_paragraph()
    p2.text = f"₹ {data.get('spends', '0')}"
    p2.font.size = Pt(36)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(202, 138, 4)

# ----------------------------------------------------------------------
# Template 5: Medvarsity Performance
# ----------------------------------------------------------------------
def render_medvarsity(slide, data):
    BG_CANVAS = RGBColor(248, 250, 252)
    NAVY = RGBColor(15, 45, 110)
    CARD_BG = RGBColor(255, 255, 255)
    BORDER = RGBColor(226, 232, 240)

    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_CANVAS
    bg.line.fill.background()

    b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.25), Inches(12.533), Inches(0.85))
    b.fill.solid()
    b.fill.fore_color.rgb = NAVY
    b.line.fill.background()
    p = b.text_frame.paragraphs[0]
    p.text = str(data.get("title", "MEDVARSITY PERFORMANCE DASHBOARD")).upper()
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    p2 = b.text_frame.add_paragraph()
    p2.text = str(data.get("period", "Data Till 11th September 2026"))
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = RGBColor(191, 219, 254)
    p2.alignment = PP_ALIGN.CENTER

    top_kpis = data.get("top_kpis", [])
    kpi_y = Inches(1.2)
    w_k = Inches(1.98)
    gap_k = Inches(0.13)
    for i, k in enumerate(top_kpis[:6]):
        x = Inches(0.4) + i * (w_k + gap_k)
        c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, kpi_y, w_k, Inches(1.1))
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_BG
        c.line.color.rgb = BORDER
        p = c.text_frame.paragraphs[0]
        p.text = str(k.get("label", ""))
        p.font.size = Pt(8.5)
        p.font.color.rgb = RGBColor(100, 116, 139)
        p2 = c.text_frame.add_paragraph()
        p2.text = f"{str(k.get('value', ''))}  ↑"
        p2.font.size = Pt(18)
        p2.font.bold = True
        p2.font.color.rgb = NAVY

    col_y = Inches(2.42)
    col_h = Inches(3.45)
    c1_w, c2_w, c3_w = Inches(3.8), Inches(4.5), Inches(3.95)

    c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), col_y, c1_w, col_h)
    c1.fill.solid()
    c1.fill.fore_color.rgb = CARD_BG
    c1.line.color.rgb = BORDER
    tb1 = slide.shapes.add_textbox(Inches(0.5), col_y + Inches(0.08), c1_w - Inches(0.2), Inches(0.35))
    p = tb1.text_frame.paragraphs[0]
    p.text = "Marketing Spend Summary"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY

    m_spends = data.get("marketing_spends", [])
    tbl_s = slide.shapes.add_table(len(m_spends), 2, Inches(0.5), col_y + Inches(0.45), c1_w - Inches(0.2), col_h - Inches(0.6)).table
    for r, item in enumerate(m_spends):
        tbl_s.cell(r, 0).text = str(item.get("label", ""))
        tbl_s.cell(r, 1).text = str(item.get("value", ""))
        p0 = tbl_s.cell(r, 0).text_frame.paragraphs[0]
        p0.font.size = Pt(7.5)
        p0.font.color.rgb = RGBColor(30, 41, 59)
        p1 = tbl_s.cell(r, 1).text_frame.paragraphs[0]
        p1.font.size = Pt(8.5)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        p1.alignment = PP_ALIGN.RIGHT

    c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.35), col_y, c2_w, col_h)
    c2.fill.solid()
    c2.fill.fore_color.rgb = CARD_BG
    c2.line.color.rgb = BORDER
    tb2 = slide.shapes.add_textbox(Inches(4.45), col_y + Inches(0.08), c2_w - Inches(0.2), Inches(0.35))
    p = tb2.text_frame.paragraphs[0]
    p.text = "Leads Assigned to Sales Team"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY

    s_items = data.get("sales_assigned", [])
    tbl_a = slide.shapes.add_table(len(s_items) + 1, 4, Inches(4.45), col_y + Inches(0.45), c2_w - Inches(0.2), col_h - Inches(0.6)).table
    for col_i, h in enumerate(["Manager", "New", "Reactivated", "Total"]):
        c = tbl_a.cell(0, col_i)
        c.text = h
        p = c.text_frame.paragraphs[0]
        p.font.size = Pt(8)
        p.font.bold = True
        p.font.color.rgb = NAVY
    for r, item in enumerate(s_items):
        tbl_a.cell(r + 1, 0).text = f"👤 {str(item.get('manager', ''))}"
        tbl_a.cell(r + 1, 1).text = str(item.get("new", ""))
        tbl_a.cell(r + 1, 2).text = str(item.get("reactivated", ""))
        tbl_a.cell(r + 1, 3).text = str(item.get("total", ""))
        for col_i in range(4):
            p = tbl_a.cell(r + 1, col_i).text_frame.paragraphs[0]
            p.font.size = Pt(7.5)
            p.font.color.rgb = RGBColor(30, 41, 59)
            if col_i > 0: p.alignment = PP_ALIGN.CENTER

    c3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.98), col_y, c3_w, col_h)
    c3.fill.solid()
    c3.fill.fore_color.rgb = CARD_BG
    c3.line.color.rgb = BORDER
    tb3 = slide.shapes.add_textbox(Inches(9.08), col_y + Inches(0.08), c3_w - Inches(0.2), Inches(0.35))
    p = tb3.text_frame.paragraphs[0]
    p.text = "Revenue Achieved - Team Wise"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY

    r_items = data.get("revenue_team", [])
    tbl_r = slide.shapes.add_table(len(r_items) + 1, 2, Inches(9.08), col_y + Inches(0.45), c3_w - Inches(0.2), Inches(2.2)).table
    tbl_r.cell(0, 0).text = "Team Member"
    tbl_r.cell(0, 1).text = "Revenue"
    for col_i in [0, 1]:
        p = tbl_r.cell(0, col_i).text_frame.paragraphs[0]
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = NAVY
    for r, item in enumerate(r_items):
        tbl_r.cell(r + 1, 0).text = f"👤 {str(item.get('member', ''))}"
        tbl_r.cell(r + 1, 1).text = f"₹ {item.get('revenue', '')}"
        p0 = tbl_r.cell(r + 1, 0).text_frame.paragraphs[0]
        p0.font.size = Pt(8)
        p1 = tbl_r.cell(r + 1, 1).text_frame.paragraphs[0]
        p1.font.size = Pt(8.5)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(14, 116, 144)
        p1.alignment = PP_ALIGN.RIGHT

    ins_y = Inches(6.0)
    w_ins = Inches(3.02)
    gap_ins = Inches(0.15)
    insights = data.get("insights", [])
    color_map = {
        "dark_blue": RGBColor(15, 23, 42),
        "green": RGBColor(16, 185, 129),
        "purple": RGBColor(99, 102, 241),
        "pink": RGBColor(236, 72, 153)
    }
    for i, ins in enumerate(insights[:4]):
        x = Inches(0.4) + i * (w_ins + gap_ins)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, ins_y, w_ins, Inches(1.25))
        card.fill.solid()
        card.fill.fore_color.rgb = color_map.get(ins.get("color", "dark_blue"), RGBColor(15, 23, 42))
        card.line.fill.background()
        p = card.text_frame.paragraphs[0]
        p.text = str(ins.get("title", "")).upper()
        p.font.size = Pt(7.5)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p2 = card.text_frame.add_paragraph()
        p2.text = str(ins.get("desc", ""))
        p2.font.size = Pt(7.5)
        p2.font.color.rgb = RGBColor(241, 245, 249)

# Master Builder
def build_deck_for_all(results, filename="Executive_Exact_Dashboards.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for data in results:
        slide = prs.slides.add_slide(blank_layout)
        t_type = data.get("template_type", "LEADS_OVERVIEW")
        if t_type == "LEADS_OVERVIEW":
            render_leads_overview(slide, data)
        elif t_type == "ASSIMILATE":
            render_assimilate(slide, data)
        elif t_type == "YOUTUBE":
            render_youtube(slide, data)
        elif t_type == "MEDVARSITY":
            render_medvarsity(slide, data)
        else:
            render_leads_overview(slide, data)

    prs.save(filename)
    return filename

# ----------------------------------------------------------------------
# Streamlit Interface
# ----------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload Dashboard Screenshots (Upload 1 or all 5)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if uploaded_files:
    cols = st.columns(min(len(uploaded_files), 5))
    loaded = []
    for idx, f in enumerate(uploaded_files):
        img = Image.open(f)
        loaded.append((f.name, img))
        with cols[idx % 5]:
            st.image(img, caption=f.name, use_container_width=True)

    st.markdown("---")
    if st.button("🚀 Generate Exact PPT Slides", type="primary", use_container_width=True):
        if not api_key:
            st.error("Please enter your Gemini API Key in the sidebar.")
        else:
            all_results = []
            progress = st.progress(0)
            status = st.empty()

            for i, (name, img) in enumerate(loaded):
                status.text(f"Extracting & formatting ({i+1}/{len(loaded)}): {name}...")
                try:
                    res = extract_dashboard_data(img, api_key)
                    all_results.append(res)
                except Exception as e:
                    st.error(f"Error on {name}: {str(e)}")
                progress.progress((i + 1) / len(loaded))

            if all_results:
                pptx_out = "Executive_Exact_Dashboards.pptx"
                build_deck_for_all(all_results, pptx_out)
                st.success(f"Generated {len(all_results)} exact executive presentation slides!")
                with open(pptx_out, "rb") as f:
                    st.download_button(
                        label="📥 Download Exact Executive Presentation (.pptx)",
                        data=f,
                        file_name=pptx_out,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
