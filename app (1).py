import os
import json
import time
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from google import genai
from google.genai import types

# ----------------------------------------------------------------------
# Page Configuration & Executive Theme Styling
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="ExecSlide AI | Executive PPT Dashboard",
    page_icon="📊",
    layout="wide"
)

# Custom Executive Dark Navy & Teal Styling
st.markdown("""
    <style>
    .main { background-color: #0b1329; color: #f8fafc; }
    .stMetric { background-color: #16223f; border: 1px solid #1e3a5f; border-radius: 8px; padding: 12px; }
    h1, h2, h3 { color: #f8fafc; font-weight: 700; }
    div[data-testid="stExpander"] { background-color: #16223f; border: 1px solid #1e3a5f; border-radius: 8px; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 ExecSlide: Image-to-Executive Dashboard")
st.markdown("Transform dashboard screenshots, spreadsheets, or metric graphics directly into **C-Suite ready PowerPoint slides** with automated strategic takeaways.")

# ----------------------------------------------------------------------
# API Key Handling (Sidebar or Streamlit Secrets)
# ----------------------------------------------------------------------
default_key = ""
if "GEMINI_API_KEY" in st.secrets:
    default_key = st.secrets["GEMINI_API_KEY"]
elif os.getenv("GEMINI_API_KEY"):
    default_key = os.getenv("GEMINI_API_KEY")

api_key = st.sidebar.text_input("Gemini API Key", type="password", value=default_key, help="Get your key at [https://aistudio.google.com](https://aistudio.google.com)")

if not api_key:
    st.sidebar.warning("⚠️ Enter your Gemini API Key to enable automated visual extraction.")
    st.sidebar.info("Tip: In Streamlit Cloud, you can also store this under App Settings > Secrets as GEMINI_API_KEY.")

# ----------------------------------------------------------------------
# Analysis Prompt (Executive Level)
# ----------------------------------------------------------------------
ANALYSIS_PROMPT = """
You are an elite Chief of Staff and Strategy Director preparing a presentation for the CEO & VP.
Analyze the provided dashboard images and extract all relevant data, cross-referencing and consolidating performance.
Return a STRICT JSON object with this exact structure:
{
    "dashboard_title": "Concise Executive Title (e.g., Q3 PORTFOLIO & REVENUE STRATEGY REVIEW)",
    "reporting_period": "Reporting timeframe extracted from the images",
    "headline_kpis": [
        {"label": "TOTAL REVENUE", "value": "₹58.29 L", "subtext": "4.51% Spend-to-Rev"},
        {"label": "MARKETING SPENDS", "value": "₹2.63 L", "subtext": "Consolidated Digital"},
        {"label": "TOTAL CONVERSIONS", "value": "72 Deals", "subtext": "8.9% Lead-to-Deal"},
        {"label": "TOP PERFORMER", "value": "Bala (₹27.83L)", "subtext": "47.8% Total Revenue"}
    ],
    "channel_breakdown": [
        {"program": "Program / Category", "leads": "Metric Count", "spend": "Spend Value", "revenue": "Revenue Value", "note": "Strategic Note"}
    ],
    "key_insights": [
        "Concise, high-impact bullet point focusing on capital efficiency and ROI.",
        "Concise bullet point identifying bottlenecks or conversion capacity upsides.",
        "Concise bullet point recommending concrete operational resource reallocations."
    ]
}
Do NOT enclose the response in markdown backticks or code blocks. Output pure JSON only.
"""

def extract_and_analyze(images, key):
    client = genai.Client(api_key=key)
    contents = [ANALYSIS_PROMPT]
    for img in images:
        contents.append(img)
    
    # Priority list of exact models matching your Google AI Studio account
    candidate_models = [
        "gemini-3.8-flash",
        "gemini-3.5-flash-lite",
        "gemini-3.1-pro",
        "gemini-2.5-flash"
    ]
    
    last_err = None
    for model_name in candidate_models:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                text = response.text.strip()
                # Clean accidental markdown wrapping
                if text.startswith("```json"):
                    text = text[7:]
                elif text.startswith("```"):
                    text = text[3:]
                if text.endswith("```"):
                    text = text[:-3]
                return json.loads(text.strip())
            except Exception as e:
                last_err = e
                time.sleep(2)
                continue
                
    raise last_err

# ----------------------------------------------------------------------
# PPTX Generator: Deep Navy, Dark Blue & Teal Theme
# ----------------------------------------------------------------------
def create_executive_pptx(data, filename="Executive_Dashboard.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Color Palette
    BG_COLOR = RGBColor(11, 19, 41)       # Deep Navy
    CARD_BG = RGBColor(22, 34, 63)        # Card Slate Blue
    TEAL_ACCENT = RGBColor(34, 211, 238)  # Vibrant Cyan/Teal
    HEADER_BLUE = RGBColor(14, 116, 144)  # Table Header Dark Teal
    WHITE = RGBColor(255, 255, 255)
    GRAY = RGBColor(156, 163, 175)

    # 1. Slide Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_COLOR
    bg.line.fill.background()

    # 2. Slide Header
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.9))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(data.get("dashboard_title", "EXECUTIVE PERFORMANCE OVERVIEW")).upper()
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = WHITE

    p2 = tf.add_paragraph()
    p2.text = f"Consolidated Strategy Briefing | Reporting Period: {data.get('reporting_period', 'Active Window')}"
    p2.font.size = Pt(12)
    p2.font.color.rgb = TEAL_ACCENT

    # 3. Top Row KPI Cards (4 Cards)
    kpis = data.get("headline_kpis", [])[:4]
    card_w, card_h = Inches(2.75), Inches(1.2)
    start_x, start_y = Inches(0.8), Inches(1.4)
    gap = Inches(0.24)

    for i, kpi in enumerate(kpis):
        x = start_x + i * (card_w + gap)
        rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, start_y, card_w, card_h)
        rect.fill.solid()
        rect.fill.fore_color.rgb = CARD_BG
        rect.line.color.rgb = TEAL_ACCENT
        rect.line.width = Pt(1)

        ktf = rect.text_frame
        ktf.margin_left = Inches(0.15)
        ktf.margin_top = Inches(0.1)

        kp0 = ktf.paragraphs[0]
        kp0.text = str(kpi.get("label", "")).upper()
        kp0.font.size = Pt(9)
        kp0.font.color.rgb = GRAY

        kp1 = ktf.add_paragraph()
        kp1.text = str(kpi.get("value", ""))
        kp1.font.size = Pt(20)
        kp1.font.bold = True
        kp1.font.color.rgb = WHITE

        kp2 = ktf.add_paragraph()
        kp2.text = str(kpi.get("subtext", ""))
        kp2.font.size = Pt(8.5)
        kp2.font.color.rgb = TEAL_ACCENT

    # 4. Strategic Insights Card (Left Half)
    left_x, left_y, left_w, left_h = Inches(0.8), Inches(2.8), Inches(5.8), Inches(4.2)
    left_rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_x, left_y, left_w, left_h)
    left_rect.fill.solid()
    left_rect.fill.fore_color.rgb = CARD_BG
    left_rect.line.color.rgb = RGBColor(30, 58, 95)

    ltf = left_rect.text_frame
    ltf.margin_left = Inches(0.3)
    ltf.margin_top = Inches(0.25)
    ltf.margin_right = Inches(0.3)
    ltf.word_wrap = True

    lp0 = ltf.paragraphs[0]
    lp0.text = "C-SUITE STRATEGIC OBSERVATIONS"
    lp0.font.size = Pt(13)
    lp0.font.bold = True
    lp0.font.color.rgb = TEAL_ACCENT

    for insight in data.get("key_insights", []):
        lip = ltf.add_paragraph()
        lip.text = f"• {insight}"
        lip.font.size = Pt(10)
        lip.font.color.rgb = WHITE
        lip.space_before = Pt(8)

    # 5. Consolidated Data Table (Right Half)
    table_x, table_y, table_w, table_h = Inches(6.8), Inches(2.8), Inches(5.7), Inches(4.2)
    channels = data.get("channel_breakdown", [])
    rows = min(len(channels) + 1, 6)
    table_shape = slide.shapes.add_table(rows, 4, table_x, table_y, table_w, table_h)
    table = table_shape.table

    headers = ["Program / Stream", "Leads/Vol", "Spends", "Revenue"]
    for col_idx, h in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = HEADER_BLUE
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = WHITE

    for row_idx, item in enumerate(channels[:rows-1]):
        r = row_idx + 1
        vals = [
            str(item.get("program", "")),
            str(item.get("leads", "")),
            str(item.get("spend", "")),
            str(item.get("revenue", ""))
        ]
        for c_idx, val in enumerate(vals):
            cell = table.cell(r, c_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(9)
                p.font.color.rgb = WHITE

    prs.save(filename)
    return filename

# ----------------------------------------------------------------------
# Streamlit App UI Flow
# ----------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload Dashboard Screenshots (JPG / PNG)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True,
    help="Select one or multiple images showing your reports, funnels, or sales metrics."
)

if uploaded_files:
    cols = st.columns(min(len(uploaded_files), 5))
    loaded_images = []
    for idx, file in enumerate(uploaded_files):
        img = Image.open(file)
        loaded_images.append(img)
        with cols[idx % 5]:
            st.image(img, caption=file.name, use_container_width=True)

    st.markdown("---")
    if st.button("🚀 Synthesize Metrics & Generate C-Suite PPT", type="primary", use_container_width=True):
        if not api_key:
            st.error("Please provide a Gemini API Key in the sidebar or via Streamlit Secrets.")
        else:
            with st.spinner("Analyzing visuals, extracting KPI tables, and compiling strategic takeaways..."):
                try:
                    analysis_result = extract_and_analyze(loaded_images, api_key)
                    
                    st.success("Analysis Complete!")
                    st.subheader(analysis_result.get("dashboard_title", "Executive Overview"))
                    st.caption(f"Period: {analysis_result.get('reporting_period', 'Current')}")
                    
                    # KPIs in Streamlit UI
                    kpis = analysis_result.get("headline_kpis", [])
                    if kpis:
                        kpi_cols = st.columns(len(kpis))
                        for i, kpi in enumerate(kpis):
                            with kpi_cols[i]:
                                st.metric(label=kpi.get("label"), value=kpi.get("value"), delta=kpi.get("subtext"))
                    
                    c1, c2 = st.columns([1, 1])
                    with c1:
                        st.markdown("### 💡 Strategic Observations")
                        for ins in analysis_result.get("key_insights", []):
                            st.info(ins)
                    with c2:
                        st.markdown("### 📋 Program Summary")
                        st.dataframe(analysis_result.get("channel_breakdown", []), use_container_width=True)

                    # Export PPTX
                    pptx_filename = "CEO_Executive_Dashboard.pptx"
                    create_executive_pptx(analysis_result, pptx_filename)
                    with open(pptx_filename, "rb") as f:
                        st.download_button(
                            label="📥 Download Executive Presentation (.pptx)",
                            data=f,
                            file_name=pptx_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Error during processing: {str(e)}")
