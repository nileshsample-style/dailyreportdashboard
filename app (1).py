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
    page_title="ExecSlide AI | Multi-Slide Executive Dashboard Generator",
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

st.title("📊 ExecSlide: Individual Image-to-Executive Dashboard")
st.markdown("Upload multiple data screenshots. The AI converts **each image into its own dedicated, C-Suite ready presentation slide** with exact extracted tables, top-row KPIs, and strategic observations for the VP/CEO.")

# ----------------------------------------------------------------------
# API Key Handling (Sidebar or Streamlit Secrets)
# ----------------------------------------------------------------------
default_key = ""
if "GEMINI_API_KEY" in st.secrets:
    default_key = st.secrets["GEMINI_API_KEY"]
elif os.getenv("GEMINI_API_KEY"):
    default_key = os.getenv("GEMINI_API_KEY")

api_key = st.sidebar.text_input("Gemini API Key", type="password", value=default_key, help="Get your key at https://aistudio.google.com")

if not api_key:
    st.sidebar.warning("⚠️ Enter your Gemini API Key to enable automated visual extraction.")
    st.sidebar.info("Tip: In Streamlit Cloud, store this under App Settings > Secrets as GEMINI_API_KEY.")

# ----------------------------------------------------------------------
# Executive Vision Prompt for INDIVIDUAL Image Extraction
# ----------------------------------------------------------------------
SINGLE_IMAGE_PROMPT = """
You are an elite Chief of Staff and Strategy Director preparing an executive slide for the CEO and VP.
Examine this single dashboard image and extract all its specific numbers, data tables, and metrics.
Return a STRICT JSON object representing this exact template:
{
    "slide_title": "Concise Executive Title (e.g., YOUTUBE CHANNEL GROWTH & RETENTION)",
    "reporting_period": "Exact reporting date or period identified from the image",
    "headline_kpis": [
        {"label": "KPI 1 NAME", "value": "12,345", "subtext": "Contextual subtext or % change"},
        {"label": "KPI 2 NAME", "value": "₹50,000", "subtext": "Contextual subtext"},
        {"label": "KPI 3 NAME", "value": "94.2%", "subtext": "Target achievement"},
        {"label": "KPI 4 NAME", "value": "4.2 Days", "subtext": "Velocity indicator"},
        {"label": "KPI 5 NAME", "value": "₹155.65", "subtext": "Unit economic / CPL"}
    ],
    "table_headers": ["Category / Dimension", "Volume / Count", "Share / Performance", "Status / Next Step"],
    "table_rows": [
        ["Row Item 1", "Val 1", "Val 2", "Val 3"],
        ["Row Item 2", "Val 1", "Val 2", "Val 3"],
        ["Row Item 3", "Val 1", "Val 2", "Val 3"]
    ],
    "key_insights": [
        "Concise, high-impact bullet on top-line performance or acquisition velocity.",
        "Concise observation on funnel efficiency, operational leakage, or bottlenecks.",
        "Concise, high-priority recommendation for executive decision making and resource allocation."
    ]
}
Note:
- headline_kpis should have 4 to 5 key metrics extracted from the image.
- table_rows should accurately capture the primary breakdown table from the image (up to 7 rows).
- Do NOT include markdown blocks like ```json. Output raw JSON only.
"""

def extract_single_image(img, key):
    client = genai.Client(api_key=key)
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
                    contents=[SINGLE_IMAGE_PROMPT, img],
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
                time.sleep(2)
                continue
    raise last_err

# ----------------------------------------------------------------------
# PPTX Generator: Multi-Slide Deck with Navy & Teal Executive Design
# ----------------------------------------------------------------------
def build_deck(slides_data, filename="Executive_Individual_Dashboards.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Theme Colors
    BG_COLOR = RGBColor(11, 19, 41)       # Deep Navy
    CARD_BG = RGBColor(22, 34, 63)        # Slate Blue Container
    TEAL = RGBColor(34, 211, 238)         # Cyan/Teal Accent
    WHITE = RGBColor(255, 255, 255)
    GRAY = RGBColor(156, 163, 175)
    HEADER_BLUE = RGBColor(14, 116, 144)  # Table Header Dark Teal
    GREEN = RGBColor(52, 211, 153)

    for data in slides_data:
        slide = prs.slides.add_slide(blank_layout)

        # 1. Slide Background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_COLOR
        bg.line.fill.background()

        # 2. Header
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = str(data.get("slide_title", "EXECUTIVE DASHBOARD")).upper()
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = WHITE

        p2 = tf.add_paragraph()
        p2.text = f"Reporting Period: {data.get('reporting_period', 'September 2026')} | Executive Strategy Review"
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEAL

        # 3. Top Row KPI Cards
        kpis = data.get("headline_kpis", [])[:5]
        n = max(len(kpis), 1)
        total_w = Inches(11.733)
        gap = Inches(0.18)
        card_w = (total_w - (n - 1) * gap) / n
        start_x = Inches(0.8)
        card_h = Inches(1.2)
        start_y = Inches(1.4)

        for i, kpi in enumerate(kpis):
            x = start_x + i * (card_w + gap)
            rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, start_y, card_w, card_h)
            rect.fill.solid()
            rect.fill.fore_color.rgb = CARD_BG
            rect.line.color.rgb = TEAL
            rect.line.width = Pt(1)

            ktf = rect.text_frame
            ktf.margin_left = Inches(0.12)
            ktf.margin_top = Inches(0.08)

            kp0 = ktf.paragraphs[0]
            kp0.text = str(kpi.get("label", "")).upper()
            kp0.font.size = Pt(9)
            kp0.font.color.rgb = GRAY

            kp1 = ktf.add_paragraph()
            kp1.text = str(kpi.get("value", ""))
            kp1.font.size = Pt(19)
            kp1.font.bold = True
            kp1.font.color.rgb = WHITE

            kp2 = ktf.add_paragraph()
            kp2.text = str(kpi.get("subtext", ""))
            kp2.font.size = Pt(8.5)
            kp2.font.color.rgb = GREEN

        # 4. Insights Card (Left Side)
        left_x, left_y, left_w, left_h = Inches(0.8), Inches(2.8), Inches(5.6), Inches(4.3)
        left_rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_x, left_y, left_w, left_h)
        left_rect.fill.solid()
        left_rect.fill.fore_color.rgb = CARD_BG
        left_rect.line.color.rgb = RGBColor(30, 58, 95)

        ltf = left_rect.text_frame
        ltf.word_wrap = True
        ltf.margin_left = Inches(0.25)
        ltf.margin_top = Inches(0.2)
        ltf.margin_right = Inches(0.25)

        lp0 = ltf.paragraphs[0]
        lp0.text = "STRATEGIC OBSERVATIONS (VP / CEO)"
        lp0.font.size = Pt(12)
        lp0.font.bold = True
        lp0.font.color.rgb = TEAL

        for ins in data.get("key_insights", []):
            lip = ltf.add_paragraph()
            lip.text = f"• {ins}"
            lip.font.size = Pt(9.5)
            lip.font.color.rgb = WHITE
            lip.space_before = Pt(7)

        # 5. Data Breakdown Table (Right Side)
        headers = data.get("table_headers", ["Category", "Volume", "Share", "Status"])
        rows_data = data.get("table_rows", [])[:7]
        table_x, table_y, table_w, table_h = Inches(6.6), Inches(2.8), Inches(5.9), Inches(4.3)

        num_rows = len(rows_data) + 1
        num_cols = len(headers)
        table_shape = slide.shapes.add_table(num_rows, num_cols, table_x, table_y, table_w, table_h)
        table = table_shape.table

        for c, h_text in enumerate(headers):
            cell = table.cell(0, c)
            cell.text = str(h_text)
            cell.fill.solid()
            cell.fill.fore_color.rgb = HEADER_BLUE
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(9.5)
                p.font.bold = True
                p.font.color.rgb = WHITE

        for r, row_vals in enumerate(rows_data):
            for c, val in enumerate(row_vals[:num_cols]):
                cell = table.cell(r + 1, c)
                cell.text = str(val)
                cell.fill.solid()
                cell.fill.fore_color.rgb = CARD_BG
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(8.5)
                    p.font.color.rgb = WHITE

    prs.save(filename)
    return filename

# ----------------------------------------------------------------------
# Streamlit Web UI Flow
# ----------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "Upload Your Raw Dashboard Screenshots (JPG / PNG)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True,
    help="Upload your screenshots (e.g., YouTube Metrics, Assimilate, HBF, Fellowship, Medvarsity)."
)

if uploaded_files:
    cols = st.columns(min(len(uploaded_files), 5))
    loaded_images = []
    for idx, file in enumerate(uploaded_files):
        img = Image.open(file)
        loaded_images.append((file.name, img))
        with cols[idx % 5]:
            st.image(img, caption=file.name, use_container_width=True)

    st.markdown("---")
    if st.button("🚀 Generate Individual Executive Slides for Each File", type="primary", use_container_width=True):
        if not api_key:
            st.error("Please provide a Gemini API Key in the sidebar or via Streamlit Secrets.")
        else:
            all_slides_data = []
            progress_bar = st.progress(0)
            status_text = st.empty()

            for idx, (fname, img) in enumerate(loaded_images):
                status_text.text(f"Processing ({idx+1}/{len(loaded_images)}): {fname}...")
                try:
                    slide_result = extract_single_image(img, api_key)
                    all_slides_data.append(slide_result)
                except Exception as e:
                    st.error(f"Error processing {fname}: {str(e)}")
                progress_bar.progress((idx + 1) / len(loaded_images))

            status_text.text("Building executive multi-slide deck...")
            
            if all_slides_data:
                pptx_filename = "Executive_Individual_Dashboards.pptx"
                build_deck(all_slides_data, pptx_filename)
                st.success(f"Successfully generated {len(all_slides_data)} customized executive slides!")

                with open(pptx_filename, "rb") as f:
                    st.download_button(
                        label="📥 Download Complete Executive PPTX Deck",
                        data=f,
                        file_name=pptx_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

                # Expandable preview of each slide in Streamlit
                for i, slide_data in enumerate(all_slides_data):
                    with st.expander(f"Slide {i+1} Preview: {slide_data.get('slide_title', 'Dashboard')}", expanded=True):
                        st.caption(f"Period: {slide_data.get('reporting_period', 'N/A')}")
                        k_cols = st.columns(len(slide_data.get("headline_kpis", [])))
                        for j, kpi in enumerate(slide_data.get("headline_kpis", [])):
                            with k_cols[j]:
                                st.metric(kpi.get("label"), kpi.get("value"), kpi.get("subtext"))
                        c1, c2 = st.columns([1, 1])
                        with c1:
                            st.markdown("**Strategic Observations:**")
                            for obs in slide_data.get("key_insights", []):
                                st.info(obs)
                        with c2:
                            st.markdown("**Extracted Data Matrix:**")
                            headers = slide_data.get("table_headers", [])
                            rows = slide_data.get("table_rows", [])
                            st.table([dict(zip(headers, row)) for row in rows])
