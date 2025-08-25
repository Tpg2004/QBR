# app.py
# A professional, single-file Streamlit application to generate comprehensive, AI-powered QBR decks.
# Version 7: Enhanced PowerPoint alignment and professional layout.

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import datetime
import os
import time

# --- 1. BACKEND LOGIC: ENHANCED DATA & PRESENTATION GENERATION ---

# --- Color Palette & Helper ---
PRIMARY_COLOR_PPT = RGBColor(10, 47, 87)    # Deep Navy for PPT
ACCENT_COLOR_PPT = RGBColor(0, 122, 255)    # Professional Blue for PPT
TEXT_COLOR_PPT = RGBColor(33, 33, 33)      # Dark Gray for PPT

def rgb_to_hex(rgb_color_obj):
    """Converts a python-pptx RGBColor object to a hex string for matplotlib."""
    r, g, b = rgb_color_obj
    return f"#{r:02x}{g:02x}{b:02x}"

def get_enhanced_mock_data(customer_name):
    """Generates a rich, multi-faceted dataset for a comprehensive QBR."""
    np.random.seed(hash(customer_name) % (2**32 - 1))

    kpis = {
        "Account Health Score": np.random.randint(75, 98), "NPS": np.random.randint(30, 65),
        "Product Adoption (%)": np.random.randint(60, 95), "Active Users": f"{np.random.randint(150, 500)}",
        "Support Tickets Closed": np.random.randint(25, 100), "Renewal Date": (datetime.date.today() + datetime.timedelta(days=np.random.randint(90, 365))).strftime('%Y-%m-%d')
    }
    commit_vs_actual = pd.DataFrame({
        'Metric': ['Feature Delivery', 'Uptime SLA', 'Avg. Ticket Response'],
        'Commitment': ['5 New Features', '99.9% Uptime', '< 8 Hours'],
        'Actual': ['6 New Features', f"{99.9 + np.random.uniform(0, 0.09):.2f}% Uptime", f"{np.random.uniform(6, 7.9):.1f} Hours"],
        'Status': ['Exceeded', 'Met', 'Met']
    })
    challenges = [
        "Initial onboarding for the new analytics module was slower than anticipated.",
        "Integration with the legacy CRM system required custom development work.",
        "User adoption in the finance department is lagging behind other teams."
    ]
    learnings = [
        "A dedicated onboarding webinar for new modules significantly boosts initial adoption.",
        "Pre-sales technical discovery for legacy systems is crucial to scope integrations accurately.",
        "Targeted training sessions and identifying team champions accelerate department-level adoption."
    ]
    okrs = pd.DataFrame({
        'Objective': ["Enhance Team Collaboration", "Enhance Team Collaboration", "Improve Data-Driven Decisions", "Improve Data-Driven Decisions"],
        'Key Result': ["Increase shared dashboard usage by 20%", "Onboard the marketing team to the platform", "Complete integration with BI tool", "Train 5 team leads on advanced reporting"],
        'Status': ['Not Started', 'Not Started', 'Not Started', 'Not Started']
    })
    roadmap = {
        "Q4 2025": ["AI-Powered Insights Engine", "Mobile App V2 Launch"],
        "Q1 2026": ["Advanced API Access", "Third-Party Integration Marketplace"],
        "Q2 2026": ["Predictive Analytics Module", "Team-based Permissions V3"]
    }
    revenue_forecast = pd.DataFrame({
        'Month': pd.to_datetime([f'2025-{i}-01' for i in range(10, 13)]),
        'Forecasted Revenue ($K)': [50 + i*5 + np.random.randint(-5, 5) for i in range(3)]
    })
    action_plan = pd.DataFrame({
        'Action Item': ["Schedule onboarding for marketing team", "Finalize BI tool integration specs", "Identify and train finance dept. champions"],
        'Owner': ["John (CSM)", "Jane (Customer IT)", "Sarah (CSM)"],
        'Due Date': [(datetime.date.today() + datetime.timedelta(days=d)).strftime('%Y-%m-%d') for d in [14, 30, 45]],
        'Status': ['Not Started', 'Not Started', 'Not Started']
    })
    return locals()

def create_revenue_chart(revenue_df, customer_name, output_path="revenue_chart.png"):
    """Creates a bar chart for the revenue forecast."""
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(8, 4))
    chart_color = rgb_to_hex(ACCENT_COLOR_PPT)
    sns.barplot(x=revenue_df['Month'].dt.strftime('%b'), y='Forecasted Revenue ($K)', data=revenue_df, color=chart_color, ax=ax)
    ax.set_title(f'Next Quarter Revenue Forecast', fontsize=14, weight='bold', color=rgb_to_hex(PRIMARY_COLOR_PPT))
    ax.set_xlabel('Month', fontsize=10); ax.set_ylabel('Forecasted Revenue ($K)', fontsize=10)
    plt.tight_layout(); plt.savefig(output_path, dpi=300, transparent=True)
    return output_path

def add_table_to_slide(slide, df, x, y, cx, cy):
    """Helper function to add a styled pandas DataFrame as a table to a slide."""
    shape = slide.shapes.add_table(df.shape[0] + 1, df.shape[1], x, y, cx, cy)
    table = shape.table
    # Set column widths to be equal
    for i in range(df.shape[1]): table.columns[i].width = int(cx / df.shape[1])
    for i, col_name in enumerate(df.columns):
        cell = table.cell(0, i); cell.text = col_name
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        cell.fill.solid(); cell.fill.fore_color.rgb = PRIMARY_COLOR_PPT
    for i, row in df.iterrows():
        for j, value in enumerate(row): table.cell(i + 1, j).text = str(value)

def create_professional_qbr_deck(data):
    """Builds a comprehensive, professionally styled PowerPoint presentation."""
    prs = Presentation(); prs.slide_width = Inches(16); prs.slide_height = Inches(9)
    def add_title_slide(title_text, subtitle_text):
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title_text; slide.placeholders[1].text = subtitle_text
    def add_content_slide(title_text):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = title_text; return slide, slide.placeholders[1]

    add_title_slide(f"Quarterly Business Review: {data['customer_name']}", f"Q3 2025 Report | Prepared: {datetime.date.today().strftime('%B %d, %Y')}")
    
    slide, content_placeholder = add_content_slide("Agenda")
    content = content_placeholder.text_frame
    content.clear() # Clear existing text
    topics = ["Quarterly Snapshot & Highlights", "Commitment Review", "Challenges & Key Learnings", "Objectives for Next Quarter (OKRs)", "Strategic Growth & Product Roadmap", "Commercial Outlook", "Joint Action Plan & Next Steps"]
    for topic in topics: p = content.add_paragraph(); p.text = topic; p.level = 0; p.space_after = Pt(12)

    slide, _ = add_content_slide("Quarterly Snapshot: Key Metrics")
    # *** ALIGNMENT FIX: Calculated spacing for perfect centering of KPI boxes ***
    num_kpis = len(data['kpis']); box_width = 2.5; gap = 0.5
    total_width = num_kpis * box_width + (num_kpis - 1) * gap
    start_x = (16 - total_width) / 2
    for i, (key, value) in enumerate(data['kpis'].items()):
        left = Inches(start_x + i * (box_width + gap))
        txBox = slide.shapes.add_textbox(left, Inches(2.5), Inches(box_width), Inches(2)); tf = txBox.text_frame
        p_val = tf.add_paragraph(); p_val.text = str(value); p_val.font.bold = True; p_val.font.size = Pt(44); p_val.alignment = PP_ALIGN.CENTER
        p_key = tf.add_paragraph(); p_key.text = key; p_key.font.size = Pt(16); p_key.alignment = PP_ALIGN.CENTER

    slide, _ = add_content_slide("Commitment Review: Promises vs. Reality")
    add_table_to_slide(slide, data['commit_vs_actual'], Inches(1), Inches(2.5), Inches(14), Inches(4))

    slide, _ = add_content_slide("Challenges & Key Learnings")
    # *** ALIGNMENT FIX: Defined precise positions for Challenge/Learning boxes ***
    box_width_cl, gap_cl = 7.0, 1.0
    left_cl1 = (16 - (2 * box_width_cl + gap_cl)) / 2
    left_cl2 = left_cl1 + box_width_cl + gap_cl
    txBox1 = slide.shapes.add_textbox(Inches(left_cl1), Inches(2.5), Inches(box_width_cl), Inches(5)); tf1 = txBox1.text_frame
    tf1.text = "Challenges Faced This Quarter"; tf1.paragraphs[0].font.bold = True
    for item in data['challenges']: p = tf1.add_paragraph(); p.text = item; p.level = 1
    txBox2 = slide.shapes.add_textbox(Inches(left_cl2), Inches(2.5), Inches(box_width_cl), Inches(5)); tf2 = txBox2.text_frame
    tf2.text = "Key Lessons Learned"; tf2.paragraphs[0].font.bold = True
    for item in data['learnings']: p = tf2.add_paragraph(); p.text = item; p.level = 1

    slide, _ = add_content_slide("Objectives for Next Quarter (OKRs)")
    add_table_to_slide(slide, data['okrs'], Inches(1), Inches(2.5), Inches(14), Inches(5))

    slide, _ = add_content_slide("Strategic Growth & Product Roadmap")
    # *** ALIGNMENT FIX: Calculated spacing for roadmap columns ***
    num_roadmap = len(data['roadmap']); box_width_r = 4.5; gap_r = 1.0
    total_width_r = num_roadmap * box_width_r + (num_roadmap - 1) * gap_r
    start_x_r = (16 - total_width_r) / 2
    for i, (quarter, features) in enumerate(data['roadmap'].items()):
        left = Inches(start_x_r + i * (box_width_r + gap_r))
        txBox = slide.shapes.add_textbox(left, Inches(2.5), Inches(box_width_r), Inches(5)); tf = txBox.text_frame
        p_qtr = tf.add_paragraph(); p_qtr.text = quarter; p_qtr.font.bold = True; p_qtr.font.size = Pt(24)
        for feature in features: p_feat = tf.add_paragraph(); p_feat.text = "• " + feature; p_feat.level = 0

    slide, _ = add_content_slide("Commercial Outlook: Revenue Forecast")
    chart_path = create_revenue_chart(data['revenue_forecast'], data['customer_name'])
    slide.shapes.add_picture(chart_path, Inches(2), Inches(2.0), width=Inches(12)); os.remove(chart_path)

    slide, _ = add_content_slide("Joint Action Plan & Owners")
    add_table_to_slide(slide, data['action_plan'], Inches(1), Inches(2.5), Inches(14), Inches(4))

    add_title_slide("Thank You", "Q&A and Discussion")
    output_filename = f"QBR_{data['customer_name'].replace(' ', '_')}_{datetime.date.today()}.pptx"; prs.save(output_filename); return output_filename

# --- 2. FRONTEND UI: PROFESSIONAL STREAMLIT APPLICATION ---

st.set_page_config(page_title="AI QBR Deck Generator", page_icon="✨", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');
    body { font-family: 'Poppins', sans-serif; }
    .stApp { background: #f0f2f6; }
    .main-header { font-size: 2.5rem; font-weight: 700; text-align: center; margin-bottom: 0px; color: #0a2f57; }
    .sub-header { font-size: 1.1rem; text-align: center; color: #555; margin-bottom: 2rem; }
    .stButton>button {
        background-image: linear-gradient(to right, #007bff 0%, #0056b3 51%, #007bff 100%);
        color: white; border-radius: 10px; transition: 0.5s; background-size: 200% auto;
        font-weight: 600; border: none; height: 3em; width: 100%;
    }
    .stButton>button:hover { background-position: right center; color: #fff; text-decoration: none; }
    .stTextInput>div>div>input { background-color: #fff; border-radius: 10px; }
    .info-card {
        background: white; border-radius: 15px; padding: 25px;
        box-shadow: 0 10px 20px rgba(0,0,0,0.05); border-left: 5px solid #007bff; height: 100%;
    }
    .info-card h3 { color: #0a2f57; font-weight: 600; }
    .stProgress > div > div > div > div { background-image: linear-gradient(to right, #007bff, #0056b3); }
    .stDownloadButton>button { background-image: linear-gradient(to right, #28a745 0%, #218838 51%, #28a745 100%); }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-header'>✨ AI QBR Deck Generator</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Transform customer data into a stunning, client-ready presentation in seconds.</p>", unsafe_allow_html=True)

col1, col2 = st.columns([1, 1.5])

with col1:
    st.markdown("<div class='info-card'><h3>Controls</h3></div>", unsafe_allow_html=True)
    customer_name = st.text_input("Enter Customer Name", "Innovate Corp", label_visibility="collapsed")
    
    if st.button("🚀 Generate QBR Deck"):
        if customer_name:
            with st.spinner('Crafting your presentation...'):
                progress_bar = st.progress(0, text="Initializing...")
                enhanced_data = get_enhanced_mock_data(customer_name)
                time.sleep(1); progress_bar.progress(25, text="Generating Insights...")
                time.sleep(1); progress_bar.progress(50, text="Creating Visualizations...")
                time.sleep(1); progress_bar.progress(75, text="Assembling Deck...")
                final_deck_path = create_professional_qbr_deck(enhanced_data)
                progress_bar.progress(100, text="Done!")
                st.success(f"🎉 Your QBR deck is ready!")
                with open(final_deck_path, "rb") as file:
                    st.download_button(
                        label="⬇️ Download Presentation", data=file, file_name=final_deck_path,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                os.remove(final_deck_path)
        else:
            st.warning("Please enter a customer name.")

with col2:
    with st.expander("See What's Inside Your AI-Generated Deck", expanded=True):
        st.markdown("""
        - **Agenda**: A clear outline of the review.
        - **Quarterly Snapshot**: At-a-glance view of all key performance indicators.
        - **Commit vs. Actual**: Transparent review of promises vs. reality.
        - **Challenges & Learnings**: Honest reflection to build stronger partnerships.
        - **Next Quarter OKRs**: Collaborative goal-setting for the future.
        - **Product Roadmap**: A look ahead at exciting new developments.
        - **Revenue Forecast**: Commercial outlook and pipeline discussion.
        - **Action Plan**: Clear, actionable next steps with assigned owners.
        """)

    st.markdown("<div class='info-card'><h3>Sample Visualization</h3></div>", unsafe_allow_html=True)
    sample_data = get_enhanced_mock_data("Sample Company")
    fig, ax = plt.subplots()
    sns.barplot(x=sample_data['revenue_forecast']['Month'].dt.strftime('%b'), y='Forecasted Revenue ($K)', data=sample_data['revenue_forecast'], color=rgb_to_hex(ACCENT_COLOR_PPT), ax=ax)
    ax.set_title("Revenue Forecast Visualization")
    st.pyplot(fig, use_container_width=True)

st.sidebar.info("This is a Proof-of-Concept. All data is realistically simulated for demonstration.")
