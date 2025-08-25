# app.py
# A professional, single-file Streamlit application to generate comprehensive, AI-powered QBR decks.

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

# --- Color Palette for a Professional Look ---
PRIMARY_COLOR = RGBColor(13, 71, 161) # Deep Blue
ACCENT_COLOR = RGBColor(25, 118, 210) # Bright Blue
TEXT_COLOR = RGBColor(33, 33, 33)   # Dark Gray
BACKGROUND_COLOR = RGBColor(245, 245, 245) # Light Gray

def get_enhanced_mock_data(customer_name):
    """Generates a rich, multi-faceted dataset for a comprehensive QBR."""
    np.random.seed(hash(customer_name) % (2**32 - 1)) # Seed based on name for consistency

    # --- Quarterly Snapshot & KPIs ---
    kpis = {
        "Account Health Score": np.random.randint(75, 98),
        "NPS": np.random.randint(30, 65),
        "Product Adoption (%)": np.random.randint(60, 95),
        "Active Users": f"{np.random.randint(150, 500)}",
        "Support Tickets Closed": np.random.randint(25, 100),
        "Renewal Date": (datetime.date.today() + datetime.timedelta(days=np.random.randint(90, 365))).strftime('%Y-%m-%d')
    }

    # --- Commit vs Actual ---
    commit_vs_actual = pd.DataFrame({
        'Metric': ['Feature Delivery', 'Uptime SLA', 'Avg. Ticket Response'],
        'Commitment': ['5 New Features', '99.9% Uptime', '< 8 Hours'],
        'Actual': ['6 New Features', f"{99.9 + np.random.uniform(0, 0.09):.2f}% Uptime", f"{np.random.uniform(6, 7.9):.1f} Hours"],
        'Status': ['Exceeded', 'Met', 'Met']
    })

    # --- Challenges and Lessons ---
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

    # --- Next Quarter OKRs ---
    okrs = pd.DataFrame({
        'Objective': [
            "Enhance Team Collaboration", "Enhance Team Collaboration",
            "Improve Data-Driven Decisions", "Improve Data-Driven Decisions"
        ],
        'Key Result': [
            "Increase shared dashboard usage by 20%",
            "Onboard the marketing team to the platform",
            "Complete integration with BI tool",
            "Train 5 team leads on advanced reporting"
        ],
        'Target': [20, 1, 1, 5],
        'Actual': [0, 0, 0, 0],
        'Status': ['Not Started', 'Not Started', 'Not Started', 'Not Started']
    })

    # --- Product Roadmap ---
    roadmap = {
        "Q4 2025": ["AI-Powered Insights Engine", "Mobile App V2 Launch"],
        "Q1 2026": ["Advanced API Access", "Third-Party Integration Marketplace"],
        "Q2 2026": ["Predictive Analytics Module", "Team-based Permissions V3"]
    }

    # --- Sales & Revenue ---
    revenue_forecast = pd.DataFrame({
        'Month': pd.to_datetime([f'2025-{i}-01' for i in range(10, 13)]),
        'Forecasted Revenue ($K)': [50 + i*5 + np.random.randint(-5, 5) for i in range(3)]
    })

    # --- Action Plan ---
    action_plan = pd.DataFrame({
        'Action Item': [
            "Schedule onboarding for marketing team",
            "Finalize BI tool integration specs",
            "Identify and train finance dept. champions"
        ],
        'Owner': ["John (CSM)", "Jane (Customer IT)", "Sarah (CSM)"],
        'Due Date': [
            (datetime.date.today() + datetime.timedelta(days=14)).strftime('%Y-%m-%d'),
            (datetime.date.today() + datetime.timedelta(days=30)).strftime('%Y-%m-%d'),
            (datetime.date.today() + datetime.timedelta(days=45)).strftime('%Y-%m-%d')
        ],
        'Status': ['Not Started', 'Not Started', 'Not Started']
    })

    return {
        "customer_name": customer_name, "kpis": kpis, "commit_vs_actual": commit_vs_actual,
        "challenges": challenges, "learnings": learnings, "okrs": okrs,
        "roadmap": roadmap, "revenue_forecast": revenue_forecast, "action_plan": action_plan
    }

def create_revenue_chart(revenue_df, customer_name, output_path="revenue_chart.png"):
    """Creates a bar chart for the revenue forecast."""
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(7, 4))
    
    sns.barplot(x=revenue_df['Month'].dt.strftime('%b'), y='Forecasted Revenue ($K)', data=revenue_df, color=ACCENT_COLOR.to_hex(), ax=ax)
    
    ax.set_title(f'Next Quarter Revenue Forecast for {customer_name}', fontsize=14, weight='bold')
    ax.set_xlabel('Month', fontsize=10)
    ax.set_ylabel('Forecasted Revenue (in thousands)', fontsize=10)
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    return output_path

def add_table_to_slide(slide, df, x, y, cx, cy):
    """Helper function to add a styled pandas DataFrame as a table to a slide."""
    shape = slide.shapes.add_table(df.shape[0] + 1, df.shape[1], x, y, cx, cy)
    table = shape.table
    
    # Set column widths
    for i in range(df.shape[1]):
        table.columns[i].width = Inches(2.0)

    # Write table headers
    for i, col_name in enumerate(df.columns):
        cell = table.cell(0, i)
        cell.text = col_name
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PRIMARY_COLOR

    # Write table data
    for i, row in df.iterrows():
        for j, value in enumerate(row):
            table.cell(i + 1, j).text = str(value)

def create_professional_qbr_deck(data):
    """Builds a comprehensive, professionally styled PowerPoint presentation."""
    prs = Presentation()
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)
    
    def add_title_slide(title_text, subtitle_text):
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title, subtitle = slide.shapes.title, slide.placeholders[1]
        title.text = title_text
        subtitle.text = subtitle_text
        
    def add_content_slide(title_text):
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = title_text
        content_placeholder = slide.placeholders[1]
        return slide, content_placeholder.text_frame

    # --- Slide Generation ---
    # Slide 1: Title
    add_title_slide(f"Quarterly Business Review: {data['customer_name']}", f"Q3 2025 Report | Prepared: {datetime.date.today().strftime('%B %d, %Y')}")

    # Slide 2: Agenda
    slide, content = add_content_slide("Agenda")
    topics = [
        "Quarterly Snapshot & Highlights", "Commitment Review: Promises vs. Reality",
        "Challenges & Key Learnings", "Objectives for Next Quarter (OKRs)",
        "Strategic Growth & Product Roadmap", "Commercial Outlook: Forecast & Pipeline",
        "Joint Action Plan & Next Steps"
    ]
    for topic in topics:
        p = content.add_paragraph()
        p.text = topic
        p.level = 0
    
    # Slide 3: Quarterly Snapshot
    slide, _ = add_content_slide("Quarterly Snapshot: Key Metrics")
    for i, (key, value) in enumerate(data['kpis'].items()):
        txBox = slide.shapes.add_textbox(Inches(i*2.5 + 1), Inches(2.5), Inches(2), Inches(2))
        tf = txBox.text_frame
        p_val = tf.add_paragraph()
        p_val.text = str(value)
        p_val.font.bold = True
        p_val.font.size = Pt(44)
        p_val.alignment = PP_ALIGN.CENTER
        p_key = tf.add_paragraph()
        p_key.text = key
        p_key.font.size = Pt(16)
        p_key.alignment = PP_ALIGN.CENTER

    # Slide 4: Commit vs Actual
    slide, _ = add_content_slide("Commitment Review: Promises vs. Reality")
    add_table_to_slide(slide, data['commit_vs_actual'], Inches(1), Inches(2), Inches(14), Inches(4))

    # Slide 5: Challenges & Lessons
    slide, _ = add_content_slide("Challenges & Key Learnings")
    txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(6.5), Inches(5))
    tf = txBox.text_frame
    tf.text = "Challenges Faced This Quarter"
    tf.paragraphs[0].font.bold = True
    for item in data['challenges']:
        p = tf.add_paragraph(); p.text = item; p.level = 1
    
    txBox2 = slide.shapes.add_textbox(Inches(8.5), Inches(2), Inches(6.5), Inches(5))
    tf2 = txBox2.text_frame
    tf2.text = "Key Lessons Learned"
    tf2.paragraphs[0].font.bold = True
    for item in data['learnings']:
        p = tf2.add_paragraph(); p.text = item; p.level = 1

    # Slide 6: Next Quarter OKRs
    slide, _ = add_content_slide("Objectives for Next Quarter (OKRs)")
    add_table_to_slide(slide, data['okrs'], Inches(1), Inches(2), Inches(14), Inches(5))

    # Slide 7: Roadmap
    slide, _ = add_content_slide("Strategic Growth & Product Roadmap")
    for i, (quarter, features) in enumerate(data['roadmap'].items()):
        txBox = slide.shapes.add_textbox(Inches(i*5 + 1), Inches(2.5), Inches(4.5), Inches(4))
        tf = txBox.text_frame
        p_qtr = tf.add_paragraph(); p_qtr.text = quarter; p_qtr.font.bold = True; p_qtr.font.size = Pt(24)
        for feature in features:
            p_feat = tf.add_paragraph(); p_feat.text = feature; p_feat.level = 1

    # Slide 8: Revenue Forecast
    slide, _ = add_content_slide("Commercial Outlook: Revenue Forecast")
    chart_path = create_revenue_chart(data['revenue_forecast'], data['customer_name'])
    slide.shapes.add_picture(chart_path, Inches(2), Inches(1.8), width=Inches(12))
    os.remove(chart_path)

    # Slide 9: Action Plan
    slide, _ = add_content_slide("Joint Action Plan & Owners")
    add_table_to_slide(slide, data['action_plan'], Inches(1), Inches(2), Inches(14), Inches(4))

    # Slide 10: Thank You
    add_title_slide("Thank You", "Q&A and Discussion")

    # --- Save Presentation ---
    output_filename = f"QBR_{data['customer_name'].replace(' ', '_')}_{datetime.date.today()}.pptx"
    prs.save(output_filename)
    return output_filename

# --- 2. FRONTEND UI: PROFESSIONAL STREAMLIT APPLICATION ---

st.set_page_config(
    page_title="AI QBR Deck Generator",
    page_icon="✨",
    layout="wide"
)

# --- Custom CSS for a professional look ---
st.markdown("""
<style>
    .stApp {
        background-color: #f0f2f6;
    }
    .stButton>button {
        background-color: #1976d2;
        color: white;
        border-radius: 5px;
        height: 3em;
        width: 100%;
    }
    .stTextInput>div>div>input {
        border-radius: 5px;
    }
    .st-emotion-cache-16txtl3 {
        padding: 2rem 2rem;
    }
    h1, h2, h3 {
        color: #0d47a1;
    }
    .feature-box {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        text-align: center;
        height: 100%;
    }
    .feature-box h3 {
        color: #1976d2;
    }
    .feature-box p {
        color: #555;
    }
</style>
""", unsafe_allow_html=True)


# --- UI Layout ---
# Header Section
with st.container():
    st.title("🤖 AI-Powered QBR Deck Generator")
    st.markdown("### From raw data to a client-ready presentation in one click.")
    st.markdown("This tool automates the creation of comprehensive Quarterly Business Reviews, saving you hours of manual work. Simply enter your customer's name to generate a professional, data-driven PowerPoint deck.")
    st.divider()

# Main Content Area
col1, col2 = st.columns([1, 2])

# Left Column: Controls
with col1:
    st.header("Settings")
    customer_name = st.text_input("Enter Customer Name", "Global Tech Innovators")
    
    if st.button("🚀 Generate QBR Deck"):
        if customer_name:
            with st.spinner('Crafting your presentation... This may take a moment.'):
                try:
                    # Simulate AI process with steps
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    status_text.info("Step 1/4: Aggregating customer data...")
                    time.sleep(1.5)
                    progress_bar.progress(25)

                    status_text.info("Step 2/4: Generating AI-driven insights...")
                    time.sleep(1.5)
                    progress_bar.progress(50)
                    
                    status_text.info("Step 3/4: Creating data visualizations...")
                    time.sleep(1.5)
                    progress_bar.progress(75)

                    status_text.info("Step 4/4: Assembling professional PowerPoint deck...")
                    enhanced_data = get_enhanced_mock_data(customer_name)
                    final_deck_path = create_professional_qbr_deck(enhanced_data)
                    time.sleep(1)
                    progress_bar.progress(100)
                    
                    status_text.success(f"🎉 Your QBR deck for **{customer_name}** is ready!")
                    
                    with open(final_deck_path, "rb") as file:
                        st.download_button(
                            label="⬇️ Download Presentation",
                            data=file,
                            file_name=final_deck_path,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    os.remove(final_deck_path) # Clean up file from server
                except Exception as e:
                    st.error(f"An error occurred: {e}")
        else:
            st.warning("Please enter a customer name.")

# Right Column: Feature Highlights
with col2:
    st.header("What's Inside Your AI-Generated Deck?")
    
    f_col1, f_col2, f_col3 = st.columns(3)
    
    with f_col1:
        st.markdown("""
        <div class="feature-box">
            <h3>📊 Data-Driven Insights</h3>
            <p>Automatically synthesizes KPIs, performance metrics, and revenue forecasts into clear, actionable slides.</p>
        </div>
        """, unsafe_allow_html=True)

    with f_col2:
        st.markdown("""
        <div class="feature-box">
            <h3>📈 Strategic Roadmaps</h3>
            <p>Includes forward-looking content like next-quarter OKRs, growth initiatives, and product roadmaps.</p>
        </div>
        """, unsafe_allow_html=True)

    with f_col3:
        st.markdown("""
        <div class="feature-box">
            <h3>✅ Action-Oriented Plans</h3>
            <p>Generates clear action plans with owners and due dates to ensure accountability and follow-through.</p>
        </div>
        """, unsafe_allow_html=True)
        
    st.image("https://placehold.co/900x400/0d47a1/ffffff?text=Sample+Chart+Visualization",
             caption="Visualizations are created automatically to highlight key trends.",
             use_column_width=True)

st.sidebar.info("This is a Proof-of-Concept. All data is realistically simulated for demonstration.")
