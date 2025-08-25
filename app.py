# app.py
# A professional, single-file Streamlit application to generate comprehensive, AI-powered QBR decks.
# Version 9: Added a professional login page.

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import datetime
import os
import time

# --- 1. BACKEND LOGIC: ENHANCED DATA & PRESENTATION GENERATION ---

# --- Color Palette & Helper ---
# A more sophisticated corporate color palette
PALETTE = {
    "navy": RGBColor(15, 32, 62),
    "blue": RGBColor(0, 122, 255),
    "teal": RGBColor(8, 143, 143),
    "gray": RGBColor(128, 128, 128),
    "light_gray": RGBColor(240, 240, 240),
    "white": RGBColor(255, 255, 255)
}

def rgb_to_hex(rgb_color_obj):
    """Converts a python-pptx RGBColor object to a hex string for matplotlib."""
    r, g, b = rgb_color_obj; return f"#{r:02x}{g:02x}{b:02x}"

def get_enhanced_mock_data(customer_name):
    """Generates a rich, multi-faceted dataset for a comprehensive QBR."""
    np.random.seed(hash(customer_name) % (2**32 - 1))
    kpis = {
        "Account Health": f"{np.random.randint(75, 98)}/100", "NPS Score": f"{np.random.randint(30, 65)}",
        "Adoption Rate": f"{np.random.randint(60, 95)}%", "Active Users": f"{np.random.randint(150, 500)}",
    }
    commit_vs_actual = pd.DataFrame({
        'Metric': ['Feature Delivery', 'Uptime SLA', 'Avg. Ticket Response'], 'Commitment': ['5 New Features', '99.9% Uptime', '< 8 Hours'],
        'Actual': ['6 New Features', f"{99.9 + np.random.uniform(0, 0.09):.2f}%", f"{np.random.uniform(6, 7.9):.1f} Hours"], 'Status': ['Exceeded', 'Met', 'Met']
    })
    challenges = ["Slower than anticipated onboarding for the new analytics module.", "Integration with legacy CRM required custom development.", "User adoption in the finance department is lagging."]
    learnings = ["Dedicated onboarding webinars significantly boost initial adoption.", "Pre-sales technical discovery for legacy systems is crucial.", "Targeted training and identifying team champions accelerate adoption."]
    okrs = pd.DataFrame({
        'Objective': ["Enhance Collaboration", "Improve Data-Driven Decisions"],
        'Key Result': ["Increase shared dashboard usage by 20%", "Train 5 team leads on advanced reporting"], 'Status': ['Not Started', 'Not Started']
    })
    roadmap = {"Next Quarter": ["AI-Powered Insights Engine", "Mobile App V2 Launch"], "Following Quarter": ["Advanced API Access", "Integration Marketplace"]}
    revenue_forecast = pd.DataFrame({'Month': pd.to_datetime([f'2025-{i}-01' for i in range(10, 13)]), 'Forecasted Revenue ($K)': [50 + i*5 + np.random.randint(-5, 5) for i in range(3)]})
    action_plan = pd.DataFrame({
        'Action Item': ["Schedule marketing team onboarding", "Finalize BI tool integration specs"], 'Owner': ["John (CSM)", "Jane (Customer IT)"],
        'Due Date': [(datetime.date.today() + datetime.timedelta(days=d)).strftime('%Y-%m-%d') for d in [14, 30]], 'Status': ['Not Started', 'Not Started']
    })
    return locals()

def create_revenue_chart(revenue_df, output_path="revenue_chart.png"):
    """Creates a visually improved bar chart for the revenue forecast."""
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(8, 4))
    sns.barplot(x=revenue_df['Month'].dt.strftime('%b'), y='Forecasted Revenue ($K)', data=revenue_df, color=rgb_to_hex(PALETTE["blue"]), ax=ax)
    ax.set_title('Next Quarter Revenue Forecast', fontsize=14, weight='bold', color=rgb_to_hex(PALETTE["navy"]))
    ax.set_xlabel(''); ax.set_ylabel('Forecasted Revenue ($K)', fontsize=10)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    sns.despine(left=True, bottom=True)
    plt.tight_layout(); plt.savefig(output_path, dpi=300, transparent=True)
    return output_path

def add_master_elements(slide, customer_name):
    """Adds consistent footer and design elements to each slide."""
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(8.5), Inches(15), Inches(0.4))
    p = footer.text_frame.paragraphs[0]
    p.text = f"QBR for {customer_name}  |  {datetime.date.today().strftime('%B %Y')}"
    p.font.size = Pt(10); p.font.color.rgb = PALETTE["gray"]
    
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.5), Inches(16), Inches(0.05))
    accent.fill.solid(); accent.fill.fore_color.rgb = PALETTE["blue"]
    accent.line.fill.background()

def add_table_to_slide(slide, df, x, y, cx, cy):
    """Adds a professionally styled table to a slide with zebra striping."""
    shape = slide.shapes.add_table(df.shape[0] + 1, df.shape[1], x, y, cx, cy)
    table = shape.table
    for i in range(df.shape[1]): table.columns[i].width = int(cx / df.shape[1])
    for i, col_name in enumerate(df.columns):
        cell = table.cell(0, i); cell.text = col_name; cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]; p.font.bold = True; p.font.color.rgb = PALETTE["white"]
        cell.fill.solid(); cell.fill.fore_color.rgb = PALETTE["navy"]
    for i, row in df.iterrows():
        for j, value in enumerate(row):
            cell = table.cell(i + 1, j); cell.text = str(value); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i % 2 == 0: # Zebra striping for readability
                cell.fill.solid(); cell.fill.fore_color.rgb = PALETTE["light_gray"]

def create_professional_qbr_deck(data):
    """Builds the final, professionally styled PowerPoint presentation."""
    prs = Presentation(); prs.slide_width = Inches(16); prs.slide_height = Inches(9)
    def add_title_slide(title_text, subtitle_text):
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title_text; slide.placeholders[1].text = subtitle_text
        return slide
    def add_content_slide(title_text):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = title_text; add_master_elements(slide, data['customer_name'])
        return slide, slide.placeholders[1]

    add_title_slide(f"Quarterly Business Review: {data['customer_name']}", f"Q3 2025 Report")
    slide, content_placeholder = add_content_slide("Agenda"); content = content_placeholder.text_frame; content.clear()
    topics = ["Quarterly Snapshot", "Commitment Review", "Challenges & Learnings", "Next Quarter OKRs", "Product Roadmap", "Commercial Outlook", "Action Plan"]
    for topic in topics: p = content.add_paragraph(); p.text = topic; p.level = 0; p.space_after = Pt(18)

    slide, _ = add_content_slide("Quarterly Snapshot: Key Metrics")
    num_kpis = len(data['kpis']); box_width = 3.0; gap = 0.8; total_width = num_kpis * box_width + (num_kpis - 1) * gap
    start_x = (16 - total_width) / 2
    for i, (key, value) in enumerate(data['kpis'].items()):
        left = Inches(start_x + i * (box_width + gap))
        txBox = slide.shapes.add_textbox(left, Inches(2.5), Inches(box_width), Inches(2)); tf = txBox.text_frame
        p_val = tf.add_paragraph(); p_val.text = str(value); p_val.font.bold = True; p_val.font.size = Pt(48); p_val.alignment = PP_ALIGN.CENTER
        p_key = tf.add_paragraph(); p_key.text = key; p_key.font.size = Pt(18); p_key.alignment = PP_ALIGN.CENTER

    slide, _ = add_content_slide("Commitment Review: Promises vs. Reality"); add_table_to_slide(slide, data['commit_vs_actual'], Inches(1.5), Inches(2.5), Inches(13), Inches(4))
    
    slide, _ = add_content_slide("Challenges & Key Learnings")
    box_w, gap_cl = 6.5, 1.0; left1 = (16 - (2 * box_w + gap_cl)) / 2; left2 = left1 + box_w + gap_cl
    txBox1 = slide.shapes.add_textbox(Inches(left1), Inches(2.5), Inches(box_w), Inches(5)); tf1 = txBox1.text_frame
    tf1.text = "Challenges Faced"; tf1.paragraphs[0].font.bold = True; tf1.paragraphs[0].font.size = Pt(24)
    for item in data['challenges']: p = tf1.add_paragraph(); p.text = f"• {item}"; p.space_after = Pt(8)
    txBox2 = slide.shapes.add_textbox(Inches(left2), Inches(2.5), Inches(box_w), Inches(5)); tf2 = txBox2.text_frame
    tf2.text = "Key Lessons Learned"; tf2.paragraphs[0].font.bold = True; tf2.paragraphs[0].font.size = Pt(24)
    for item in data['learnings']: p = tf2.add_paragraph(); p.text = f"• {item}"; p.space_after = Pt(8)

    slide, _ = add_content_slide("Objectives for Next Quarter (OKRs)"); add_table_to_slide(slide, data['okrs'], Inches(1.5), Inches(2.5), Inches(13), Inches(3))
    
    slide, _ = add_content_slide("Strategic Growth & Product Roadmap")
    num_r, box_w_r, gap_r = len(data['roadmap']), 6.0, 2.0; total_w_r = num_r * box_w_r + (num_r - 1) * gap_r
    start_x_r = (16 - total_w_r) / 2
    for i, (quarter, features) in enumerate(data['roadmap'].items()):
        left = Inches(start_x_r + i * (box_w_r + gap_r))
        txBox = slide.shapes.add_textbox(left, Inches(2.5), Inches(box_w_r), Inches(5)); tf = txBox.text_frame
        p_qtr = tf.add_paragraph(); p_qtr.text = quarter; p_qtr.font.bold = True; p_qtr.font.size = Pt(24)
        for feature in features: p_feat = tf.add_paragraph(); p_feat.text = f"• {feature}"; p_feat.space_after = Pt(8)

    slide, _ = add_content_slide("Commercial Outlook: Revenue Forecast"); chart_path = create_revenue_chart(data['revenue_forecast'])
    slide.shapes.add_picture(chart_path, Inches(3), Inches(2.0), width=Inches(10)); os.remove(chart_path)
    
    slide, _ = add_content_slide("Joint Action Plan & Owners"); add_table_to_slide(slide, data['action_plan'], Inches(1.5), Inches(2.5), Inches(13), Inches(3))
    
    add_title_slide("Thank You", "Q&A and Discussion")
    output_filename = f"QBR_{data['customer_name'].replace(' ', '_')}_{datetime.date.today()}.pptx"; prs.save(output_filename); return output_filename

# --- 2. FRONTEND UI: MAIN APPLICATION & LOGIN PAGE ---

def main_app():
    """This function contains the main QBR generator application UI."""
    st.markdown("<div class='header'><h1>AI QBR Deck Generator</h1><p>Instantly create stunning, data-driven presentations that impress.</p></div>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 1.5])
    with col1:
        with st.container():
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.subheader("Start Here")
            customer_name = st.text_input("Enter Customer Name", "Innovate Corp", label_visibility="collapsed")
            if st.button("🚀 Generate Presentation"):
                if customer_name:
                    with st.spinner('Analyzing data and building your deck...'):
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
            st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        with st.container():
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.subheader("How It Works")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown("<div class='feature-card'><div class='icon'>📊</div><h3>Data Synthesis</h3><p>Aggregates key metrics from all your sources.</p></div>", unsafe_allow_html=True)
            with c2:
                st.markdown("<div class='feature-card'><div class='icon'>🤖</div><h3>AI Narration</h3><p>Generates summaries and actionable insights.</p></div>", unsafe_allow_html=True)
            with c3:
                st.markdown("<div class='feature-card'><div class='icon'>🎨</div><h3>Design Automation</h3><p>Builds a professionally designed presentation.</p></div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    st.sidebar.info("This is a Proof-of-Concept. All data is realistically simulated for demonstration.")
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.rerun()

def login_page():
    """This function displays the login UI."""
    st.markdown("<div class='header'><h1>Welcome Back</h1><p>Please log in to continue.</p></div>", unsafe_allow_html=True)
    
    with st.container():
        col1, col2, col3 = st.columns([1,1,1])
        with col2:
            with st.container():
                st.markdown("<div class='card login-card'>", unsafe_allow_html=True)
                username = st.text_input("Username", placeholder="Enter your username")
                password = st.text_input("Password", type="password", placeholder="Enter your password")
                if st.button("Login"):
                    # For this demo, any non-empty credentials will work.
                    if username and password:
                        st.session_state.logged_in = True
                        st.rerun()
                    else:
                        st.error("Please enter both username and password.")
                st.markdown("</div>", unsafe_allow_html=True)


# --- MAIN SCRIPT EXECUTION ---
st.set_page_config(page_title="AI QBR Deck Generator", page_icon="✨", layout="wide")

# --- SHARED STYLES FOR BOTH PAGES ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    body { font-family: 'Inter', sans-serif; }
    .stApp { background: #f0f2f6; }
    .main-container { padding: 2rem; }
    .header { text-align: center; margin-bottom: 2rem; }
    .header h1 {
        font-size: 3rem; font-weight: 700; color: #0f203e;
        background: -webkit-linear-gradient(45deg, #007bff, #0f203e);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .header p { font-size: 1.2rem; color: #555; }
    .card {
        background: white; border-radius: 15px; padding: 25px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.07); transition: all 0.3s ease;
    }
    .card:hover { transform: translateY(-5px); box-shadow: 0 15px 30px rgba(0,0,0,0.1); }
    .login-card { padding: 35px; }
    .stButton>button {
        background-image: linear-gradient(to right, #007bff 0%, #0056b3 100%);
        color: white; border-radius: 10px; transition: 0.5s; background-size: 200% auto;
        font-weight: 600; border: none; height: 3em; width: 100%;
    }
    .stButton>button:hover { background-position: right center; }
    .stDownloadButton>button { background-image: linear-gradient(to right, #28a745, #218838); }
    .feature-card { text-align: center; padding: 1.5rem; }
    .feature-card h3 { color: #0f203e; font-weight: 600; }
    .feature-card .icon { font-size: 3rem; margin-bottom: 1rem; color: #007bff; }
</style>
""", unsafe_allow_html=True)

# --- CONDITIONAL PAGE RENDERING ---
# Initialize session state for login
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

# Check login status and display the appropriate page
if st.session_state.logged_in:
    main_app()
else:
    login_page()
