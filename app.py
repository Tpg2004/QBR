# app.py
# The Streamlit UI for our QBR Generator.

import streamlit as st
from qbr_generator import generate_qbr_for_customer # Import our backend logic
import time
import os

# --- Page Configuration ---
st.set_page_config(
    page_title="AI QBR Deck Generator",
    page_icon="✨",
    layout="centered"
)

# --- UI Elements ---
st.title("🤖 AI-Powered QBR Deck Generator")
st.markdown("Enter a customer name and let AI create a complete, data-driven QBR presentation in seconds.")

# 

st.sidebar.header("Controls")
customer_name = st.sidebar.text_input("Enter Customer Name", "Global Tech Corp")

if st.sidebar.button("🚀 Generate QBR Deck"):
    if customer_name:
        with st.spinner('Generating your presentation... This might take a moment.'):
            st.info("Step 1: Fetching and analyzing customer data...")
            time.sleep(2) # Simulate work
            st.info("Step 2: Generating insights and visualizations...")
            time.sleep(2) # Simulate work
            st.info("Step 3: Assembling the PowerPoint deck...")
            
            # Call the backend function
            try:
                final_deck_path = generate_qbr_for_customer(customer_name)
                
                st.success(f"🎉 Your QBR deck for **{customer_name}** is ready!")
                
                # Provide a download button
                with open(final_deck_path, "rb") as file:
                    st.download_button(
                        label="⬇️ Download Presentation",
                        data=file,
                        file_name=final_deck_path,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                # Clean up the generated file after download is prepared
                # os.remove(final_deck_path) # Optional: uncomment to delete the file from the server
            except Exception as e:
                st.error(f"An error occurred: {e}")
    else:
        st.warning("Please enter a customer name.")

st.sidebar.markdown("---")
st.sidebar.info("This is a Proof-of-Concept app. Data is simulated for demonstration purposes.")
