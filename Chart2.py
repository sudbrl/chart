import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from fpdf import FPDF
import streamlit as st
import tempfile
import os
import time
from supabase import create_client, Client

# --- Supabase Initialization (Secure) ---
@st.cache_resource
def init_supabase():
    """Initialize Supabase client using Streamlit Secrets."""
    if "supabase" not in st.secrets:
        st.error("⚠️ Configuration Error: Secrets missing.")
        st.stop()

    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]

    if not url or not key:
        st.error("⚠️ Configuration Error: Invalid URL or Key.")
        st.stop()

    try:
        return create_client(url, key)
    except Exception:
        st.error("⚠️ Connection Error: Could not connect to the database.")
        return None

supabase = init_supabase()

# --- Session State Management ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'user' not in st.session_state:
    st.session_state.user = None
if 'last_attempt_time' not in st.session_state:
    st.session_state.last_attempt_time = 0

# --- Authentication Logic ---
def login_page():
    st.title("User Login")
    
    with st.form("login_form"):
        st.subheader("Sign In")
        email = st.text_input("Email")
        password = st.text_input("Password", type="password")
        
        submit_button = st.form_submit_button("Log In")

        if submit_button:
            # 1. Simple Rate Limit Check
            current_time = time.time()
            if current_time - st.session_state.last_attempt_time < 2:
                st.warning("⏳ Too fast! Please wait a moment.")
                return
            st.session_state.last_attempt_time = current_time

            # 2. Supabase Auth
            try:
                response = supabase.auth.sign_in_with_password({"email": email, "password": password})
                st.session_state.authenticated = True
                st.session_state.user = response.user
                
                st.success("Login successful!")
                st.rerun()
            except Exception:
                # UPDATED: Removed {str(e)} to hide the "Debug Hint"
                st.error("❌ **Login Failed**") 
                st.warning("Please check your Email and Password.")

# --- Main Application Logic ---
def main_app():
    with st.sidebar:
        if st.session_state.user:
            st.write(f"👤 **{st.session_state.user.email}**")
        
        if st.button("Log Out"):
            if supabase:
                supabase.auth.sign_out()
            st.session_state.authenticated = False
            st.session_state.user = None
            st.rerun()

    # --- Original Loan Report Logic ---
    st.title("Comprehensive Loan Report Generator")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file:
        try:
            file_bytes = BytesIO(uploaded_file.read())

            # Branch Sheet Analysis
            branch_sheet = 'Branch'
            try:
                branch_df = pd.read_excel(file_bytes, sheet_name=branch_sheet)
            except ValueError:
                st.error(f"Sheet '{branch_sheet}' not found in the uploaded file.")
                return

            required_branch_cols = {'BranchName', 'Change'}
            if not required_branch_cols.issubset(branch_df.columns):
                st.error(f"Branch sheet must contain columns: {required_branch_cols}")
            else:
                top_5_increasing = branch_df.nlargest(5, 'Change')
                top_5_declining = branch_df.nsmallest(5, 'Change')

                combined_branches = pd.concat([
                    top_5_increasing[['BranchName', 'Change']],
                    top_5_declining[['BranchName', 'Change']]
                ])
                combined_branches['Change (Crore)'] = combined_branches['Change'] / 1e7

                fig_branch, ax_branch = plt.subplots(figsize=(12, 8))
                colors_branch = ['green' if x > 0 else 'red' for x in combined_branches['Change']]
                bars_branch = ax_branch.bar(combined_branches['BranchName'], combined_branches['Change'],
                                            color=colors_branch, edgecolor='black', alpha=0.7)
                ax_branch.set_title('Top 5 Increasing and Declining Branches (in Crores)', fontsize=16)
                ax_branch.set_xlabel('Branch Name', fontsize=14)
                ax_branch.set_ylabel('Loan Change (in currency)', fontsize=14)
                ax_branch.set_xticks(range(len(combined_branches['BranchName'])))
                ax_branch.set_xticklabels(combined_branches['BranchName'], rotation=45, ha='right')
                
                for bar, change_crore in zip(bars_branch, combined_branches['Change (Crore)']):
                    yval = bar.get_height()
                    va_pos = 'bottom' if yval > 0 else 'top'
                    label_y = yval + (yval * 0.05) if yval > 0 else yval - (yval * 0.05)
                    
                    ax_branch.text(bar.get_x() + bar.get_width() / 2, yval, f'Rs.{abs(change_crore):.2f} Cr',
                                   ha='center', va=va_pos, fontsize=10)
                
                st.subheader("Branch Loan Change Analysis")
                st.pyplot(fig_branch)

            file_bytes.seek(0)

            # Compare Sheet Analysis
            compare_sheet = 'Compare'
            try:
                df_compare = pd.read_excel(file_bytes, sheet_name=compare_sheet)
            except ValueError:
                st.error(f"Sheet '{compare_sheet}' not found in the uploaded file.")
                return

            required_compare_cols = {'index', 'Change'}
            if not required_compare_cols.issubset(df_compare.columns):
                st.error(f"Compare sheet must contain columns: {required_compare_cols}")
            else:
                df_compare['Change_Crore_NRs'] = df_compare['Change'] / 1e7
                df_compare_sorted = df_compare.sort_values('Change_Crore_NRs', ascending=False)

                colors_compare = [
                    'green' if i < 3 else 'blue' if i < len(df_compare_sorted) - 3 else 'red'
                    for i in range(len(df_compare_sorted))
                ]

                fig_compare, ax_compare = plt.subplots(figsize=(12, 8))
                ax_compare.barh(df_compare_sorted['index'], df_compare_sorted['Change_Crore_NRs'], color=colors_compare)
                ax_compare.set_xlabel('Change in Balance (Crore NRs)')
                ax_compare.set_ylabel('Loan Type')
                ax_compare.set_title('Change in Loan Balances Across Loan Types (in Crore NRs)')
                ax_compare.grid(axis='x', linestyle='--', alpha=0.7)
                st.subheader("Loan Type Balance Change Analysis")
                st.pyplot(fig_compare)

            if st.button("Generate PDF Report"):
                with st.spinner("Generating PDF..."):
                    pdf = FPDF()
                    pdf.set_auto_page_break(auto=True, margin=15)
                    pdf.add_page()
                    pdf.set_font("Arial", 'B', 16)
                    pdf.cell(200, 10, txt="Comprehensive Loan Report", ln=True, align='C')
                    pdf.ln(10)

                    if 'fig_branch' in locals():
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                            fig_branch.savefig(tmpfile.name, format="png", bbox_inches='tight')
                            pdf.image(tmpfile.name, x=10, y=30, w=190)
                            os.unlink(tmpfile.name)

                    pdf.add_page()
                    pdf.cell(200, 10, txt="Loan Type Balance Change", ln=True, align='C')
                    pdf.ln(10)

                    if 'fig_compare' in locals():
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                            fig_compare.savefig(tmpfile.name, format="png", bbox_inches='tight')
                            pdf.image(tmpfile.name, x=10, y=30, w=190)
                            os.unlink(tmpfile.name)

                    pdf_buffer = BytesIO()
                    pdf_output = pdf.output(dest='S').encode('latin1')
                    pdf_buffer.write(pdf_output)
                    pdf_buffer.seek(0)

                    st.download_button(label="Download Comprehensive PDF Report",
                                       data=pdf_buffer,
                                       file_name="Comprehensive_Loan_Report.pdf",
                                       mime="application/pdf")

        except Exception as e:
            st.error(f"Error processing file: {e}")

# --- App Flow Control ---
if st.session_state.authenticated:
    main_app()
else:
    login_page()
