import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import base64
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import json
import os
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import Flow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    GOOGLE_APIS_AVAILABLE = True
except ImportError:
    GOOGLE_APIS_AVAILABLE = False

def add_custom_css():
    """Add custom CSS for Typeform-like styling"""
    st.markdown("""
    <style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* Global Styles */
    .stApp {
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: #2c3e50;
        padding-top: 0 !important;
    }

    /* Remove default Streamlit padding */
    .main .block-container {
        padding-top: 0 !important;
        padding-bottom: 0 !important;
    }

    /* Hide Streamlit elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Main container */
    .main-container {
        max-width: 800px;
        margin: 0 auto;
        padding: 2rem;
        background: rgba(255, 255, 255, 0.95);
        border-radius: 20px;
        box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        backdrop-filter: blur(10px);
        margin-top: 0;
        margin-bottom: 2rem;
    }

    /* Hero Section */
    .hero-section {
        text-align: center;
        padding: 2rem 0 2rem 0;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        margin: -2rem -2rem 3rem -2rem;
        border-radius: 20px 20px 0 0;
        color: white;
    }

    .hero-title {
        font-size: 3.5rem;
        font-weight: 700;
        margin-bottom: 1rem;
        background: linear-gradient(45deg, #fff, #f8f9fa);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    .hero-subtitle {
        font-size: 1.3rem;
        font-weight: 400;
        opacity: 0.9;
        margin-bottom: 0;
        max-width: 600px;
        margin-left: auto;
        margin-right: auto;
        line-height: 1.6;
    }

    /* Upload Section */
    .upload-section {
        background: #f8f9ff;
        border-radius: 16px;
        padding: 3rem 2rem;
        text-align: center;
        margin: 2rem 0;
        border: 2px dashed #667eea;
        transition: all 0.3s ease;
    }

    .upload-section:hover {
        border-color: #764ba2;
        background: #f5f7ff;
        transform: translateY(-2px);
    }

    .upload-icon {
        font-size: 3rem;
        margin-bottom: 1rem;
        color: #667eea;
    }

    .upload-title {
        font-size: 1.5rem;
        font-weight: 600;
        color: #2c3e50;
        margin-bottom: 0.5rem;
    }

    .upload-description {
        color: #64748b;
        font-size: 1rem;
        margin-bottom: 2rem;
    }

    /* File uploader styling */
    .stFileUploader > div {
        border: none !important;
        background: transparent !important;
    }

    .stFileUploader label {
        font-weight: 600 !important;
        color: #2c3e50 !important;
        font-size: 1rem !important;
    }

    /* Success message */
    .success-card {
        background: linear-gradient(135deg, #4ade80 0%, #22c55e 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        text-align: center;
        box-shadow: 0 4px 12px rgba(34, 197, 94, 0.3);
    }

    /* Info card */
    .info-card {
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        font-weight: 500;
    }

    /* Error card */
    .error-card {
        background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        margin: 1rem 0;
        text-align: center;
        box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3);
    }

    /* Button styling */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        padding: 0.75rem 2rem !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3) !important;
    }

    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 24px rgba(102, 126, 234, 0.4) !important;
    }

    /* Expander styling */
    .streamlit-expanderHeader {
        background: #f1f5f9 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }

    /* Download section */
    .download-section {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border-radius: 16px;
        padding: 2rem;
        margin: 2rem 0;
        border: 1px solid #e2e8f0;
    }

    /* Progress indicators */
    .step-indicator {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 2rem 0;
    }

    .step {
        width: 40px;
        height: 40px;
        border-radius: 50%;
        background: #e2e8f0;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 1rem;
        font-weight: 600;
        color: #64748b;
        position: relative;
    }

    .step.active {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }

    .step.completed {
        background: #22c55e;
        color: white;
    }

    .step::after {
        content: '';
        position: absolute;
        top: 50%;
        left: 100%;
        width: 60px;
        height: 2px;
        background: #e2e8f0;
        z-index: -1;
    }

    .step:last-child::after {
        display: none;
    }

    .step.completed::after {
        background: #22c55e;
    }

    /* Cards */
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        border: 1px solid #e2e8f0;
        margin: 1rem 0;
    }

    /* Animations */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(30px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }

    .fade-in-up {
        animation: fadeInUp 0.6s ease-out;
    }

    /* Data preview */
    .stDataFrame {
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1) !important;
    }

    /* Charts */
    .stPlotlyChart {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        margin: 1rem 0;
    }

    /* Navigation tabs */
    .nav-tabs {
        display: flex;
        justify-content: center;
        background: rgba(255, 255, 255, 0.9);
        border-radius: 16px;
        padding: 0.5rem;
        margin: 2rem 0;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        backdrop-filter: blur(10px);
    }

    .nav-tab {
        padding: 0.75rem 2rem;
        border-radius: 12px;
        margin: 0 0.25rem;
        font-weight: 600;
        font-size: 1rem;
        cursor: pointer;
        transition: all 0.3s ease;
        color: #64748b;
        background: transparent;
        border: none;
        min-width: 120px;
    }

    .nav-tab:hover {
        background: rgba(102, 126, 234, 0.1);
        color: #667eea;
        transform: translateY(-1px);
    }

    .nav-tab.active {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
    }

    /* Filter section */
    .filter-section {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border-radius: 16px;
        padding: 1.5rem;
        margin: 1rem 0;
        border: 1px solid #e2e8f0;
    }

    .filter-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #2c3e50;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    /* Section headers */
    .section-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 16px;
        margin: 2rem 0 1rem 0;
        text-align: center;
    }

    .section-title {
        font-size: 1.8rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }

    .section-subtitle {
        font-size: 1rem;
        opacity: 0.9;
    }

    /* Stats cards */
    .stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin: 1.5rem 0;
    }

    .stat-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        border: 1px solid #e2e8f0;
        transition: transform 0.3s ease;
    }

    .stat-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 16px rgba(0,0,0,0.15);
    }

    .stat-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #667eea;
        margin-bottom: 0.5rem;
    }

    .stat-label {
        color: #64748b;
        font-weight: 500;
        text-transform: uppercase;
        font-size: 0.85rem;
        letter-spacing: 0.5px;
    }

    /* Team filter */
    .stSelectbox > div > div {
        background: white !important;
        border-radius: 8px !important;
        border: 2px solid #e2e8f0 !important;
        transition: border-color 0.3s ease !important;
    }

    .stSelectbox > div > div:focus-within {
        border-color: #667eea !important;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1) !important;
    }

    /* Radio button navigation styling */
    .stRadio > div {
        background: rgba(255, 255, 255, 0.9) !important;
        border-radius: 16px !important;
        padding: 0.5rem !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1) !important;
        backdrop-filter: blur(10px) !important;
        margin: 1rem 0 !important;
    }

    .stRadio > div > div {
        flex-direction: column !important;
        gap: 0.5rem !important;
        align-items: stretch !important;
    }

    .stRadio > div > div > label {
        background: white !important;
        border: 2px solid #e2e8f0 !important;
        border-radius: 12px !important;
        padding: 1rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        color: #2c3e50 !important;
        transition: all 0.3s ease !important;
        cursor: pointer !important;
        text-align: left !important;
        margin: 0 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05) !important;
    }

    .stRadio > div > div > label:hover {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%) !important;
        border-color: #667eea !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.15) !important;
    }

    .stRadio > div > div > label[data-baseweb="radio"] > div:first-child {
        display: none !important;
    }

    .stRadio > div > div > label[aria-checked="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        border-color: #667eea !important;
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4) !important;
        transform: translateY(-2px) !important;
    }

    /* Style for the default "Select a team..." option */
    .stRadio > div > div > label:first-child {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%) !important;
        color: #64748b !important;
        font-style: italic !important;
        border: 2px dashed #cbd5e1 !important;
    }

    .stRadio > div > div > label:first-child:hover {
        background: linear-gradient(135deg, #f1f5f9 0%, #e2e8f0 100%) !important;
        border-color: #94a3b8 !important;
    }

    /* Hide radio circles */
    .stRadio [role="radiogroup"] label > div:first-child {
        display: none !important;
    }

    /* Team selection cards */
    .team-selection-card {
        background: white;
        border: 2px solid #e2e8f0;
        border-radius: 16px;
        padding: 2rem;
        margin: 1rem 0;
        text-align: center;
        transition: all 0.3s ease;
        cursor: pointer;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }

    .team-selection-card:hover {
        border-color: #667eea;
        transform: translateY(-4px);
        box-shadow: 0 8px 24px rgba(102, 126, 234, 0.15);
    }

    .team-selection-card.selected {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-color: #667eea;
        box-shadow: 0 8px 24px rgba(102, 126, 234, 0.3);
    }

    .team-icon {
        font-size: 3rem;
        margin-bottom: 1rem;
    }

    .team-name {
        font-size: 1.3rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }

    .team-description {
        font-size: 0.9rem;
        opacity: 0.8;
    }

    /* Welcome section */
    .welcome-section {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border-radius: 16px;
        padding: 3rem 2rem;
        text-align: center;
        margin: 2rem 0;
        border: 1px solid #e2e8f0;
    }

    .welcome-title {
        font-size: 2rem;
        font-weight: 700;
        color: #2c3e50;
        margin-bottom: 1rem;
    }

    .welcome-subtitle {
        color: #64748b;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }

    /* Enhanced tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem !important;
        background: rgba(255, 255, 255, 0.9) !important;
        border-radius: 16px !important;
        padding: 0.5rem !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1) !important;
        backdrop-filter: blur(10px) !important;
        margin: 1rem 0 !important;
    }

    .stTabs [data-baseweb="tab"] {
        background: transparent !important;
        color: #64748b !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        border-radius: 12px !important;
        padding: 0.75rem 1.5rem !important;
        border: none !important;
        transition: all 0.3s ease !important;
        margin: 0 !important;
    }

    .stTabs [data-baseweb="tab"]:hover {
        background: rgba(102, 126, 234, 0.1) !important;
        color: #667eea !important;
        transform: translateY(-1px) !important;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3) !important;
    }

    .stTabs [data-baseweb="tab-panel"] {
        background: rgba(255, 255, 255, 0.95) !important;
        border-radius: 16px !important;
        padding: 2rem !important;
        margin-top: 1rem !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08) !important;
        backdrop-filter: blur(10px) !important;
    }

    /* Enhanced section headers within tabs */
    .tab-section-header {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border-radius: 12px;
        padding: 1.5rem;
        margin: 1rem 0 2rem 0;
        text-align: center;
        border-left: 4px solid #667eea;
    }

    .tab-section-title {
        font-size: 1.3rem;
        font-weight: 700;
        color: #2c3e50;
        margin-bottom: 0.5rem;
    }

    .tab-section-subtitle {
        color: #64748b;
        font-size: 0.95rem;
    }

    /* Improved info cards within tabs */
    .tab-info-card {
        background: linear-gradient(135deg, #fff 0%, #f8fafc 100%);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }

    .tab-info-card.warning {
        border-left: 4px solid #f59e0b;
        background: linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%);
    }

    .tab-info-card.info {
        border-left: 4px solid #3b82f6;
        background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
    }

    .tab-info-card.success {
        border-left: 4px solid #10b981;
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
    }
    </style>
    """, unsafe_allow_html=True)

def create_step_indicator(current_step):
    """Create a step indicator similar to Typeform"""
    steps = [
        {"number": "1", "title": "Upload", "completed": current_step >= 1},
        {"number": "2", "title": "Analyze", "completed": current_step >= 2},
        {"number": "3", "title": "Download", "completed": current_step >= 3}
    ]

    step_html = '<div class="step-indicator">'
    for i, step in enumerate(steps):
        if step["completed"]:
            step_class = "step completed" if current_step > i + 1 else "step active"
        else:
            step_class = "step"

        step_html += f'<div class="{step_class}">{step["number"]}</div>'

    step_html += '</div>'
    st.markdown(step_html, unsafe_allow_html=True)

def categorize_file(filename):
    """Categorize file based on filename patterns"""
    filename_lower = filename.lower()

    # Extract team name with specific team mappings
    team = None
    team_mappings = {
        'andrew': "Andrew's Team",
        'build': 'Build Team',
        'people': 'People and Marketing Team',
        'marketing': 'People and Marketing Team',
        'hr': 'People and Marketing Team',
        'operations': 'Finance and Operations Team',
        'finance': 'Finance and Operations Team'
    }

    for keyword, team_name in team_mappings.items():
        if keyword in filename_lower:
            team = team_name
            break

    # Determine section/category
    section = None
    if any(word in filename_lower for word in ['comment', 'feedback', 'response']):
        section = 'Comments'
    elif any(word in filename_lower for word in ['theme', 'topic', 'category']):
        section = 'Themes'
    elif any(word in filename_lower for word in ['question', 'survey', 'form']):
        section = 'Questions'
    else:
        # Default to general if no specific category found
        section = 'General'

    return {
        'section': section,
        'team': team,
        'filename': filename
    }

def load_and_process_multiple_files(uploaded_files):
    """Load and process multiple survey data files"""
    # Team-first structure
    team_data = {
        "Andrew's Team": {'Themes': [], 'Questions': [], 'Comments': []},
        'Build Team': {'Themes': [], 'Questions': [], 'Comments': []},
        'People and Marketing Team': {'Themes': [], 'Questions': [], 'Comments': []},
        'Finance and Operations Team': {'Themes': [], 'Questions': [], 'Comments': []}
    }

    processing_errors = []
    files_processed = []

    for uploaded_file in uploaded_files:
        try:
            # Load the Excel file
            df = pd.read_excel(uploaded_file)

            # Clean the data
            df = df.dropna(how='all').dropna(axis=1, how='all')
            df.columns = [str(col).strip() for col in df.columns]

            # Categorize based on filename
            file_info = categorize_file(uploaded_file.name)
            section = file_info['section']
            team = file_info['team']
            filename = file_info['filename']

            # Add metadata to dataframe
            df['_source_file'] = filename
            df['_detected_team'] = team
            df['_file_section'] = section

            # Store in team-first structure
            if team and team in team_data and section in ['Themes', 'Questions', 'Comments']:
                team_data[team][section].append({
                    'data': df,
                    'filename': filename
                })
                files_processed.append({
                    'filename': filename,
                    'team': team,
                    'section': section
                })
            else:
                # If team not recognized, try to infer or put in a default location
                if not team:
                    processing_errors.append(f"Could not detect team from filename: {filename}")
                else:
                    processing_errors.append(f"Unrecognized team '{team}' in file: {filename}")

        except Exception as e:
            processing_errors.append(f"Error processing {uploaded_file.name}: {str(e)}")

    return team_data, files_processed, processing_errors

def get_team_section_data(team_data, team_name, section):
    """Get data for a specific team and section"""
    if team_name not in team_data or section not in team_data[team_name]:
        return pd.DataFrame()

    combined_df = pd.DataFrame()
    team_section_files = team_data[team_name][section]

    for file_info in team_section_files:
        if not combined_df.empty:
            combined_df = pd.concat([combined_df, file_info['data']], ignore_index=True)
        else:
            combined_df = file_info['data'].copy()

    return combined_df

def show_team_themes_analysis(team_data, team_name):
    """Show themes analysis for a specific team"""
    st.markdown(f"""
    <div class="section-header fade-in-up">
        <div class="section-title">🎯 Themes Analysis - {team_name}</div>
        <div class="section-subtitle">Detailed breakdown of survey themes and scores</div>
    </div>
    """, unsafe_allow_html=True)

    filtered_df = get_team_section_data(team_data, team_name, 'Themes')

    if filtered_df.empty:
        st.info(f"No themes data found for {team_name}. Upload Excel files with 'themes' in the filename.")
        return

    # Show file sources
    if '_source_file' in filtered_df.columns:
        source_files = filtered_df['_source_file'].unique()
        st.markdown(f"**📁 Data Sources:** {', '.join(source_files)}")
        st.markdown("---")

    # Show themes data similar to previous implementation
    if 'Theme' in filtered_df.columns:
        try:
            if 'Score' in filtered_df.columns:
                theme_df = filtered_df[['Theme', 'Score']].copy()
                theme_df['Score_Display'] = theme_df['Score'].astype(str)
                numeric_scores = pd.to_numeric(theme_df['Score'], errors='coerce')
                theme_df['Score_Display'] = numeric_scores.combine_first(theme_df['Score'].astype(str))
                theme_df = theme_df.dropna(subset=['Theme'])
                display_df = theme_df[['Theme', 'Score_Display']].rename(columns={'Score_Display': 'Score'})
            else:
                display_df = filtered_df[['Theme']].dropna()

            st.markdown("### 📋 Themes Overview")
            st.dataframe(
                display_df.head(20),
                width='stretch',
                column_config={
                    "Theme": st.column_config.TextColumn("Theme", width="large"),
                    "Score": st.column_config.TextColumn("Score", width="medium") if 'Score' in display_df.columns else None
                }
            )

            # Create visualization if scores are available
            if 'Score' in filtered_df.columns:
                try:
                    filtered_df_copy = filtered_df.copy()
                    filtered_df_copy['Score_Numeric'] = pd.to_numeric(filtered_df_copy['Score'], errors='coerce')
                    valid_scores = filtered_df_copy.dropna(subset=['Score_Numeric', 'Theme'])

                    if not valid_scores.empty:
                        theme_scores = valid_scores.groupby('Theme')['Score_Numeric'].mean().sort_values(ascending=False).head(10)

                        if not theme_scores.empty:
                            fig = px.bar(
                                x=theme_scores.values,
                                y=theme_scores.index,
                                orientation='h',
                                title=f"🏆 Top 10 Themes by Average Score - {team_name}",
                                labels={'x': 'Average Score', 'y': 'Theme'}
                            )
                            fig.update_layout(height=400)
                            st.plotly_chart(fig, width='stretch')
                except Exception as e:
                    st.warning(f"Could not create theme visualization: {str(e)}")

        except Exception as e:
            st.warning(f"Error displaying themes data: {str(e)}")
    else:
        st.info("No 'Theme' column found in the data.")

def analyze_questions_data(filtered_df):
    """Analyze questions data to extract affirmations, themes, and scores"""
    questions_data = []

    # Look for question/affirmation column
    question_col = None
    for col in filtered_df.columns:
        if any(word in col.lower() for word in ['question', 'affirmation', 'statement', 'item']):
            question_col = col
            break

    # Look for theme column
    theme_col = None
    for col in filtered_df.columns:
        if any(word in col.lower() for word in ['theme', 'category', 'domain']):
            theme_col = col
            break

    # Look for score column
    score_col = None
    for col in filtered_df.columns:
        if any(word in col.lower() for word in ['score', 'rating', 'value']) and not col.startswith('_'):
            score_col = col
            break

    if question_col and not filtered_df[question_col].isna().all():
        # Process row-by-row data (each row is a question/affirmation)
        for idx, row in filtered_df.iterrows():
            question_text = row.get(question_col)
            if pd.isna(question_text) or str(question_text).strip() == '':
                continue

            # Get theme
            theme = str(row.get(theme_col, "Not specified")) if theme_col else "Not specified"

            # Get score
            score_value = None
            if score_col and pd.notna(row.get(score_col)):
                score_raw = row.get(score_col)
                try:
                    score_numeric = pd.to_numeric(score_raw, errors='coerce')
                    if pd.notna(score_numeric):
                        score_value = f"{score_numeric:.2f}"
                    else:
                        score_value = str(score_raw)
                except:
                    score_value = str(score_raw)

            question_info = {
                "Affirmation": str(question_text).strip(),
                "Theme": theme,
                "Score": score_value if score_value else "No score"
            }
            questions_data.append(question_info)

    return questions_data, question_col, theme_col, score_col

def analyze_comments_data(filtered_df):
    """Analyze comments data to extract affirmations, themes, and comments"""
    comments_data = []

    # Look for affirmation/question column
    question_col = None
    for col in filtered_df.columns:
        if any(word in col.lower() for word in ['question', 'affirmation', 'statement', 'item']):
            question_col = col
            break

    # Look for theme column
    theme_col = None
    for col in filtered_df.columns:
        if any(word in col.lower() for word in ['theme', 'category', 'domain']):
            theme_col = col
            break

    # Look for comment column
    comment_col = None
    for col in filtered_df.columns:
        if any(word in col.lower() for word in ['comment', 'feedback', 'response', 'text', 'opinion']) and not col.startswith('_'):
            comment_col = col
            break

    if comment_col and not filtered_df[comment_col].isna().all():
        # Process row-by-row data (each row has affirmation, theme, and comment)
        for idx, row in filtered_df.iterrows():
            comment_text = row.get(comment_col)
            if pd.isna(comment_text) or str(comment_text).strip() == '':
                continue

            # Get affirmation/question
            affirmation = str(row.get(question_col, "Not specified")) if question_col else "Not specified"

            # Get theme
            theme = str(row.get(theme_col, "Not specified")) if theme_col else "Not specified"

            comment_info = {
                "Affirmation": affirmation.strip(),
                "Theme": theme,
                "Comment": str(comment_text).strip()
            }

            # Add team source if available
            if '_team_source' in row:
                comment_info['_team_source'] = row['_team_source']
            comments_data.append(comment_info)

    return comments_data, question_col, theme_col, comment_col

def show_team_questions_analysis(team_data, team_name):
    """Show questions analysis for a specific team"""
    st.markdown(f"""
    <div class="section-header fade-in-up">
        <div class="section-title">❓ Questions Analysis - {team_name}</div>
        <div class="section-subtitle">Survey questions breakdown and response patterns</div>
    </div>
    """, unsafe_allow_html=True)

    filtered_df = get_team_section_data(team_data, team_name, 'Questions')

    if filtered_df.empty:
        st.info(f"No questions data found for {team_name}. Upload Excel files with 'questions' or 'survey' in the filename.")
        return

    # Show file sources
    if '_source_file' in filtered_df.columns:
        source_files = filtered_df['_source_file'].unique()
        st.markdown(f"**📁 Data Sources:** {', '.join(source_files)}")
        st.markdown("---")

    st.markdown("### 📊 Questions Overview")

    # Use new questions analysis function
    questions_data, question_col, theme_col, score_col = analyze_questions_data(filtered_df)

    # Show data structure detection
    st.markdown(f"**📋 Data Structure Detected:**")
    st.markdown(f"- **Questions Column:** {question_col or '❌ Not found'}")
    st.markdown(f"- **Theme Column:** {theme_col or '❌ Not found'}")
    st.markdown(f"- **Score Column:** {score_col or '❌ Not found'}")
    st.markdown("---")

    if not questions_data:
        st.info("Expected Excel format: Each row should contain a question/affirmation with columns for Question, Theme, and Score")
        st.markdown("**Available columns in your data:**")
        for col in filtered_df.columns:
            if not col.startswith('_'):
                st.markdown(f"- {col}")
        return

    st.markdown(f"**📊 Found {len(questions_data)} questions/affirmations**")

    if questions_data:
        # Display enhanced questions table
        questions_df = pd.DataFrame(questions_data)
        st.dataframe(questions_df, width='stretch')

        # Show detailed question cards for better readability
        st.markdown("### 📋 Detailed Questions Analysis")

        # Add theme filter
        themes = list(set([q['Theme'] for q in questions_data if q['Theme'] != 'Not specified']))
        if themes:
            st.markdown("### 🔍 Filter by Theme")
            selected_theme = st.selectbox(
                "Choose theme to filter questions:",
                options=['All themes'] + sorted(themes),
                key=f"theme_filter_{team_name}"
            )

            if selected_theme != 'All themes':
                questions_data = [q for q in questions_data if q['Theme'] == selected_theme]
                st.markdown(f"**Filtered to {len(questions_data)} questions for theme: {selected_theme}**")

        # Sort alphabetically by affirmation
        questions_data_sorted = sorted(questions_data, key=lambda x: x['Affirmation'])

        # Display questions in card format grouped by theme
        current_theme = None
        for question in questions_data_sorted:
            # Group by theme
            if question['Theme'] != current_theme:
                current_theme = question['Theme']
                st.markdown(f"#### 🎯 {current_theme}")

            # Determine card style based on score (try to parse score for styling)
            card_style = "info"
            score_display = ""

            # Try to determine numeric value for styling
            try:
                if question['Score'] != "No score":
                    score_numeric = pd.to_numeric(question['Score'], errors='coerce')
                    if pd.notna(score_numeric):
                        if score_numeric >= 4.0:
                            card_style = "success"
                        elif score_numeric <= 2.5:
                            card_style = "warning"
                score_display = f"<div style='font-size: 1.2rem; font-weight: 700; color: #667eea; margin: 0.5rem 0;'>📊 Score: {question['Score']}</div>"
            except:
                score_display = f"<div style='font-size: 1rem; color: #64748b; margin: 0.5rem 0;'>📊 {question['Score']}</div>"

            st.markdown(f"""
            <div class="tab-info-card {card_style}">
                <div style="font-weight: 700; font-size: 1.1rem; margin-bottom: 0.5rem; color: #2c3e50;">
                    💬 {question['Affirmation']}
                </div>
                {score_display}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No valid question data found for analysis.")

def generate_narrative_analysis(sentiment_data, team_name):
    """Generate narrative bullet points about comment patterns and insights"""
    if sentiment_data['total_comments'] == 0:
        return []

    total = sentiment_data['total_comments']
    positive_pct = sentiment_data['positive_count'] / total * 100
    negative_pct = sentiment_data['negative_count'] / total * 100
    neutral_pct = sentiment_data['neutral_count'] / total * 100

    narrative_points = []

    # Overall sentiment narrative
    if positive_pct > 70:
        narrative_points.append(f"**Strong Positive Climate:** The team demonstrates exceptional satisfaction with {positive_pct:.0f}% of feedback being positive, indicating a healthy and motivated work environment.")
    elif positive_pct > 50:
        narrative_points.append(f"**Generally Positive Outlook:** With {positive_pct:.0f}% positive sentiment, the team shows good overall satisfaction, though there's room for addressing concerns.")
    elif negative_pct > 40:
        narrative_points.append(f"**Areas of Concern:** {negative_pct:.0f}% of feedback expresses dissatisfaction, suggesting significant challenges that require immediate attention and action.")
    else:
        narrative_points.append(f"**Mixed Perspectives:** The team shows balanced viewpoints with {positive_pct:.0f}% positive and {negative_pct:.0f}% negative sentiment, indicating diverse experiences.")

    # Engagement patterns
    if sentiment_data['avg_length'] > 150:
        narrative_points.append(f"**High Engagement:** Detailed responses (avg {sentiment_data['avg_length']:.0f} characters) show team members are deeply invested and willing to provide comprehensive feedback.")
    elif sentiment_data['avg_length'] > 80:
        narrative_points.append(f"**Moderate Engagement:** Responses show thoughtful consideration with meaningful detail, indicating good participation in the feedback process.")
    else:
        narrative_points.append(f"**Concise Feedback:** Brief responses (avg {sentiment_data['avg_length']:.0f} characters) suggest either efficiency in communication or potential hesitancy to elaborate.")

    # Theme patterns
    top_themes = sentiment_data['key_themes'][:3]
    if len(top_themes) >= 2:
        if 'communication' in top_themes and 'leadership' in top_themes:
            narrative_points.append("**Leadership & Communication Focus:** Comments frequently address leadership effectiveness and communication clarity, indicating these are priority areas for the team.")
        elif 'teamwork' in top_themes and 'culture' in top_themes:
            narrative_points.append("**Team Dynamics Emphasis:** Strong focus on collaborative relationships and workplace culture suggests team cohesion is a key concern.")
        elif 'workload' in top_themes and 'resources' in top_themes:
            narrative_points.append("**Operational Challenges:** Recurring mentions of workload and resource availability point to potential capacity or support issues.")
        elif 'growth' in top_themes:
            narrative_points.append("**Development-Oriented:** Significant attention to professional growth and development opportunities shows a forward-thinking, ambitious team culture.")
        else:
            narrative_points.append(f"**Key Focus Areas:** Team discussions center around {', '.join(top_themes)}, highlighting the primary concerns and interests of {team_name}.")

    # Positive patterns
    if positive_pct > 30:
        positive_themes = []
        if 'teamwork' in top_themes:
            positive_themes.append("collaborative relationships")
        if 'leadership' in top_themes and positive_pct > 50:
            positive_themes.append("leadership support")
        if 'growth' in top_themes:
            positive_themes.append("development opportunities")
        if 'culture' in top_themes and positive_pct > 50:
            positive_themes.append("positive work environment")

        if positive_themes:
            narrative_points.append(f"**Strengths Highlighted:** Team members particularly appreciate {', '.join(positive_themes)}, representing core strengths to build upon.")
        else:
            narrative_points.append("**Positive Momentum:** Despite challenges, team members recognize and value several aspects of their current work experience.")

    # Areas for improvement
    if negative_pct > 20:
        concern_areas = []
        if 'communication' in top_themes and negative_pct > 30:
            concern_areas.append("communication clarity and frequency")
        if 'workload' in top_themes:
            concern_areas.append("workload management and balance")
        if 'resources' in top_themes:
            concern_areas.append("resource availability and support")
        if 'processes' in top_themes and negative_pct > 25:
            concern_areas.append("workflow efficiency and procedures")

        if concern_areas:
            narrative_points.append(f"**Improvement Opportunities:** Feedback consistently points to {', '.join(concern_areas)} as areas requiring attention and strategic improvement.")
        else:
            narrative_points.append(f"**Challenge Areas:** With {negative_pct:.0f}% of feedback indicating concerns, there are clear opportunities for enhancing the team experience.")

    # Neutral feedback insights
    if neutral_pct > 40:
        narrative_points.append("**Observational Feedback:** A significant portion of responses are neutral, suggesting either satisfaction with status quo or a wait-and-see approach to changes.")

    return narrative_points

def generate_comprehensive_insights(sentiment_data, all_comments):
    """Generate comprehensive thematic insights based on actual comment analysis"""
    insights = []

    # Analyze common patterns in comments (without revealing actual text)
    comment_text = " ".join(all_comments).lower()
    total_comments = len(all_comments)
    avg_length = sum(len(comment) for comment in all_comments) / total_comments if total_comments > 0 else 0

    # Define keyword categories with more comprehensive terms
    theme_keywords = {
        'workload_time': ['time', 'hours', 'workload', 'overload', 'busy', 'deadline', 'pressure', 'stress', 'overwhelmed', 'capacity', 'bandwidth'],
        'clarity_vision': ['uncertainty', 'unclear', 'vision', 'direction', 'confused', 'clarity', 'priorities', 'goals', 'strategy', 'roadmap', 'purpose'],
        'culture_engagement': ['culture', 'events', 'engagement', 'fun', 'team', 'connection', 'balance', 'office', 'workplace', 'morale', 'atmosphere'],
        'communication': ['communication', 'feedback', 'transparent', 'updates', 'information', 'know', 'understand', 'listening', 'sharing'],
        'growth_development': ['growth', 'development', 'learning', 'career', 'skills', 'opportunities', 'advancement', 'training', 'mentorship'],
        'leadership_management': ['leadership', 'management', 'support', 'guidance', 'decision', 'leader', 'manager', 'supervisor'],
        'resources_tools': ['resources', 'tools', 'budget', 'equipment', 'technology', 'support', 'infrastructure'],
        'process_efficiency': ['process', 'efficiency', 'workflow', 'procedures', 'systems', 'organization', 'structure']
    }

    # Count keyword frequency and sentiment context
    theme_analysis = {}
    for theme, keywords in theme_keywords.items():
        keyword_count = sum(comment_text.count(keyword) for keyword in keywords)
        if keyword_count > 0:
            # Analyze sentiment context around these keywords
            positive_context = sum(1 for comment in all_comments
                                 if any(keyword in comment.lower() for keyword in keywords)
                                 and any(pos_word in comment.lower() for pos_word in ['good', 'great', 'excellent', 'positive', 'love', 'appreciate', 'satisfied']))
            negative_context = sum(1 for comment in all_comments
                                 if any(keyword in comment.lower() for keyword in keywords)
                                 and any(neg_word in comment.lower() for neg_word in ['bad', 'poor', 'lack', 'need', 'problem', 'issue', 'concern', 'difficult']))

            theme_analysis[theme] = {
                'count': keyword_count,
                'positive': positive_context,
                'negative': negative_context,
                'total_mentions': positive_context + negative_context
            }

    # Sort themes by relevance (total mentions)
    sorted_themes = sorted(theme_analysis.items(), key=lambda x: x[1]['total_mentions'], reverse=True)

    # Generate insights based on actual data patterns
    for theme, data in sorted_themes[:4]:  # Top 4 themes
        if data['total_mentions'] == 0:
            continue

        sentiment_ratio = data['positive'] / data['total_mentions'] if data['total_mentions'] > 0 else 0

        if theme == 'workload_time' and data['total_mentions'] >= 2:
            if sentiment_ratio < 0.3:
                insights.append({
                    'title': 'Workload & Time Management',
                    'description': f'Analysis of {data["total_mentions"]} workload-related comments reveals significant concerns about time constraints and capacity. The predominantly negative sentiment ({data["negative"]} negative vs {data["positive"]} positive mentions) suggests this is impacting productivity and employee well-being.'
                })
            else:
                insights.append({
                    'title': 'Workload Balance',
                    'description': f'Comments about workload and time management show mixed sentiment across {data["total_mentions"]} mentions. While some challenges exist, there\'s recognition of efforts to manage capacity effectively.'
                })

        elif theme == 'clarity_vision' and data['total_mentions'] >= 2:
            if sentiment_ratio < 0.4:
                insights.append({
                    'title': 'Organizational Clarity & Direction',
                    'description': f'Feedback indicates uncertainty about organizational direction appears in {data["total_mentions"]} comments. The pattern suggests employees are seeking clearer communication about company vision, priorities, and strategic direction.'
                })
            else:
                insights.append({
                    'title': 'Strategic Alignment',
                    'description': f'Comments about vision and direction show {data["total_mentions"]} mentions with generally positive sentiment, indicating good alignment with organizational goals.'
                })

        elif theme == 'culture_engagement' and data['total_mentions'] >= 2:
            if sentiment_ratio < 0.5:
                insights.append({
                    'title': 'Culture & Employee Engagement',
                    'description': f'Cultural and engagement themes appear in {data["total_mentions"]} comments with mixed sentiment. Feedback suggests opportunities to enhance team connection, workplace atmosphere, and engagement initiatives.'
                })
            else:
                insights.append({
                    'title': 'Positive Culture Momentum',
                    'description': f'Culture and engagement feedback across {data["total_mentions"]} mentions shows predominantly positive sentiment, indicating strong team dynamics and workplace satisfaction.'
                })

        elif theme == 'communication' and data['total_mentions'] >= 2:
            if sentiment_ratio < 0.4:
                insights.append({
                    'title': 'Communication & Information Sharing',
                    'description': f'Communication patterns emerge in {data["total_mentions"]} comments, primarily highlighting needs for improved transparency, feedback mechanisms, and information flow across the organization.'
                })
            else:
                insights.append({
                    'title': 'Communication Strengths',
                    'description': f'Communication feedback shows positive patterns across {data["total_mentions"]} mentions, indicating effective information sharing and feedback processes.'
                })

        elif theme == 'growth_development' and data['total_mentions'] >= 2:
            insights.append({
                'title': 'Professional Development Focus',
                'description': f'Development and growth themes appear in {data["total_mentions"]} comments, showing employee interest in career advancement, skill building, and learning opportunities within the organization.'
            })

    # If no strong themes emerge, provide general analysis
    if not insights:
        positive_pct = sentiment_data['positive_count'] / sentiment_data['total_comments'] * 100 if sentiment_data['total_comments'] > 0 else 0
        insights.append({
            'title': 'Overall Comment Analysis',
            'description': f'Across {total_comments} comments with an average length of {avg_length:.0f} characters, {positive_pct:.0f}% express positive sentiment. The feedback indicates engaged employees providing thoughtful input on their work experience.'
        })

    return insights

def generate_team_narrative_report(team_name, team_files):
    """Generate a comprehensive narrative report for a specific team"""
    report = f"# {team_name} - Survey Analysis Report\n\n"
    report += f"Generated on: {pd.Timestamp.now().strftime('%B %d, %Y')}\n\n"

    # Process each data type
    for category, files in team_files.items():
        if not files:
            continue

        report += f"## {category} Analysis\n\n"

        for file_info in files:
            if 'data' not in file_info or file_info['data'].empty:
                continue

            data = file_info['data']

            if category == 'Themes':
                report += generate_themes_narrative(data, team_name)
            elif category == 'Questions':
                report += generate_questions_narrative(data, team_name)
            elif category == 'Comments':
                report += generate_comments_narrative(data, team_name)

    return report

def generate_company_wide_narrative_report(team_data):
    """Generate a comprehensive narrative report for company-wide analysis"""
    report = "# Company-Wide Survey Analysis Report\n\n"
    report += f"Generated on: {pd.Timestamp.now().strftime('%B %d, %Y')}\n\n"
    report += "## Executive Summary\n\n"

    team_names = ["Andrew's Team", 'Build Team', 'People and Marketing Team', 'Finance and Operations Team']

    # Aggregate data
    all_themes_data = []
    all_questions_data = []
    all_comments_data = []

    for team_name in team_names:
        if team_name in team_data:
            for file_info in team_data[team_name].get('Themes', []):
                if 'data' in file_info and not file_info['data'].empty:
                    themes_data = file_info['data'].copy()
                    themes_data['_team_source'] = team_name
                    all_themes_data.append(themes_data)

            for file_info in team_data[team_name].get('Questions', []):
                if 'data' in file_info and not file_info['data'].empty:
                    questions_data = file_info['data'].copy()
                    questions_data['_team_source'] = team_name
                    all_questions_data.append(questions_data)

            for file_info in team_data[team_name].get('Comments', []):
                if 'data' in file_info and not file_info['data'].empty:
                    comments_data = file_info['data'].copy()
                    comments_data['_team_source'] = team_name
                    all_comments_data.append(comments_data)

    # Generate comprehensive analysis
    if all_themes_data:
        report += "## Themes Analysis\n\n"
        report += generate_company_themes_narrative(all_themes_data, team_names)

    if all_questions_data:
        report += "## Questions Analysis\n\n"
        report += generate_company_questions_narrative(all_questions_data, team_names)

    if all_comments_data:
        report += "## Comments Analysis\n\n"
        report += generate_company_comments_narrative(all_comments_data, team_names)

    return report

def generate_themes_narrative(data, team_name):
    """Generate narrative text for themes data"""
    narrative = f"### {team_name} - Themes Overview\n\n"

    # Look for theme and score columns
    theme_col = score_col = None
    for col in data.columns:
        if any(word in col.lower() for word in ['theme', 'category', 'domain']):
            theme_col = col
            break
    for col in data.columns:
        if any(word in col.lower() for word in ['score', 'rating', 'value']) and not col.startswith('_'):
            score_col = col
            break

    if theme_col and score_col:
        try:
            data['Score_Numeric'] = pd.to_numeric(data[score_col], errors='coerce')
            valid_data = data.dropna(subset=['Score_Numeric', theme_col])

            if not valid_data.empty:
                theme_scores = valid_data.groupby(theme_col)['Score_Numeric'].agg(['mean', 'count']).round(2)
                theme_scores = theme_scores.sort_values('mean', ascending=False)

                narrative += f"Analysis of {len(valid_data)} theme responses reveals the following patterns:\n\n"

                # Top themes
                top_themes = theme_scores.head(3)
                narrative += "**Top Performing Themes:**\n"
                for theme, row in top_themes.iterrows():
                    narrative += f"- {theme}: {row['mean']:.2f} average score ({row['count']} responses)\n"

                # Low themes
                low_themes = theme_scores.tail(3)
                narrative += "\n**Areas for Improvement:**\n"
                for theme, row in low_themes.iterrows():
                    narrative += f"- {theme}: {row['mean']:.2f} average score ({row['count']} responses)\n"

                narrative += "\n"
        except Exception:
            narrative += "Theme data structure could not be processed for detailed analysis.\n\n"

    return narrative

def generate_questions_narrative(data, team_name):
    """Generate narrative text for questions data"""
    narrative = f"### {team_name} - Questions Analysis\n\n"

    questions_data, question_col, theme_col, score_col = analyze_questions_data(data)

    if questions_data:
        narrative += f"Analysis of {len(questions_data)} questions reveals key insights about team perceptions:\n\n"

        # Group by theme
        theme_summary = {}
        for item in questions_data:
            theme = item.get('Theme', 'Not specified')
            if theme not in theme_summary:
                theme_summary[theme] = {'count': 0, 'scores': []}
            theme_summary[theme]['count'] += 1
            if item.get('Score') and item['Score'] != "No score":
                try:
                    score_numeric = pd.to_numeric(item['Score'], errors='coerce')
                    if pd.notna(score_numeric):
                        theme_summary[theme]['scores'].append(score_numeric)
                except:
                    pass

        # Sort by average score
        themes_with_scores = []
        for theme, data_item in theme_summary.items():
            avg_score = sum(data_item['scores']) / len(data_item['scores']) if data_item['scores'] else 0
            themes_with_scores.append((theme, data_item, avg_score))

        sorted_themes = sorted(themes_with_scores, key=lambda x: x[2] if x[2] > 0 else 999)

        narrative += "**Theme Performance (ordered by score):**\n"
        for theme, data_item, avg_score in sorted_themes:
            score_text = f"{avg_score:.2f}" if avg_score > 0 else "No scores"
            narrative += f"- {theme}: {score_text} ({data_item['count']} questions)\n"

        narrative += "\n"

    return narrative

def generate_comments_narrative(data, team_name):
    """Generate narrative text for comments data"""
    narrative = f"### {team_name} - Comments Analysis\n\n"

    comments_data, question_col, theme_col, comment_col = analyze_comments_data(data)

    if comments_data:
        all_comments = [item['Comment'] for item in comments_data if item.get('Comment') and item['Comment'] != "No comment"]

        if all_comments:
            sentiment_analysis = analyze_comment_sentiment(all_comments)
            narrative_points = generate_narrative_analysis(sentiment_analysis, team_name)

            narrative += f"Analysis of {len(all_comments)} comments from {team_name}:\n\n"

            # Overall sentiment
            positive_pct = sentiment_analysis['positive_count'] / sentiment_analysis['total_comments'] * 100
            narrative += f"**Overall Sentiment:** {positive_pct:.0f}% positive sentiment across {sentiment_analysis['total_comments']} comments\n\n"

            # Key insights
            narrative += "**Key Insights:**\n"
            for point in narrative_points[:3]:
                narrative += f"- {point.replace('**', '').replace('*', '')}\n"

            narrative += "\n"

    return narrative

def generate_company_themes_narrative(all_themes_data, team_names):
    """Generate company-wide themes narrative"""
    narrative = ""

    combined_themes = pd.concat(all_themes_data, ignore_index=True) if all_themes_data else pd.DataFrame()

    if not combined_themes.empty:
        theme_col = score_col = None
        for col in combined_themes.columns:
            if any(word in col.lower() for word in ['theme', 'category', 'domain']):
                theme_col = col
                break
        for col in combined_themes.columns:
            if any(word in col.lower() for word in ['score', 'rating', 'value']) and not col.startswith('_'):
                score_col = col
                break

        if theme_col and score_col:
            combined_themes['Score_Numeric'] = pd.to_numeric(combined_themes[score_col], errors='coerce')
            valid_data = combined_themes.dropna(subset=['Score_Numeric', theme_col])

            if not valid_data.empty:
                theme_scores = valid_data.groupby(theme_col)['Score_Numeric'].agg(['mean', 'count']).round(2)
                theme_scores = theme_scores.sort_values('mean', ascending=False)

                narrative += f"Company-wide theme analysis across {len(valid_data)} responses from all teams:\n\n"

                narrative += "**Top Performing Themes:**\n"
                for theme, row in theme_scores.head(5).iterrows():
                    narrative += f"- {theme}: {row['mean']:.2f} average ({row['count']} total responses)\n"

                narrative += "\n**Areas Needing Attention:**\n"
                for theme, row in theme_scores.tail(5).iterrows():
                    narrative += f"- {theme}: {row['mean']:.2f} average ({row['count']} total responses)\n"

                narrative += "\n"

    return narrative

def generate_company_questions_narrative(all_questions_data, team_names):
    """Generate company-wide questions narrative"""
    narrative = ""

    combined_questions = pd.concat(all_questions_data, ignore_index=True) if all_questions_data else pd.DataFrame()

    if not combined_questions.empty:
        questions_data, question_col, theme_col, score_col = analyze_questions_data(combined_questions)

        if questions_data:
            # Group by theme and analyze
            theme_summary = {}
            for item in questions_data:
                theme = item.get('Theme', 'Not specified')
                if theme not in theme_summary:
                    theme_summary[theme] = {'count': 0, 'scores': []}
                theme_summary[theme]['count'] += 1
                if item.get('Score') and item['Score'] != "No score":
                    try:
                        score_numeric = pd.to_numeric(item['Score'], errors='coerce')
                        if pd.notna(score_numeric):
                            theme_summary[theme]['scores'].append(score_numeric)
                    except:
                        pass

            narrative += f"Company-wide questions analysis covering {len(questions_data)} questions across all teams:\n\n"

            # Sort by score
            themes_with_scores = []
            for theme, data in theme_summary.items():
                avg_score = sum(data['scores']) / len(data['scores']) if data['scores'] else 0
                themes_with_scores.append((theme, data, avg_score))

            sorted_themes = sorted(themes_with_scores, key=lambda x: x[2] if x[2] > 0 else 999)

            narrative += "**Theme Performance Summary:**\n"
            for theme, data, avg_score in sorted_themes:
                score_text = f"{avg_score:.2f}" if avg_score > 0 else "No scores available"
                narrative += f"- {theme}: {score_text} ({data['count']} questions)\n"

            narrative += "\n"

    return narrative

def generate_company_comments_narrative(all_comments_data, team_names):
    """Generate company-wide comments narrative"""
    narrative = ""

    combined_comments = pd.concat(all_comments_data, ignore_index=True) if all_comments_data else pd.DataFrame()

    if not combined_comments.empty:
        comments_data, question_col, theme_col, comment_col = analyze_comments_data(combined_comments)

        if comments_data:
            all_comments = [item['Comment'] for item in comments_data if item.get('Comment') and item['Comment'] != "No comment"]

            if all_comments:
                sentiment_analysis = analyze_comment_sentiment(all_comments)
                insights = generate_comprehensive_insights(sentiment_analysis, all_comments)

                narrative += f"Company-wide comment analysis across {len(all_comments)} comments from all teams:\n\n"

                positive_pct = sentiment_analysis['positive_count'] / sentiment_analysis['total_comments'] * 100
                narrative += f"**Overall Sentiment:** {positive_pct:.0f}% positive sentiment company-wide\n\n"

                narrative += "**Key Organizational Insights:**\n"
                for insight in insights:
                    narrative += f"**{insight['title']}:** {insight['description']}\n\n"

    return narrative

# Google Docs Integration Functions
def setup_google_credentials():
    """Setup Google OAuth credentials"""
    if not GOOGLE_APIS_AVAILABLE:
        return None, "Google APIs not available. Install required packages."

    # OAuth 2.0 configuration
    CLIENT_CONFIG = {
        "web": {
            "client_id": st.secrets.get("google_client_id", ""),
            "client_secret": st.secrets.get("google_client_secret", ""),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [st.secrets.get("redirect_uri", "http://localhost:8501")]
        }
    }

    SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive.file']

    return CLIENT_CONFIG, SCOPES

def get_google_auth_url():
    """Generate Google OAuth authorization URL"""
    try:
        client_config, scopes = setup_google_credentials()
        if client_config is None:
            return None, "Google configuration not available"

        flow = Flow.from_client_config(
            client_config,
            scopes=scopes,
            redirect_uri=client_config['web']['redirect_uris'][0]
        )

        auth_url, _ = flow.authorization_url(prompt='consent')
        return auth_url, None
    except Exception as e:
        return None, str(e)

def handle_google_callback(auth_code):
    """Handle Google OAuth callback and get credentials"""
    try:
        client_config, scopes = setup_google_credentials()
        if client_config is None:
            return None, "Google configuration not available"

        flow = Flow.from_client_config(
            client_config,
            scopes=scopes,
            redirect_uri=client_config['web']['redirect_uris'][0]
        )

        flow.fetch_token(code=auth_code)
        credentials = flow.credentials

        return credentials, None
    except Exception as e:
        return None, str(e)

def create_google_doc(title, content, credentials):
    """Create a Google Doc with the provided content"""
    try:
        service = build('docs', 'v1', credentials=credentials)

        # Create a new document
        document = {
            'title': title
        }
        doc = service.documents().create(body=document).execute()
        doc_id = doc.get('documentId')

        # Convert markdown-style content to Google Docs format
        requests = []

        # Split content into lines and process
        lines = content.split('\n')
        current_position = 1

        for line in lines:
            if line.strip():
                # Handle headers
                if line.startswith('# '):
                    # Main title
                    text = line[2:].strip()
                    requests.append({
                        'insertText': {
                            'location': {'index': current_position},
                            'text': text + '\n'
                        }
                    })
                    requests.append({
                        'updateParagraphStyle': {
                            'range': {
                                'startIndex': current_position,
                                'endIndex': current_position + len(text)
                            },
                            'paragraphStyle': {
                                'namedStyleType': 'TITLE'
                            },
                            'fields': 'namedStyleType'
                        }
                    })
                    current_position += len(text) + 1

                elif line.startswith('## '):
                    # Section headers
                    text = line[3:].strip()
                    requests.append({
                        'insertText': {
                            'location': {'index': current_position},
                            'text': text + '\n'
                        }
                    })
                    requests.append({
                        'updateParagraphStyle': {
                            'range': {
                                'startIndex': current_position,
                                'endIndex': current_position + len(text)
                            },
                            'paragraphStyle': {
                                'namedStyleType': 'HEADING_1'
                            },
                            'fields': 'namedStyleType'
                        }
                    })
                    current_position += len(text) + 1

                elif line.startswith('### '):
                    # Subsection headers
                    text = line[4:].strip()
                    requests.append({
                        'insertText': {
                            'location': {'index': current_position},
                            'text': text + '\n'
                        }
                    })
                    requests.append({
                        'updateParagraphStyle': {
                            'range': {
                                'startIndex': current_position,
                                'endIndex': current_position + len(text)
                            },
                            'paragraphStyle': {
                                'namedStyleType': 'HEADING_2'
                            },
                            'fields': 'namedStyleType'
                        }
                    })
                    current_position += len(text) + 1

                elif line.startswith('**') and line.endswith('**'):
                    # Bold text
                    text = line[2:-2].strip()
                    requests.append({
                        'insertText': {
                            'location': {'index': current_position},
                            'text': text + '\n'
                        }
                    })
                    requests.append({
                        'updateTextStyle': {
                            'range': {
                                'startIndex': current_position,
                                'endIndex': current_position + len(text)
                            },
                            'textStyle': {
                                'bold': True
                            },
                            'fields': 'bold'
                        }
                    })
                    current_position += len(text) + 1

                else:
                    # Regular text
                    text = line.strip()
                    if text:
                        requests.append({
                            'insertText': {
                                'location': {'index': current_position},
                                'text': text + '\n'
                            }
                        })
                        current_position += len(text) + 1
            else:
                # Empty line
                requests.append({
                    'insertText': {
                        'location': {'index': current_position},
                        'text': '\n'
                    }
                })
                current_position += 1

        # Apply all formatting
        if requests:
            service.documents().batchUpdate(
                documentId=doc_id,
                body={'requests': requests}
            ).execute()

        # Return the document URL
        doc_url = f"https://docs.google.com/document/d/{doc_id}/edit"
        return doc_url, None

    except HttpError as error:
        return None, f"Google API error: {error}"
    except Exception as e:
        return None, str(e)

def show_google_docs_integration(report_content, report_title):
    """Show Google Docs integration UI"""
    if not GOOGLE_APIS_AVAILABLE:
        st.error("Google Docs integration requires additional packages. Please contact your administrator.")
        return

    st.markdown("### 🔗 Google Docs Integration")

    # Check if we have Google credentials configured
    if not all([
        st.secrets.get("google_client_id"),
        st.secrets.get("google_client_secret"),
        st.secrets.get("redirect_uri")
    ]):
        st.error("Google Docs integration is not configured. Please add your Google API credentials to Streamlit secrets.")
        st.info("""
        To enable Google Docs integration:
        1. Create a project in Google Cloud Console
        2. Enable Google Docs API and Google Drive API
        3. Create OAuth 2.0 credentials
        4. Add the credentials to your Streamlit secrets:
           - `google_client_id`
           - `google_client_secret`
           - `redirect_uri` (usually your Streamlit app URL)
        """)
        return

    # Check if user is authenticated
    if 'google_credentials' not in st.session_state:
        st.info("Click below to authenticate with Google and create documents automatically.")

        auth_url, error = get_google_auth_url()
        if error:
            st.error(f"Error generating auth URL: {error}")
            return

        if st.button("🔐 Login with Google"):
            st.markdown(f"[Click here to authenticate with Google]({auth_url})")
            st.info("After authentication, copy the authorization code from the URL and paste it below.")

        # Input for authorization code
        auth_code = st.text_input("Authorization Code", help="Paste the code from Google OAuth redirect URL")
        if auth_code:
            credentials, error = handle_google_callback(auth_code)
            if error:
                st.error(f"Error handling callback: {error}")
            else:
                st.session_state.google_credentials = credentials
                st.success("Successfully authenticated with Google!")
                st.rerun()
    else:
        # User is authenticated, show create document option
        st.success("✅ Authenticated with Google")

        if st.button("📄 Create Google Doc"):
            with st.spinner("Creating Google Document..."):
                doc_url, error = create_google_doc(
                    report_title,
                    report_content,
                    st.session_state.google_credentials
                )

                if error:
                    st.error(f"Error creating document: {error}")
                else:
                    st.success("Document created successfully!")
                    st.markdown(f"[📖 Open your Google Doc]({doc_url})")

        if st.button("🔓 Logout from Google"):
            if 'google_credentials' in st.session_state:
                del st.session_state.google_credentials
            st.rerun()

def analyze_comment_sentiment(comments_list):
    """Analyze sentiment of comments without exposing actual content"""
    if not comments_list:
        return {
            'total_comments': 0,
            'positive_count': 0,
            'negative_count': 0,
            'neutral_count': 0,
            'avg_length': 0,
            'key_themes': []
        }

    positive_words = [
        'good', 'great', 'excellent', 'amazing', 'wonderful', 'fantastic', 'outstanding', 'perfect',
        'love', 'like', 'appreciate', 'satisfied', 'happy', 'pleased', 'impressed', 'positive',
        'strong', 'effective', 'successful', 'helpful', 'supportive', 'clear', 'transparent',
        'collaborative', 'innovative', 'efficient', 'smooth', 'well', 'better', 'improved',
        'progress', 'growth', 'success', 'achievement', 'opportunity', 'benefit'
    ]

    negative_words = [
        'bad', 'terrible', 'awful', 'horrible', 'disappointing', 'frustrating', 'annoying',
        'hate', 'dislike', 'unsatisfied', 'unhappy', 'disappointed', 'concerned', 'worried',
        'problem', 'issue', 'challenge', 'difficulty', 'confusion', 'unclear', 'poor',
        'ineffective', 'unsuccessful', 'lacking', 'missing', 'insufficient', 'inadequate',
        'slow', 'delayed', 'complicated', 'confusing', 'overwhelming', 'stressful'
    ]

    total_comments = len(comments_list)
    positive_count = 0
    negative_count = 0
    neutral_count = 0
    lengths = []

    # Common themes extraction (anonymized)
    theme_words = {
        'communication': ['communication', 'communicate', 'talk', 'discuss', 'meeting', 'update'],
        'leadership': ['leadership', 'leader', 'management', 'manager', 'direction', 'guidance'],
        'teamwork': ['team', 'collaboration', 'together', 'support', 'help', 'cooperation'],
        'processes': ['process', 'procedure', 'workflow', 'system', 'method', 'approach'],
        'growth': ['growth', 'development', 'learning', 'training', 'skill', 'improve'],
        'workload': ['workload', 'busy', 'time', 'deadline', 'pressure', 'stress'],
        'resources': ['resource', 'tool', 'equipment', 'budget', 'funding', 'support'],
        'culture': ['culture', 'environment', 'atmosphere', 'morale', 'values', 'mission']
    }

    theme_mentions = {theme: 0 for theme in theme_words.keys()}

    for comment in comments_list:
        if pd.isna(comment) or str(comment).strip() == '':
            continue

        comment_text = str(comment).lower()
        lengths.append(len(comment_text))

        # Count positive and negative words
        pos_score = sum(1 for word in positive_words if word in comment_text)
        neg_score = sum(1 for word in negative_words if word in comment_text)

        # Classify sentiment
        if pos_score > neg_score:
            positive_count += 1
        elif neg_score > pos_score:
            negative_count += 1
        else:
            neutral_count += 1

        # Count theme mentions
        for theme, keywords in theme_words.items():
            if any(keyword in comment_text for keyword in keywords):
                theme_mentions[theme] += 1

    # Get top themes
    top_themes = sorted(theme_mentions.items(), key=lambda x: x[1], reverse=True)[:3]
    key_themes = [theme for theme, count in top_themes if count > 0]

    return {
        'total_comments': total_comments,
        'positive_count': positive_count,
        'negative_count': negative_count,
        'neutral_count': neutral_count,
        'avg_length': sum(lengths) / len(lengths) if lengths else 0,
        'key_themes': key_themes,
        'theme_distribution': theme_mentions
    }

def show_team_comments_analysis(team_data, team_name):
    """Show comments analysis for a specific team"""
    st.markdown(f"""
    <div class="section-header fade-in-up">
        <div class="section-title">💬 Comments Analysis - {team_name}</div>
        <div class="section-subtitle">Survey comments and feedback</div>
    </div>
    """, unsafe_allow_html=True)

    filtered_df = get_team_section_data(team_data, team_name, 'Comments')

    if filtered_df.empty:
        st.info(f"No comments data found for {team_name}. Upload Excel files with 'comments' or 'feedback' in the filename.")
        return

    # Show file sources
    if '_source_file' in filtered_df.columns:
        source_files = filtered_df['_source_file'].unique()
        st.markdown(f"**📁 Data Sources:** {', '.join(source_files)}")
        st.markdown("---")

    # Look for comment columns
    comment_columns = [col for col in filtered_df.columns
                      if any(keyword in col.lower() for keyword in
                      ['comment', 'feedback', 'response', 'text', 'opinion', 'note']) and not col.startswith('_')]

    # If no specific comment columns found, look for text columns
    if not comment_columns:
        comment_columns = [col for col in filtered_df.columns
                         if filtered_df[col].dtype == 'object' and
                         not col.startswith('_') and
                         col.lower() not in ['team', 'name', 'id']]

    # Try to use the new row-based structure first
    comments_data, question_col, theme_col, comment_col = analyze_comments_data(filtered_df)

    if comments_data:
        st.markdown("### 📝 Comments & Sentiment Analysis")

        # Collect all comments for sentiment analysis
        all_comments = []
        theme_comments = {}

        for item in comments_data:
            if item.get('Comment') and item['Comment'] != "No comment":
                all_comments.append(item['Comment'])
                theme = item.get('Theme', 'Not specified')
                if theme not in theme_comments:
                    theme_comments[theme] = []
                theme_comments[theme].append(item['Comment'])

        if all_comments:
            # Perform sentiment analysis
            sentiment_analysis = analyze_comment_sentiment(all_comments)

            # Display overview metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Comments", sentiment_analysis['total_comments'])
            with col2:
                st.metric("Positive", sentiment_analysis['positive_count'],
                         delta=f"{sentiment_analysis['positive_count']/sentiment_analysis['total_comments']*100:.1f}%")
            with col3:
                st.metric("Negative", sentiment_analysis['negative_count'],
                         delta=f"{sentiment_analysis['negative_count']/sentiment_analysis['total_comments']*100:.1f}%")
            with col4:
                st.metric("Neutral", sentiment_analysis['neutral_count'],
                         delta=f"{sentiment_analysis['neutral_count']/sentiment_analysis['total_comments']*100:.1f}%")

            # Sentiment overview narrative
            st.markdown("### 📊 Sentiment Overview")

            positive_pct = sentiment_analysis['positive_count'] / sentiment_analysis['total_comments'] * 100
            negative_pct = sentiment_analysis['negative_count'] / sentiment_analysis['total_comments'] * 100
            neutral_pct = sentiment_analysis['neutral_count'] / sentiment_analysis['total_comments'] * 100

            # Generate overview analysis
            if positive_pct > 60:
                sentiment_tone = "predominantly positive"
                sentiment_icon = "😊"
            elif negative_pct > 40:
                sentiment_tone = "concerning with notable negative feedback"
                sentiment_icon = "😟"
            elif neutral_pct > 50:
                sentiment_tone = "neutral with mixed feelings"
                sentiment_icon = "😐"
            else:
                sentiment_tone = "balanced with varied perspectives"
                sentiment_icon = "⚖️"

            # Generate narrative analysis
            narrative_points = generate_narrative_analysis(sentiment_analysis, team_name)

            st.markdown(f"""
            <div class="tab-info-card info">
                <h4>{sentiment_icon} Overall Sentiment Analysis</h4>
                <p>Analysis of {sentiment_analysis['total_comments']} comments from {team_name} reveals the following patterns and insights:</p>
            </div>
            """, unsafe_allow_html=True)

            # Display narrative bullet points
            for point in narrative_points:
                st.markdown(f"• {point}")

            st.markdown("---")

            # Theme-based sentiment breakdown
            if len(theme_comments) > 1:
                st.markdown("### 🎯 Sentiment by Theme")

                for theme, comments in theme_comments.items():
                    if theme != 'Not specified':
                        theme_sentiment = analyze_comment_sentiment(comments)
                        theme_positive_pct = theme_sentiment['positive_count'] / theme_sentiment['total_comments'] * 100 if theme_sentiment['total_comments'] > 0 else 0

                        if theme_positive_pct > 60:
                            theme_icon = "✅"
                            theme_status = "Positive"
                        elif theme_positive_pct < 40:
                            theme_icon = "⚠️"
                            theme_status = "Needs Attention"
                        else:
                            theme_icon = "ℹ️"
                            theme_status = "Mixed"

                        st.markdown(f"""
                        <div class="tab-info-card {'success' if theme_positive_pct > 60 else 'warning' if theme_positive_pct < 40 else 'info'}">
                            <strong>{theme_icon} {theme}</strong> - {theme_status}<br>
                            <small>{theme_sentiment['total_comments']} comments | {theme_positive_pct:.0f}% positive sentiment</small>
                        </div>
                        """, unsafe_allow_html=True)
        else:
            st.info("Comments detected but no text content found for analysis.")

    elif comment_columns:
        # Fallback to old column-based structure
        st.markdown("### 📝 Comments & Sentiment Analysis")

        # Collect all comments from columns
        all_comments = []
        for col in comment_columns:
            comments = filtered_df[col].dropna().tolist()
            all_comments.extend(comments)

        if all_comments:
            sentiment_analysis = analyze_comment_sentiment(all_comments)

            # Display metrics and analysis (same as above)
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Comments", sentiment_analysis['total_comments'])
            with col2:
                st.metric("Positive", sentiment_analysis['positive_count'])
            with col3:
                st.metric("Negative", sentiment_analysis['negative_count'])
            with col4:
                st.metric("Neutral", sentiment_analysis['neutral_count'])

            st.markdown("### 📊 Sentiment Overview")
            positive_pct = sentiment_analysis['positive_count'] / sentiment_analysis['total_comments'] * 100

            st.markdown(f"""
            <div class="tab-info-card info">
                <h4>💬 Comment Analysis Summary</h4>
                <p>Analysis of {sentiment_analysis['total_comments']} comments reveals a {positive_pct:.1f}% positive sentiment rate.</p>
                <p><strong>Key themes discussed:</strong> {", ".join(sentiment_analysis['key_themes']) if sentiment_analysis['key_themes'] else "Various topics"}</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.info("No comments found for sentiment analysis.")
    else:
        st.info("No comment data found. Upload Excel files with 'comment', 'feedback', or similar columns.")

def show_insights_section(categorized_data, team_filter, all_teams):
    """Show the insights overview section"""
    st.markdown("""
    <div class="section-header fade-in-up">
        <div class="section-title">📊 Survey Insights</div>
        <div class="section-subtitle">Overall survey performance and key metrics</div>
    </div>
    """, unsafe_allow_html=True)

    # Get filtered data for insights
    filtered_df = get_filtered_data(categorized_data, 'Insights', team_filter)

    if filtered_df.empty:
        # If no specific insights data, combine data from all sections
        all_data = pd.DataFrame()
        for section in ['Comments', 'Themes', 'Questions']:
            section_data = get_filtered_data(categorized_data, section, team_filter)
            if not section_data.empty:
                all_data = pd.concat([all_data, section_data], ignore_index=True)
        filtered_df = all_data

    # Create stats cards
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value">{len(filtered_df)}</div>
            <div class="stat-label">Total Responses</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        avg_score = 0
        if 'Score' in filtered_df.columns:
            try:
                # Try to convert to numeric and calculate mean
                numeric_scores = pd.to_numeric(filtered_df['Score'], errors='coerce')
                avg_score = numeric_scores.mean() if not numeric_scores.isna().all() else 0
            except:
                avg_score = 0

        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value">{avg_score:.1f}</div>
            <div class="stat-label">Average Score</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        completion_rate = 85.2  # Placeholder
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value">{completion_rate}%</div>
            <div class="stat-label">Completion Rate</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        themes_count = filtered_df.get('Theme', pd.Series()).nunique() if 'Theme' in filtered_df.columns else 0
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-value">{themes_count}</div>
            <div class="stat-label">Unique Themes</div>
        </div>
        """, unsafe_allow_html=True)

    # Show overall insights
    if not filtered_df.empty:
        try:
            insights = analyze_survey_data(filtered_df)
            narrative_text = generate_survey_narrative(insights)

            st.markdown('<div class="fade-in-up" style="margin-top: 2rem;">', unsafe_allow_html=True)
            st.markdown("### 📝 Executive Summary")
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
                        padding: 2rem; border-radius: 16px; margin: 1rem 0;
                        border-left: 4px solid #667eea;">
                {narrative_text[:500]}...
            </div>
            """, unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        except Exception as e:
            st.markdown('<div class="fade-in-up" style="margin-top: 2rem;">', unsafe_allow_html=True)
            st.markdown("### 📝 Executive Summary")
            st.info("Analysis in progress. Please ensure your data contains properly formatted numeric scores for detailed insights.")
            st.markdown('</div>', unsafe_allow_html=True)

def show_themes_section(categorized_data, team_filter, all_teams):
    """Show the themes analysis section"""
    st.markdown("""
    <div class="section-header fade-in-up">
        <div class="section-title">🎯 Themes Analysis</div>
        <div class="section-subtitle">Detailed breakdown of survey themes and scores</div>
    </div>
    """, unsafe_allow_html=True)

    # Get filtered data for themes
    filtered_df = get_filtered_data(categorized_data, 'Themes', team_filter)

    # Show file sources
    if not filtered_df.empty and '_source_file' in filtered_df.columns:
        source_files = filtered_df['_source_file'].unique()
        st.markdown(f"**📁 Data Sources:** {', '.join(source_files)}")
        st.markdown("---")

    if not filtered_df.empty and 'Theme' in filtered_df.columns:
        # Show themes data
        try:
            if 'Score' in filtered_df.columns:
                # Create a clean dataframe with proper data types
                theme_df = filtered_df[['Theme', 'Score']].copy()

                # Convert scores to numeric, keeping original values for display
                theme_df['Score_Display'] = theme_df['Score'].astype(str)

                # Try to convert to numeric for those that can be converted
                numeric_scores = pd.to_numeric(theme_df['Score'], errors='coerce')
                theme_df['Score_Display'] = numeric_scores.combine_first(theme_df['Score'].astype(str))

                # Remove rows where both Theme and Score are null
                theme_df = theme_df.dropna(subset=['Theme'])

                display_df = theme_df[['Theme', 'Score_Display']].rename(columns={'Score_Display': 'Score'})
            else:
                display_df = filtered_df[['Theme']].dropna()

            st.markdown("### 📋 Themes Overview")
            st.dataframe(
                display_df.head(20),
                width='stretch',
                column_config={
                    "Theme": st.column_config.TextColumn("Theme", width="large"),
                    "Score": st.column_config.TextColumn("Score", width="medium") if 'Score' in display_df.columns else None
                }
            )
        except Exception as e:
            st.warning(f"Error displaying themes data: {str(e)}")
            # Fallback to simple display
            if 'Theme' in filtered_df.columns:
                simple_themes = filtered_df['Theme'].dropna().head(20)
                st.write("**Themes:**")
                for theme in simple_themes:
                    st.write(f"• {theme}")

        # Create theme visualization
        if 'Score' in filtered_df.columns:
            try:
                # Convert scores to numeric and group by theme
                filtered_df_copy = filtered_df.copy()
                filtered_df_copy['Score_Numeric'] = pd.to_numeric(filtered_df_copy['Score'], errors='coerce')

                # Only include rows with valid numeric scores
                valid_scores = filtered_df_copy.dropna(subset=['Score_Numeric', 'Theme'])

                if not valid_scores.empty:
                    theme_scores = valid_scores.groupby('Theme')['Score_Numeric'].mean().sort_values(ascending=False).head(10)

                    if not theme_scores.empty:
                        fig = px.bar(
                            x=theme_scores.values,
                            y=theme_scores.index,
                            orientation='h',
                            title="🏆 Top 10 Themes by Average Score",
                            labels={'x': 'Average Score', 'y': 'Theme'}
                        )
                        fig.update_layout(height=400)
                        st.plotly_chart(fig, width='stretch')
                    else:
                        st.info("No valid numeric scores found for theme visualization.")
                else:
                    st.info("No valid theme and score data available for visualization.")
            except Exception as e:
                st.warning(f"Could not create theme visualization: {str(e)}")
    else:
        st.info("No theme data available for the selected filter.")

def show_comments_section(categorized_data, team_filter, all_teams):
    """Show the comments analysis section"""
    st.markdown("""
    <div class="section-header fade-in-up">
        <div class="section-title">💬 Comments Analysis</div>
        <div class="section-subtitle">Survey comments and feedback by team</div>
    </div>
    """, unsafe_allow_html=True)

    # Get filtered data for comments
    filtered_df = get_filtered_data(categorized_data, 'Comments', team_filter)

    # Show file sources
    if not filtered_df.empty and '_source_file' in filtered_df.columns:
        source_files = filtered_df['_source_file'].unique()
        st.markdown(f"**📁 Data Sources:** {', '.join(source_files)}")
        st.markdown("---")

    if not filtered_df.empty:
        # Look for comment columns (more comprehensive search)
        comment_columns = [col for col in filtered_df.columns
                          if any(keyword in col.lower() for keyword in
                          ['comment', 'feedback', 'response', 'text', 'opinion', 'note'])]

        # If no specific comment columns found, look for text columns
        if not comment_columns:
            comment_columns = [col for col in filtered_df.columns
                             if filtered_df[col].dtype == 'object' and
                             not col.startswith('_') and
                             col.lower() not in ['team', 'name', 'id']]

        if comment_columns:
            st.markdown("### 📝 Comments & Feedback")

            # Show statistics
            total_comments = 0
            for col in comment_columns:
                total_comments += filtered_df[col].notna().sum()

            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Comments", total_comments)
            with col2:
                st.metric("Comment Categories", len(comment_columns))

            # Display comments
            for col in comment_columns[:3]:  # Show first 3 comment columns
                comments = filtered_df[col].dropna().head(10)
                if not comments.empty:
                    st.markdown(f"#### {col.replace('_', ' ').title()}")

                    # Show team breakdown if available
                    if '_detected_team' in filtered_df.columns and team_filter == "All Teams":
                        team_comments = filtered_df.groupby('_detected_team')[col].count()
                        team_summary = ", ".join([f"{team}: {count}" for team, count in team_comments.items() if pd.notna(team)])
                        if team_summary:
                            st.markdown(f"**Team Breakdown:** {team_summary}")

                    for i, (idx, comment) in enumerate(comments.items(), 1):
                        team_info = ""
                        if '_detected_team' in filtered_df.columns:
                            team = filtered_df.loc[idx, '_detected_team']
                            if pd.notna(team):
                                team_info = f" <small>({team})</small>"

                        st.markdown(f"""
                        <div class="metric-card" style="margin: 0.5rem 0;">
                            <strong>Comment #{i}:</strong>{team_info}<br>
                            {str(comment)[:300]}{"..." if len(str(comment)) > 300 else ""}
                        </div>
                        """, unsafe_allow_html=True)
        else:
            st.info("No comment data found. Upload Excel files with 'comment', 'feedback', or similar columns.")

def show_questions_section(categorized_data, team_filter, all_teams):
    """Show the questions analysis section"""
    st.markdown("""
    <div class="section-header fade-in-up">
        <div class="section-title">❓ Questions Analysis</div>
        <div class="section-subtitle">Survey questions breakdown and response patterns</div>
    </div>
    """, unsafe_allow_html=True)

    # Get filtered data for questions
    filtered_df = get_filtered_data(categorized_data, 'Questions', team_filter)

    # Show file sources
    if not filtered_df.empty and '_source_file' in filtered_df.columns:
        source_files = filtered_df['_source_file'].unique()
        st.markdown(f"**📁 Data Sources:** {', '.join(source_files)}")
        st.markdown("---")

    st.markdown("### 📊 Questions Overview")

    # Use new questions analysis function
    questions_data, question_col, theme_col, score_col = analyze_questions_data(filtered_df)

    # Show data structure detection
    st.markdown(f"**📋 Data Structure Detected:**")
    st.markdown(f"- **Questions Column:** {question_col or '❌ Not found'}")
    st.markdown(f"- **Theme Column:** {theme_col or '❌ Not found'}")
    st.markdown(f"- **Score Column:** {score_col or '❌ Not found'}")
    st.markdown("---")

    if not questions_data:
        st.info("Expected Excel format: Each row should contain a question/affirmation with columns for Question, Theme, and Score")
        st.markdown("**Available columns in your data:**")
        for col in filtered_df.columns:
            if not col.startswith('_'):
                st.markdown(f"- {col}")
        return

    st.markdown(f"**📊 Found {len(questions_data)} questions/affirmations**")

    if questions_data:
        # Display enhanced questions table
        questions_df = pd.DataFrame(questions_data)
        st.dataframe(questions_df, width='stretch')

        # Show detailed question cards for better readability
        st.markdown("### 📋 Detailed Questions Analysis")

        # Add theme filter
        themes = list(set([q['Theme'] for q in questions_data if q['Theme'] != 'Not specified']))
        if themes:
            st.markdown("### 🔍 Filter by Theme")
            selected_theme = st.selectbox(
                "Choose theme to filter questions:",
                options=['All themes'] + sorted(themes),
                key=f"theme_filter_{team_name}"
            )

            if selected_theme != 'All themes':
                questions_data = [q for q in questions_data if q['Theme'] == selected_theme]
                st.markdown(f"**Filtered to {len(questions_data)} questions for theme: {selected_theme}**")

        # Sort alphabetically by affirmation
        questions_data_sorted = sorted(questions_data, key=lambda x: x['Affirmation'])

        # Display questions in card format grouped by theme
        current_theme = None
        for question in questions_data_sorted:
            # Group by theme
            if question['Theme'] != current_theme:
                current_theme = question['Theme']
                st.markdown(f"#### 🎯 {current_theme}")

            # Determine card style based on score (try to parse score for styling)
            card_style = "info"
            score_display = ""

            # Try to determine numeric value for styling
            try:
                if question['Score'] != "No score":
                    score_numeric = pd.to_numeric(question['Score'], errors='coerce')
                    if pd.notna(score_numeric):
                        if score_numeric >= 4.0:
                            card_style = "success"
                        elif score_numeric <= 2.5:
                            card_style = "warning"
                score_display = f"<div style='font-size: 1.2rem; font-weight: 700; color: #667eea; margin: 0.5rem 0;'>📊 Score: {question['Score']}</div>"
            except:
                score_display = f"<div style='font-size: 1rem; color: #64748b; margin: 0.5rem 0;'>📊 {question['Score']}</div>"

            st.markdown(f"""
            <div class="tab-info-card {card_style}">
                <div style="font-weight: 700; font-size: 1.1rem; margin-bottom: 0.5rem; color: #2c3e50;">
                    💬 {question['Affirmation']}
                </div>
                {score_display}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No valid question data found for analysis.")

def show_company_wide_analysis(team_data):
    """Show company-wide analysis across all teams"""
    st.markdown("""
    <div class="section-header fade-in-up">
        <div class="section-title">🌍 Company-Wide Analysis</div>
        <div class="section-subtitle">Consolidated insights across all teams</div>
    </div>
    """, unsafe_allow_html=True)

    # Aggregate data from all teams
    all_themes_data = []
    all_questions_data = []
    all_comments_data = []

    team_names = ["Andrew's Team", 'Build Team', 'People and Marketing Team', 'Finance and Operations Team']

    for team_name in team_names:
        if team_name in team_data:
            # Collect themes data
            for file_info in team_data[team_name].get('Themes', []):
                if 'data' in file_info and not file_info['data'].empty:
                    df_copy = file_info['data'].copy()
                    df_copy['_team_source'] = team_name
                    all_themes_data.append(df_copy)

            # Collect questions data
            for file_info in team_data[team_name].get('Questions', []):
                if 'data' in file_info and not file_info['data'].empty:
                    df_copy = file_info['data'].copy()
                    df_copy['_team_source'] = team_name
                    all_questions_data.append(df_copy)

            # Collect comments data
            for file_info in team_data[team_name].get('Comments', []):
                if 'data' in file_info and not file_info['data'].empty:
                    df_copy = file_info['data'].copy()
                    df_copy['_team_source'] = team_name
                    all_comments_data.append(df_copy)

    # Show overview metrics
    st.markdown("### 📊 Company Overview")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        active_teams = 0
        for t in team_names:
            if t in team_data and any(len(section_files) > 0 for section_files in team_data[t].values()):
                active_teams += 1
        st.metric("Active Teams", active_teams)
    with col2:
        st.metric("Theme Files", len(all_themes_data))
    with col3:
        st.metric("Question Files", len(all_questions_data))
    with col4:
        st.metric("Comment Files", len(all_comments_data))

    # Show analysis sections
    st.markdown("---")

    # Export options for company-wide analysis
    st.markdown("### 📄 Export Options")

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("📋 Copy Report Text", help="Generate text report to copy/paste"):
            report_content = generate_company_wide_narrative_report(team_data)
            st.text_area(
                "Copy this report to Google Docs:",
                value=report_content,
                height=400,
                help="Copy all text and paste into a new Google Doc"
            )

    with col2:
        if st.button("💾 Download Report", help="Download report as text file"):
            report_content = generate_company_wide_narrative_report(team_data)
            report_title = f"Company-Wide Survey Analysis - {pd.Timestamp.now().strftime('%B %d, %Y')}"

            # Create downloadable file
            st.download_button(
                label="📄 Download as .txt",
                data=report_content,
                file_name=f"company_wide_analysis_{pd.Timestamp.now().strftime('%Y%m%d')}.txt",
                mime="text/plain",
                help="Download and then copy/paste into Google Docs"
            )

    # Create tabs for different analysis types
    company_tabs = st.tabs(["🎯 Themes Summary", "❓ Questions Summary", "💬 Comments Summary"])

    with company_tabs[0]:
        show_company_wide_themes(all_themes_data, team_names)

    with company_tabs[1]:
        show_company_wide_questions(all_questions_data, team_names)

    with company_tabs[2]:
        show_company_wide_comments(all_comments_data, team_names)

def show_company_wide_themes(all_themes_data, team_names):
    """Show consolidated themes analysis"""
    if not all_themes_data:
        st.info("No themes data available across teams.")
        return

    st.markdown("#### 🎯 Themes Analysis Across All Teams")

    # Combine all themes data
    combined_themes = pd.concat(all_themes_data, ignore_index=True) if all_themes_data else pd.DataFrame()

    if combined_themes.empty:
        st.info("No themes data to analyze.")
        return

    # Look for theme and score columns
    theme_col = None
    score_col = None

    for col in combined_themes.columns:
        if any(word in col.lower() for word in ['theme', 'category', 'domain']):
            theme_col = col
            break

    for col in combined_themes.columns:
        if any(word in col.lower() for word in ['score', 'rating', 'value']) and not col.startswith('_'):
            score_col = col
            break

    if theme_col and score_col:
        try:
            # Convert scores to numeric
            combined_themes['Score_Numeric'] = pd.to_numeric(combined_themes[score_col], errors='coerce')
            valid_data = combined_themes.dropna(subset=['Score_Numeric', theme_col])

            if not valid_data.empty:
                # Calculate average scores by theme
                theme_scores = valid_data.groupby(theme_col)['Score_Numeric'].agg(['mean', 'count']).round(2)
                theme_scores.columns = ['Average Score', 'Response Count']
                theme_scores = theme_scores.sort_values('Average Score', ascending=False)

                st.markdown("**Top Performing Themes:**")
                for theme, row in theme_scores.head(5).iterrows():
                    score_color = "#10b981" if row['Average Score'] >= 4.0 else "#f59e0b" if row['Average Score'] <= 2.5 else "#3b82f6"
                    st.markdown(f"""
                    <div class="tab-info-card info">
                        <strong>{theme}</strong><br>
                        <span style="color: {score_color}; font-weight: 700;">📊 {row['Average Score']:.2f}</span>
                        <small style="color: #64748b;">({row['Response Count']} responses)</small>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown("**Low Performing Themes:**")
                for theme, row in theme_scores.tail(5).iterrows():
                    score_color = "#ef4444" if row['Average Score'] <= 2.5 else "#f59e0b" if row['Average Score'] <= 3.5 else "#3b82f6"
                    st.markdown(f"""
                    <div class="tab-info-card warning">
                        <strong>{theme}</strong><br>
                        <span style="color: {score_color}; font-weight: 700;">📊 {row['Average Score']:.2f}</span>
                        <small style="color: #64748b;">({row['Response Count']} responses)</small>
                    </div>
                    """, unsafe_allow_html=True)

        except Exception as e:
            st.warning(f"Could not process themes data: {str(e)}")
    else:
        st.info("Themes data structure not recognized. Expected columns for 'Theme' and 'Score'.")

def show_company_wide_questions(all_questions_data, team_names):
    """Show consolidated questions analysis"""
    if not all_questions_data:
        st.info("No questions data available across teams.")
        return

    st.markdown("#### ❓ Questions Analysis Across All Teams")

    # Combine all questions data
    combined_questions = pd.concat(all_questions_data, ignore_index=True) if all_questions_data else pd.DataFrame()

    if combined_questions.empty:
        st.info("No questions data to analyze.")
        return

    # Use the analyze_questions_data function
    questions_data, question_col, theme_col, score_col = analyze_questions_data(combined_questions)

    if questions_data:
        # Group by theme and show summary
        theme_summary = {}
        for item in questions_data:
            theme = item.get('Theme', 'Not specified')
            if theme not in theme_summary:
                theme_summary[theme] = {'count': 0, 'scores': []}
            theme_summary[theme]['count'] += 1
            if item.get('Score') and item['Score'] != "No score":
                try:
                    score_numeric = pd.to_numeric(item['Score'], errors='coerce')
                    if pd.notna(score_numeric):
                        theme_summary[theme]['scores'].append(score_numeric)
                except:
                    pass

        st.markdown("**Question Themes Summary (Ordered by Score):**")
        # Sort by average score (lowest to highest)
        themes_with_scores = []
        for theme, data in theme_summary.items():
            avg_score = sum(data['scores']) / len(data['scores']) if data['scores'] else 0
            themes_with_scores.append((theme, data, avg_score))

        # Sort by score (lowest first)
        sorted_themes = sorted(themes_with_scores, key=lambda x: x[2] if x[2] > 0 else 999)

        for theme, data, avg_score in sorted_themes:
            score_display = f"📊 {avg_score:.2f}" if avg_score > 0 else "📊 No scores"
            score_color = "#ef4444" if avg_score <= 2.5 else "#f59e0b" if avg_score <= 3.5 else "#10b981" if avg_score >= 4.0 else "#3b82f6"
            card_class = "warning" if avg_score <= 3.0 else "info"
            st.markdown(f"""
            <div class="tab-info-card {card_class}">
                <strong>{theme}</strong><br>
                <span style="color: {score_color}; font-weight: 700;">{score_display}</span> <small>({data['count']} questions)</small>
            </div>
            """, unsafe_allow_html=True)

        st.markdown(f"**Total Questions Analyzed:** {len(questions_data)} across all teams")
    else:
        st.info("Questions data structure not recognized. Expected columns for 'Question', 'Theme', and 'Score'.")

def show_company_wide_comments(all_comments_data, team_names):
    """Show consolidated comments analysis"""
    if not all_comments_data:
        st.info("No comments data available across teams.")
        return

    st.markdown("#### 💬 Comments Analysis Across All Teams")

    # Combine all comments data
    combined_comments = pd.concat(all_comments_data, ignore_index=True) if all_comments_data else pd.DataFrame()

    if combined_comments.empty:
        st.info("No comments data to analyze.")
        return

    # Use the analyze_comments_data function
    comments_data, question_col, theme_col, comment_col = analyze_comments_data(combined_comments)

    if comments_data:
        # Collect all comments for sentiment analysis
        all_comments = [item['Comment'] for item in comments_data if item.get('Comment') and item['Comment'] != "No comment"]

        if all_comments:
            # Perform company-wide sentiment analysis
            sentiment_analysis = analyze_comment_sentiment(all_comments)

            # Display high-level metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Comments", sentiment_analysis['total_comments'])
            with col2:
                positive_pct = sentiment_analysis['positive_count'] / sentiment_analysis['total_comments'] * 100
                st.metric("Positive Sentiment", f"{positive_pct:.1f}%")
            with col3:
                st.metric("Key Themes", len(sentiment_analysis['key_themes']))

            # Generate comprehensive company-wide insights
            st.markdown("**Overall Comments Insights:**")

            # Create thematic insights based on sentiment analysis
            insights = generate_comprehensive_insights(sentiment_analysis, all_comments)

            for insight in insights:
                st.markdown(f"**{insight['title']}:** {insight['description']}")
                st.markdown("")
        else:
            st.info("No comment content found for analysis.")
    else:
        # Fallback to column-based analysis
        st.info("Comments data structure not recognized. Expected columns for 'Comment', 'Theme', etc.")

def main():
    st.set_page_config(
        page_title="AI Fund Survey Results",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed"
    )

    # Add custom CSS
    add_custom_css()

    # Main container
    st.markdown('<div class="main-container">', unsafe_allow_html=True)

    # Hero Section
    st.markdown("""
    <div class="hero-section">
        <div class="hero-title">AI Fund Survey Results</div>
        <div class="hero-subtitle">Team-based survey analysis with themes, questions, and comments</div>
    </div>
    """, unsafe_allow_html=True)

    # Multiple file upload section
    st.markdown("""
    <div class="upload-section fade-in-up">
        <div class="upload-icon">📊</div>
        <div class="upload-title">Upload Your Survey Files</div>
        <div class="upload-description">Upload multiple Excel files to automatically organize by teams and categories<br>
        <small style="color: #94a3b8;">File naming examples: "andrew_themes.xlsx", "build_comments.xlsx", "people_questions.xlsx", "finance_themes.xlsx"</small></div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Upload multiple Excel files - they will be automatically categorized by team and type"
    )

    if uploaded_files:
        team_data, files_processed, processing_errors = load_and_process_multiple_files(uploaded_files)

        # Show processing errors if any
        if processing_errors:
            for error in processing_errors:
                st.markdown(f"""
                <div class="error-card">
                    <strong>⚠️ Processing Warning:</strong> {error}
                </div>
                """, unsafe_allow_html=True)

        # Show file categorization summary
        total_files = len(uploaded_files)
        successful_files = len(files_processed)

        st.markdown(f"""
        <div class="success-card fade-in-up">
            <strong>✅ {successful_files}/{total_files} files processed successfully!</strong><br>
            Files automatically categorized by teams and analysis type
        </div>
        """, unsafe_allow_html=True)

        # Show categorization summary
        if files_processed:
            with st.expander("📁 File Categorization Summary", expanded=False):
                for file_info in files_processed:
                    st.markdown(f"• **{file_info['filename']}** → {file_info['team']} - {file_info['section']}")

        # Navigation Layout - Two columns
        nav_col1, nav_col2 = st.columns([1, 1])

        with nav_col1:
            st.markdown("## 🏢 Company-Wide Analysis")
            show_company_wide = st.button(
                "📊 View All Teams Summary",
                key="company_wide_analysis",
                help="View consolidated analysis across all teams",
                width='stretch'
            )

        with nav_col2:
            st.markdown("## 🏢 Select Team")

            # Team Navigation - Primary navigation
            team_names = ["Andrew's Team", 'Build Team', 'People and Marketing Team', 'Finance and Operations Team']

            # Add a default option to prevent auto-selection
            team_options = ['Select a team...'] + team_names
            selected_team_index = st.radio(
                "Choose a team to analyze:",
                range(len(team_options)),
                format_func=lambda x: team_options[x],
                horizontal=False,
                label_visibility="collapsed",
                index=0  # Start with "Select a team..." option
            )

            selected_team = team_options[selected_team_index] if selected_team_index > 0 else None

        # Show content based on selection
        if show_company_wide:
            # Show company-wide analysis
            show_company_wide_analysis(team_data)
        elif selected_team is None:
            # Welcome screen when no team is selected
            st.markdown("""
            <div class="welcome-section">
                <div class="welcome-title">👋 Welcome to AI Fund Survey Results</div>
                <div class="welcome-subtitle">Select a team above to start analyzing survey results</div>
                <div style="margin-top: 2rem;">
                    <h4>📊 Available Teams</h4>
            """, unsafe_allow_html=True)

            # Show team overview cards
            cols = st.columns(2)
            for i, team in enumerate(team_names):
                with cols[i % 2]:
                    team_sections = team_data[team]
                    total_files = sum(len(files) for files in team_sections.values())
                    available_categories = [section for section, files in team_sections.items() if files]

                    if total_files > 0:
                        status_color = "#10b981"  # green
                        status_text = f"✅ {total_files} file{'s' if total_files > 1 else ''}"
                        categories_text = ", ".join(available_categories) if available_categories else "No data"
                    else:
                        status_color = "#64748b"  # gray
                        status_text = "📝 No data"
                        categories_text = "Upload files to get started"

                    st.markdown(f"""
                    <div class="team-selection-card" style="margin: 0.5rem 0;">
                        <div class="team-name">{team}</div>
                        <div style="color: {status_color}; font-weight: 600; margin: 0.5rem 0;">{status_text}</div>
                        <div class="team-description">{categories_text}</div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("</div></div>", unsafe_allow_html=True)

        else:
            # Show team analysis when a team is selected
            st.markdown(f"""
            <div class="section-header fade-in-up">
                <div class="section-title">📊 {selected_team} Analysis</div>
                <div class="section-subtitle">Choose an analysis category below</div>
            </div>
            """, unsafe_allow_html=True)

            # Show data availability for selected team
            team_sections = team_data[selected_team]
            available_sections = []
            for section, files in team_sections.items():
                if files:
                    available_sections.append(f"{section} ({len(files)} file{'s' if len(files) > 1 else ''})")

            if available_sections:
                st.markdown(f"**📁 Available Data:** {', '.join(available_sections)}")
            else:
                st.info(f"No data files found for {selected_team}. Upload files with team keywords in the filename.")

            # Secondary navigation - Analysis categories within team
            if any(team_data[selected_team].values()):  # If team has any data
                st.markdown("### 📈 Analysis Categories")

                # Export options for individual team
                st.markdown("#### 📄 Export Options")

                col1, col2 = st.columns([1, 1])
                with col1:
                    if st.button(f"📋 Copy {selected_team} Text", help="Generate text report to copy/paste"):
                        report_content = generate_team_narrative_report(selected_team, team_data[selected_team])
                        st.text_area(
                            f"Copy this {selected_team} report to Google Docs:",
                            value=report_content,
                            height=400,
                            help="Copy all text and paste into a new Google Doc"
                        )

                with col2:
                    if st.button(f"💾 Download {selected_team} Report", help="Download report as text file"):
                        report_content = generate_team_narrative_report(selected_team, team_data[selected_team])

                        # Create downloadable file
                        safe_team_name = selected_team.replace(" ", "_").replace("'", "").lower()
                        st.download_button(
                            label="📄 Download as .txt",
                            data=report_content,
                            file_name=f"{safe_team_name}_analysis_{pd.Timestamp.now().strftime('%Y%m%d')}.txt",
                            mime="text/plain",
                            help="Download and then copy/paste into Google Docs"
                        )

                analysis_tabs = st.tabs(["🎯 Themes", "❓ Questions", "💬 Comments"])

                with analysis_tabs[0]:
                    show_team_themes_analysis(team_data, selected_team)

                with analysis_tabs[1]:
                    show_team_questions_analysis(team_data, selected_team)

                with analysis_tabs[2]:
                    show_team_comments_analysis(team_data, selected_team)

    else:
        # Initial upload prompt with team examples
        st.markdown("""
        <div class="info-card fade-in-up">
            <strong>💡 Smart Team File Organization:</strong><br>
            Name your files to automatically assign them to teams:<br><br>
            <strong>Team Keywords:</strong><br>
            <small>• "andrew" → Andrew's Team</small><br>
            <small>• "build" → Build Team</small><br>
            <small>• "people" or "marketing" → People and Marketing Team</small><br>
            <small>• "operations" or "finance" → Finance and Operations Team</small><br><br>
            <strong>Category Keywords:</strong><br>
            <small>• "themes", "topics" → Themes Analysis</small><br>
            <small>• "questions", "survey" → Questions Analysis</small><br>
            <small>• "comments", "feedback" → Comments Analysis</small><br><br>
            <strong>Example filenames:</strong><br>
            <small>• "andrew_themes.xlsx", "build_comments.xlsx", "people_questions.xlsx", "finance_themes.xlsx"</small>
        </div>
        """, unsafe_allow_html=True)

    # Close main container
    st.markdown('</div>', unsafe_allow_html=True)

def analyze_survey_data(df):
    """Analyze survey data focusing on themes, scores, and participation rates"""
    insights = {
        'total_themes': len(df),
        'themes': [],
        'score_analysis': {},
        'participation_analysis': {},
        'combined_insights': {}
    }

    # Analyze themes (Column A)
    if 'Theme' in df.columns:
        theme_data = df['Theme'].dropna().astype(str)
        insights['themes'] = theme_data.tolist()
        insights['unique_themes'] = len(theme_data.unique())

    # Analyze scores (Column B)
    if 'Score' in df.columns:
        score_data = pd.to_numeric(df['Score'], errors='coerce').dropna()
        if len(score_data) > 0:
            insights['score_analysis'] = {
                'mean': round(score_data.mean(), 2),
                'median': round(score_data.median(), 2),
                'min': score_data.min(),
                'max': score_data.max(),
                'std': round(score_data.std(), 2) if len(score_data) > 1 else 0,
                'count': len(score_data),
                'high_scores': len(score_data[score_data >= score_data.quantile(0.75)]),
                'low_scores': len(score_data[score_data <= score_data.quantile(0.25)])
            }

    # Analyze participation rates (Column C)
    if 'Participation_Rate' in df.columns:
        # Handle both percentage format (85%) and decimal format (0.85)
        participation_data = df['Participation_Rate'].dropna().astype(str)
        numeric_participation = []

        for val in participation_data:
            try:
                if '%' in str(val):
                    # Convert percentage to decimal
                    numeric_val = float(str(val).replace('%', '')) / 100
                else:
                    numeric_val = float(val)
                    # If value is greater than 1, assume it's a percentage
                    if numeric_val > 1:
                        numeric_val = numeric_val / 100
                numeric_participation.append(numeric_val)
            except:
                continue

        if numeric_participation:
            participation_series = pd.Series(numeric_participation)
            insights['participation_analysis'] = {
                'mean': round(participation_series.mean(), 3),
                'median': round(participation_series.median(), 3),
                'min': round(participation_series.min(), 3),
                'max': round(participation_series.max(), 3),
                'count': len(participation_series),
                'high_participation': len(participation_series[participation_series >= 0.75]),
                'low_participation': len(participation_series[participation_series <= 0.5])
            }

    # Combined analysis - correlate scores with participation
    if 'Score' in df.columns and 'Participation_Rate' in df.columns:
        valid_rows = df[['Score', 'Participation_Rate', 'Theme']].dropna()
        if len(valid_rows) > 0:
            # Find best and worst performing themes
            score_data = pd.to_numeric(valid_rows['Score'], errors='coerce')
            valid_rows_with_scores = valid_rows[score_data.notna()].copy()

            if len(valid_rows_with_scores) > 0:
                valid_rows_with_scores['Score_Numeric'] = pd.to_numeric(valid_rows_with_scores['Score'], errors='coerce')
                best_theme = valid_rows_with_scores.loc[valid_rows_with_scores['Score_Numeric'].idxmax()]
                worst_theme = valid_rows_with_scores.loc[valid_rows_with_scores['Score_Numeric'].idxmin()]

                insights['combined_insights'] = {
                    'best_theme': {
                        'name': best_theme['Theme'],
                        'score': best_theme['Score_Numeric'],
                        'participation': best_theme['Participation_Rate']
                    },
                    'worst_theme': {
                        'name': worst_theme['Theme'],
                        'score': worst_theme['Score_Numeric'],
                        'participation': worst_theme['Participation_Rate']
                    }
                }

    return insights

def generate_survey_narrative(insights):
    """Generate a narrative report focused on survey themes, scores, and participation"""
    narrative = []

    # Introduction
    narrative.append("## Survey Theme Analysis Report")
    narrative.append("")
    narrative.append(f"This analysis examines **{insights['total_themes']} survey themes** with their corresponding scores and participation rates.")
    narrative.append("")

    # Score Analysis
    if insights.get('score_analysis'):
        score_data = insights['score_analysis']
        narrative.append("### 📊 Score Performance Analysis")
        narrative.append("")
        narrative.append(f"**Overall Score Metrics:**")
        narrative.append(f"- Average Score: **{score_data['mean']}** (out of {score_data['max']})")
        narrative.append(f"- Score Range: {score_data['min']} to {score_data['max']}")
        narrative.append(f"- Median Score: {score_data['median']}")
        narrative.append("")

        # Performance categorization
        high_pct = round((score_data['high_scores'] / score_data['count']) * 100, 1)
        low_pct = round((score_data['low_scores'] / score_data['count']) * 100, 1)

        narrative.append(f"**Performance Distribution:**")
        narrative.append(f"- High-performing themes (top 25%): **{score_data['high_scores']} themes** ({high_pct}%)")
        narrative.append(f"- Low-performing themes (bottom 25%): **{score_data['low_scores']} themes** ({low_pct}%)")
        narrative.append("")

    # Participation Analysis
    if insights.get('participation_analysis'):
        part_data = insights['participation_analysis']
        narrative.append("### 👥 Participation Rate Analysis")
        narrative.append("")
        narrative.append(f"**Participation Metrics:**")
        narrative.append(f"- Average Participation: **{part_data['mean']*100:.1f}%**")
        narrative.append(f"- Participation Range: {part_data['min']*100:.1f}% to {part_data['max']*100:.1f}%")
        narrative.append(f"- Median Participation: {part_data['median']*100:.1f}%")
        narrative.append("")

        high_part_pct = round((part_data['high_participation'] / part_data['count']) * 100, 1)
        low_part_pct = round((part_data['low_participation'] / part_data['count']) * 100, 1)

        narrative.append(f"**Engagement Levels:**")
        narrative.append(f"- High engagement themes (≥75% participation): **{part_data['high_participation']} themes** ({high_part_pct}%)")
        narrative.append(f"- Low engagement themes (≤50% participation): **{part_data['low_participation']} themes** ({low_part_pct}%)")
        narrative.append("")

    # Combined Insights
    if insights.get('combined_insights'):
        combined = insights['combined_insights']
        narrative.append("### 🏆 Theme Performance Highlights")
        narrative.append("")

        if 'best_theme' in combined:
            best = combined['best_theme']
            narrative.append(f"**Top Performing Theme:**")
            narrative.append(f"- **{best['name']}**")
            narrative.append(f"- Score: {best['score']}")
            narrative.append(f"- Participation: {best['participation']}")
            narrative.append("")

        if 'worst_theme' in combined:
            worst = combined['worst_theme']
            narrative.append(f"**Lowest Performing Theme:**")
            narrative.append(f"- **{worst['name']}**")
            narrative.append(f"- Score: {worst['score']}")
            narrative.append(f"- Participation: {worst['participation']}")
            narrative.append("")

    # Summary and Recommendations
    narrative.append("### 🎯 Key Insights & Recommendations")
    narrative.append("")

    if insights.get('score_analysis') and insights.get('participation_analysis'):
        score_avg = insights['score_analysis']['mean']
        part_avg = insights['participation_analysis']['mean']

        if score_avg >= 7 and part_avg >= 0.7:
            narrative.append("🟢 **Strong Performance:** Both scores and participation rates show positive trends.")
        elif score_avg >= 6 and part_avg >= 0.6:
            narrative.append("🟡 **Moderate Performance:** Scores and participation are at acceptable levels with room for improvement.")
        else:
            narrative.append("🔴 **Improvement Needed:** Focus on strategies to boost both scores and engagement.")

        narrative.append("")
        narrative.append("**Recommendations:**")

        if part_avg < 0.6:
            narrative.append("- **Increase Engagement:** Focus on themes with low participation rates")

        if score_avg < 6:
            narrative.append("- **Improve Content Quality:** Address low-scoring themes")

        narrative.append("- **Monitor Trends:** Track performance over time to identify patterns")

    return "\n".join(narrative)

def create_survey_visualizations(df, insights):
    """Create visualizations for survey themes, scores, and participation"""
    charts = []

    try:
        # Score distribution chart
        if 'Score' in df.columns:
            score_data = pd.to_numeric(df['Score'], errors='coerce').dropna()
            if len(score_data) > 0:
                fig = px.histogram(
                    x=score_data,
                    title="📊 Score Distribution",
                    nbins=min(10, len(score_data.unique())),
                    labels={'x': 'Score', 'y': 'Number of Themes'}
                )
                fig.update_layout(height=400, showlegend=False)
                charts.append(('Score Distribution', fig))

        # Participation rate distribution
        if 'Participation_Rate' in df.columns:
            # Convert participation rates to percentages
            participation_data = df['Participation_Rate'].dropna().astype(str)
            numeric_participation = []

            for val in participation_data:
                try:
                    if '%' in str(val):
                        numeric_val = float(str(val).replace('%', ''))
                    else:
                        numeric_val = float(val)
                        if numeric_val <= 1:
                            numeric_val = numeric_val * 100
                    numeric_participation.append(numeric_val)
                except:
                    continue

            if numeric_participation:
                fig = px.histogram(
                    x=numeric_participation,
                    title="👥 Participation Rate Distribution",
                    nbins=min(10, len(set(numeric_participation))),
                    labels={'x': 'Participation Rate (%)', 'y': 'Number of Themes'}
                )
                fig.update_layout(height=400, showlegend=False)
                charts.append(('Participation Distribution', fig))

        # Score vs Participation scatter plot
        if 'Score' in df.columns and 'Participation_Rate' in df.columns and 'Theme' in df.columns:
            # Prepare data for scatter plot
            scatter_data = df[['Theme', 'Score', 'Participation_Rate']].dropna()
            scores = pd.to_numeric(scatter_data['Score'], errors='coerce')
            scatter_data_clean = scatter_data[scores.notna()].copy()

            if len(scatter_data_clean) > 0:
                scatter_data_clean['Score_Numeric'] = pd.to_numeric(scatter_data_clean['Score'])

                # Convert participation to numeric
                participation_numeric = []
                for val in scatter_data_clean['Participation_Rate']:
                    try:
                        if '%' in str(val):
                            numeric_val = float(str(val).replace('%', ''))
                        else:
                            numeric_val = float(val)
                            if numeric_val <= 1:
                                numeric_val = numeric_val * 100
                        participation_numeric.append(numeric_val)
                    except:
                        participation_numeric.append(None)

                scatter_data_clean['Participation_Numeric'] = participation_numeric
                scatter_data_clean = scatter_data_clean[scatter_data_clean['Participation_Numeric'].notna()]

                if len(scatter_data_clean) > 0:
                    fig = px.scatter(
                        scatter_data_clean,
                        x='Participation_Numeric',
                        y='Score_Numeric',
                        hover_data=['Theme'],
                        title="🎯 Score vs Participation Rate by Theme",
                        labels={
                            'Participation_Numeric': 'Participation Rate (%)',
                            'Score_Numeric': 'Score'
                        }
                    )
                    fig.update_layout(height=500)
                    charts.append(('Score vs Participation', fig))

        # Top themes by score (if enough themes)
        if len(df) > 0 and 'Theme' in df.columns and 'Score' in df.columns:
            theme_scores = df[['Theme', 'Score']].dropna()
            numeric_scores = pd.to_numeric(theme_scores['Score'], errors='coerce')
            theme_scores_clean = theme_scores[numeric_scores.notna()].copy()
            theme_scores_clean['Score_Numeric'] = pd.to_numeric(theme_scores_clean['Score'])

            if len(theme_scores_clean) >= 3:
                top_themes = theme_scores_clean.nlargest(min(10, len(theme_scores_clean)), 'Score_Numeric')

                fig = px.bar(
                    top_themes,
                    x='Score_Numeric',
                    y='Theme',
                    orientation='h',
                    title=f"🏆 Top {len(top_themes)} Themes by Score",
                    labels={'Score_Numeric': 'Score', 'Theme': 'Theme'}
                )
                fig.update_layout(height=max(400, len(top_themes) * 30))
                charts.append(('Top Themes', fig))

    except Exception as e:
        st.warning(f"Could not create some visualizations: {str(e)}")

    return charts

def create_pdf_report(insights, df):
    """Create a PDF narrative report"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18)

    # Get styles
    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )

    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        spaceBefore=20,
        textColor=colors.darkblue
    )

    subheading_style = ParagraphStyle(
        'CustomSubHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=8,
        spaceBefore=15,
        textColor=colors.navy
    )

    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=11,
        spaceAfter=8,
        alignment=TA_JUSTIFY,
        leftIndent=12
    )

    # Build the PDF content
    story = []

    # Title page
    story.append(Paragraph("Survey Theme Analysis Report", title_style))
    story.append(Spacer(1, 0.5*inch))
    story.append(Paragraph(f"Generated on: {datetime.now().strftime('%B %d, %Y')}", styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(f"Total Themes Analyzed: {insights['total_themes']}", styles['Normal']))
    story.append(PageBreak())

    # Executive Summary
    story.append(Paragraph("Executive Summary", heading_style))

    executive_summary = f"""
    This comprehensive analysis examines {insights['total_themes']} survey themes, evaluating their performance
    through score metrics and participation rates. The report provides insights into theme effectiveness,
    engagement levels, and actionable recommendations for improvement.
    """
    story.append(Paragraph(executive_summary, normal_style))
    story.append(Spacer(1, 0.3*inch))

    # Score Analysis Section
    if insights.get('score_analysis'):
        score_data = insights['score_analysis']
        story.append(Paragraph("Score Performance Analysis", heading_style))

        score_summary = f"""
        The overall performance metrics reveal an average score of <b>{score_data['mean']}</b> across all themes,
        with scores ranging from {score_data['min']} to {score_data['max']}. The median score of {score_data['median']}
        indicates a {'positive' if score_data['median'] >= 6 else 'moderate'} distribution of theme performance.
        """
        story.append(Paragraph(score_summary, normal_style))

        # Performance breakdown
        high_pct = round((score_data['high_scores'] / score_data['count']) * 100, 1)
        low_pct = round((score_data['low_scores'] / score_data['count']) * 100, 1)

        performance_text = f"""
        <b>Performance Distribution:</b><br/>
        • High-performing themes (top 25%): {score_data['high_scores']} themes ({high_pct}%)<br/>
        • Low-performing themes (bottom 25%): {score_data['low_scores']} themes ({low_pct}%)<br/>
        • Standard deviation: {score_data['std']}, indicating {'low' if score_data['std'] < 1 else 'moderate' if score_data['std'] < 2 else 'high'} variability in scores
        """
        story.append(Paragraph(performance_text, normal_style))
        story.append(Spacer(1, 0.2*inch))

    # Participation Analysis Section
    if insights.get('participation_analysis'):
        part_data = insights['participation_analysis']
        story.append(Paragraph("Participation Rate Analysis", heading_style))

        participation_summary = f"""
        Engagement analysis shows an average participation rate of <b>{part_data['mean']*100:.1f}%</b>,
        with participation ranging from {part_data['min']*100:.1f}% to {part_data['max']*100:.1f}%.
        The median participation rate of {part_data['median']*100:.1f}% suggests
        {'strong' if part_data['median'] >= 0.7 else 'moderate' if part_data['median'] >= 0.5 else 'low'} overall engagement.
        """
        story.append(Paragraph(participation_summary, normal_style))

        # Engagement breakdown
        high_part_pct = round((part_data['high_participation'] / part_data['count']) * 100, 1)
        low_part_pct = round((part_data['low_participation'] / part_data['count']) * 100, 1)

        engagement_text = f"""
        <b>Engagement Distribution:</b><br/>
        • High engagement themes (≥75% participation): {part_data['high_participation']} themes ({high_part_pct}%)<br/>
        • Low engagement themes (≤50% participation): {part_data['low_participation']} themes ({low_part_pct}%)<br/>
        • Remaining themes show moderate engagement levels
        """
        story.append(Paragraph(engagement_text, normal_style))
        story.append(Spacer(1, 0.2*inch))

    # Theme Highlights Section
    if insights.get('combined_insights'):
        combined = insights['combined_insights']
        story.append(Paragraph("Theme Performance Highlights", heading_style))

        if 'best_theme' in combined:
            best = combined['best_theme']
            best_text = f"""
            <b>Top Performing Theme:</b><br/>
            <b>{best['name']}</b><br/>
            Score: {best['score']} | Participation: {best['participation']}<br/>
            This theme demonstrates excellent performance with high scores and strong engagement.
            """
            story.append(Paragraph(best_text, normal_style))
            story.append(Spacer(1, 0.1*inch))

        if 'worst_theme' in combined:
            worst = combined['worst_theme']
            worst_text = f"""
            <b>Area for Improvement:</b><br/>
            <b>{worst['name']}</b><br/>
            Score: {worst['score']} | Participation: {worst['participation']}<br/>
            This theme requires attention to improve both content quality and engagement levels.
            """
            story.append(Paragraph(worst_text, normal_style))
        story.append(Spacer(1, 0.2*inch))

    # Detailed Theme Table
    if len(df) > 0:
        story.append(PageBreak())
        story.append(Paragraph("Detailed Theme Breakdown", heading_style))

        # Prepare table data
        table_data = [['Theme', 'Score', 'Participation Rate']]

        for _, row in df.iterrows():
            theme_name = str(row['Theme']) if pd.notna(row['Theme']) else 'N/A'
            score = str(row['Score']) if pd.notna(row['Score']) else 'N/A'

            # Format participation rate
            participation = 'N/A'
            if pd.notna(row['Participation_Rate']):
                try:
                    if '%' in str(row['Participation_Rate']):
                        participation = str(row['Participation_Rate'])
                    else:
                        num_val = float(row['Participation_Rate'])
                        if num_val <= 1:
                            participation = f"{num_val*100:.1f}%"
                        else:
                            participation = f"{num_val:.1f}%"
                except:
                    participation = str(row['Participation_Rate'])

            table_data.append([theme_name, score, participation])

        # Create table
        table = Table(table_data, colWidths=[3*inch, 1*inch, 1.5*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))

        story.append(table)
        story.append(Spacer(1, 0.3*inch))

    # Recommendations Section
    story.append(PageBreak())
    story.append(Paragraph("Strategic Recommendations", heading_style))

    recommendations = []

    if insights.get('score_analysis') and insights.get('participation_analysis'):
        score_avg = insights['score_analysis']['mean']
        part_avg = insights['participation_analysis']['mean']

        if score_avg >= 7 and part_avg >= 0.7:
            recommendations.append("Continue current strategies as both scores and participation show strong performance.")
            recommendations.append("Consider scaling successful approaches to other survey initiatives.")
        elif score_avg >= 6 and part_avg >= 0.6:
            recommendations.append("Focus on incremental improvements to elevate performance from moderate to strong levels.")
            recommendations.append("Identify and replicate best practices from top-performing themes.")
        else:
            recommendations.append("Implement comprehensive improvement strategies to address both content quality and engagement.")
            recommendations.append("Consider redesigning low-performing themes based on high-performer analysis.")

        if part_avg < 0.6:
            recommendations.append("Develop targeted engagement strategies for themes with low participation rates.")
            recommendations.append("Investigate barriers to participation and implement solutions.")

        if score_avg < 6:
            recommendations.append("Conduct content review and enhancement for low-scoring themes.")
            recommendations.append("Implement quality assurance processes for future theme development.")

    recommendations.extend([
        "Establish regular monitoring and evaluation cycles to track performance trends over time.",
        "Create feedback mechanisms to understand participant preferences and experiences.",
        "Consider A/B testing different approaches for underperforming themes."
    ])

    for i, rec in enumerate(recommendations, 1):
        story.append(Paragraph(f"{i}. {rec}", normal_style))
        story.append(Spacer(1, 0.1*inch))

    # Footer
    story.append(Spacer(1, 0.5*inch))
    footer_text = f"""
    <i>This report was automatically generated by the Survey Narrative Generator on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}.
    For questions or additional analysis, please consult with your survey administration team.</i>
    """
    story.append(Paragraph(footer_text, styles['Normal']))

    # Build PDF
    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()

    return pdf_bytes

def generate_survey_report(df, current_step=3):
    """Generate the complete survey themes report"""

    # Show step indicator
    create_step_indicator(current_step)

    with st.spinner("🔍 Analyzing your survey themes data..."):
        # Perform specialized survey analysis
        insights = analyze_survey_data(df)

        # Generate narrative
        narrative_text = generate_survey_narrative(insights)

        # Create visualizations
        charts = create_survey_visualizations(df, insights)

    # Display results with custom styling
    st.markdown("""
    <div class="success-card fade-in-up">
        <strong>🎉 Survey Theme Analysis Complete!</strong><br>
        Your comprehensive narrative report is ready below
    </div>
    """, unsafe_allow_html=True)

    # Show narrative with better formatting
    st.markdown('<div class="fade-in-up">', unsafe_allow_html=True)
    st.markdown("## 📝 Survey Theme Narrative Report")

    # Add some visual separation and better typography
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
                padding: 2rem; border-radius: 16px; margin: 1rem 0;
                border-left: 4px solid #667eea;">
        {narrative_text.replace('##', '<h3 style="color: #2c3e50; margin-top: 1.5rem; margin-bottom: 1rem;">').replace('###', '<h4 style="color: #475569; margin-top: 1rem; margin-bottom: 0.5rem;">').replace('**', '<strong>').replace('- ', '<li>').replace('\n\n', '</li></ul><br><ul style="list-style-type: none; padding-left: 0;"><li>')}
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Show visualizations with better spacing
    if charts:
        st.markdown('<div class="fade-in-up">', unsafe_allow_html=True)
        st.markdown("## 📊 Data Visualizations")

        st.markdown('<div style="margin: 2rem 0;">', unsafe_allow_html=True)
        for i, (chart_name, fig) in enumerate(charts):
            # Add some spacing between charts
            if i > 0:
                st.markdown('<div style="margin: 2rem 0;"></div>', unsafe_allow_html=True)
            st.plotly_chart(fig, width='stretch')
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Show detailed theme breakdown with better styling
    if len(df) > 0:
        st.markdown('<div class="fade-in-up">', unsafe_allow_html=True)
        st.markdown("## 📋 Detailed Theme Breakdown")

        # Create a clean display dataframe
        display_df = df.copy()

        # Format participation rates
        if 'Participation_Rate' in display_df.columns:
            formatted_participation = []
            for val in display_df['Participation_Rate']:
                try:
                    if pd.notna(val):
                        if '%' not in str(val):
                            num_val = float(val)
                            if num_val <= 1:
                                formatted_participation.append(f"{num_val*100:.1f}%")
                            else:
                                formatted_participation.append(f"{num_val:.1f}%")
                        else:
                            formatted_participation.append(str(val))
                    else:
                        formatted_participation.append("N/A")
                except:
                    formatted_participation.append("N/A")

            display_df['Participation_Rate'] = formatted_participation

        st.dataframe(
            display_df,
            width='stretch',
            column_config={
                "Theme": st.column_config.TextColumn("Theme", width="medium"),
                "Score": st.column_config.NumberColumn("Score", format="%.1f"),
                "Participation_Rate": st.column_config.TextColumn("Participation Rate")
            }
        )
        st.markdown('</div>', unsafe_allow_html=True)

    # Download options with better styling
    st.markdown('<div class="fade-in-up">', unsafe_allow_html=True)
    st.markdown("## 💾 Download Your Report")

    st.markdown("""
    <div class="download-section">
        <h4 style="color: #2c3e50; text-align: center; margin-bottom: 1.5rem;">Choose your preferred format</h4>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown("""
        <div class="metric-card" style="text-align: center;">
            <div style="font-size: 2rem; margin-bottom: 1rem;">📄</div>
            <h5 style="color: #2c3e50; margin-bottom: 0.5rem;">Markdown Report</h5>
            <p style="color: #64748b; font-size: 0.9rem; margin-bottom: 1.5rem;">Editable text format for customization</p>
        </div>
        """, unsafe_allow_html=True)

        # Create downloadable markdown content
        report_content = f"""# Survey Theme Analysis Report

{narrative_text}

## Data Summary
- Total Themes Analyzed: {insights['total_themes']}
"""

        if insights.get('score_analysis'):
            score_data = insights['score_analysis']
            report_content += f"""
## Score Analysis
- Average Score: {score_data['mean']}
- Score Range: {score_data['min']} - {score_data['max']}
- High-performing Themes: {score_data['high_scores']}
- Low-performing Themes: {score_data['low_scores']}
"""

        if insights.get('participation_analysis'):
            part_data = insights['participation_analysis']
            report_content += f"""
## Participation Analysis
- Average Participation: {part_data['mean']*100:.1f}%
- Participation Range: {part_data['min']*100:.1f}% - {part_data['max']*100:.1f}%
- High-engagement Themes: {part_data['high_participation']}
- Low-engagement Themes: {part_data['low_participation']}
"""

        report_content += f"""
## Theme Details
"""
        for _, row in df.iterrows():
            if pd.notna(row['Theme']):
                report_content += f"\n### {row['Theme']}\n"
                if pd.notna(row['Score']):
                    report_content += f"- Score: {row['Score']}\n"
                if pd.notna(row['Participation_Rate']):
                    report_content += f"- Participation Rate: {row['Participation_Rate']}\n"

        st.download_button(
            label="📄 Download Markdown",
            data=report_content,
            file_name="survey_theme_analysis_report.md",
            mime="text/markdown",
            width='stretch'
        )

    with col2:
        st.markdown("""
        <div class="metric-card" style="text-align: center;">
            <div style="font-size: 2rem; margin-bottom: 1rem;">📊</div>
            <h5 style="color: #2c3e50; margin-bottom: 0.5rem;">PDF Report</h5>
            <p style="color: #64748b; font-size: 0.9rem; margin-bottom: 1.5rem;">Professional narrative document</p>
        </div>
        """, unsafe_allow_html=True)

        # PDF Download
        try:
            with st.spinner("🎨 Generating beautiful PDF..."):
                pdf_bytes = create_pdf_report(insights, df)

            st.download_button(
                label="📊 Download PDF",
                data=pdf_bytes,
                file_name=f"survey_narrative_report_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                width='stretch'
            )
        except Exception as e:
            st.markdown(f"""
            <div style="background: #fef2f2; color: #dc2626; padding: 1rem; border-radius: 8px; text-align: center;">
                <strong>PDF Generation Error</strong><br>
                <small>{str(e)}</small>
            </div>
            """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
