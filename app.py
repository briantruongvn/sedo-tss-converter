#!/usr/bin/env python3
"""
Ngoc Son Internal TSS Converter - Streamlit Web Application
Convert Ngoc Son Internal TSS to Standard Internal TSS
"""

import streamlit as st
import pandas as pd
import io
import os
import tempfile
import zipfile
from pathlib import Path
import time
import traceback

# Import our pipeline modules
from validation_utils import ValidationError, handle_validation_error
from pipeline_validator import validate_before_pipeline
from pipeline_runner import run_pipeline_with_progress, PipelineExecutionResult
from pipeline_config import PipelineConfig, PipelineConstants

# Configure page
st.set_page_config(
    page_title="Ngoc Son Internal TSS Converter",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    .main-header {
        text-align: center;
        padding: 2rem 0;
        margin-bottom: 2rem;
    }
    
    .main-title {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 0.5rem;
    }
    
    .chart-icon {
        font-size: 2.2rem;
        background: linear-gradient(45deg, #3b82f6, #10b981, #f59e0b);
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .main-subtitle {
        font-size: 1.1rem;
        color: #6b7280;
        font-weight: 400;
        margin-top: 0.5rem;
    }
    
    .upload-section {
        background: #f8fafc;
        border: 2px dashed #d1d5db;
        border-radius: 12px;
        padding: 3rem 2rem;
        text-align: center;
        margin: 2rem 0;
        transition: all 0.3s ease;
    }
    
    .upload-section:hover {
        border-color: #3b82f6;
        background: #f0f9ff;
    }
    
    .upload-icon {
        font-size: 3rem;
        color: #9ca3af;
        margin-bottom: 1rem;
    }
    
    .upload-title {
        font-size: 1.25rem;
        font-weight: 600;
        color: #374151;
        margin-bottom: 0.5rem;
    }
    
    .upload-subtitle {
        color: #6b7280;
        margin-bottom: 2rem;
    }
    
    .file-constraints {
        font-size: 0.875rem;
        color: #9ca3af;
        margin-top: 1rem;
    }
    
    .process-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1rem 0;
    }
    
    .step-indicator {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        margin-bottom: 1rem;
    }
    
    .step-number {
        background: #3b82f6;
        color: white;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.75rem;
        font-weight: 600;
    }
    
    .step-completed {
        background: #10b981;
    }
    
    .step-current {
        background: #f59e0b;
    }
    
    .progress-bar {
        background: #e5e7eb;
        height: 8px;
        border-radius: 4px;
        overflow: hidden;
        margin: 1rem 0;
    }
    
    .progress-fill {
        background: linear-gradient(90deg, #3b82f6, #10b981);
        height: 100%;
        transition: width 0.3s ease;
    }
    
    .success-message {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 8px;
        padding: 1rem;
        color: #166534;
        margin: 1rem 0;
    }
    
    .error-message {
        background: #fef2f2;
        border: 1px solid #fecaca;
        border-radius: 8px;
        padding: 1rem;
        color: #dc2626;
        margin: 1rem 0;
    }
    
    .download-section {
        text-align: center;
        padding: 2rem;
        background: #f8fafc;
        border-radius: 12px;
        margin: 2rem 0;
    }
    
    /* Hide Streamlit default elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .stDeployButton {display: none;}
    
    /* Custom button styling */
    .stButton > button {
        background: #3b82f6;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.5rem 2rem;
        font-weight: 500;
        transition: all 0.2s ease;
    }
    
    .stButton > button:hover {
        background: #2563eb;
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

def show_header():
    """Display the main header section"""
    st.markdown("""
    <div class="main-header">
        <div class="main-title">
            <span class="chart-icon">📊</span>
            Ngoc Son Internal TSS Converter
        </div>
        <div class="main-subtitle">
            Convert Ngoc Son Internal TSS to Standard Internal TSS
        </div>
    </div>
    """, unsafe_allow_html=True)

def show_upload_section():
    """Display the file upload section"""
    st.markdown("""
    <div class="upload-section">
        <div class="upload-icon">🗂️</div>
        <div class="upload-title">Upload Excel File</div>
        <div class="upload-subtitle">Select .xlsx file to convert</div>
    </div>
    """, unsafe_allow_html=True)

def validate_uploaded_file(uploaded_file):
    """Validate the uploaded file before processing"""
    if uploaded_file is None:
        return False, "No file uploaded"
    
    # Check file extension
    if not uploaded_file.name.lower().endswith(('.xlsx', '.xls', '.xlsm')):
        return False, "Invalid file format. Please upload .xlsx, .xls, or .xlsm file"
    
    # Check file size (limit 200MB as shown in the image)
    if uploaded_file.size > 200 * 1024 * 1024:
        return False, "File too large. Maximum size is 200MB"
    
    return True, "File validation passed"

def process_pipeline(uploaded_file, progress_placeholder, status_placeholder):
    """Process the complete 8-step pipeline using centralized configuration"""
    
    try:
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)
            
            # Save uploaded file
            input_file = temp_dir / uploaded_file.name
            with open(input_file, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Create progress callback function for Streamlit
            def progress_callback(progress, current, total, status_message):
                update_progress(progress_placeholder, status_placeholder, current, total, status_message)
            
            # Run pipeline using centralized runner
            result = run_pipeline_with_progress(
                input_file=input_file,
                progress_callback=progress_callback,
                base_dir=str(temp_dir)
            )
            
            if result.success:
                return result.final_output, None
            else:
                return None, result.error
                
    except ValidationError as e:
        return None, f"Validation Error: {str(e)}"
    except Exception as e:
        return None, f"Processing Error: {str(e)}\n\nDetails:\n{traceback.format_exc()}"

def update_progress(progress_placeholder, status_placeholder, current, total, status):
    """Update progress bar and status"""
    progress = current / total
    
    with progress_placeholder.container():
        st.markdown(f"""
        <div class="progress-bar">
            <div class="progress-fill" style="width: {progress * 100}%"></div>
        </div>
        """, unsafe_allow_html=True)
        
        st.progress(progress)
    
    with status_placeholder.container():
        st.markdown(f"""
        <div class="step-indicator">
            <div class="step-number {'step-completed' if current >= total else 'step-current'}">{current}</div>
            <span>{status}</span>
        </div>
        """, unsafe_allow_html=True)

def main():
    """Main application"""
    show_header()
    
    # File upload section
    show_upload_section()
    
    uploaded_file = st.file_uploader(
        "",
        type=['xlsx', 'xls', 'xlsm'],
        help="Drag and drop file here\nLimit 200MB per file • XLSX",
        label_visibility="collapsed"
    )
    
    # Show file constraints
    st.markdown("""
    <div class="file-constraints">
        Drag and drop file here<br>
        Limit 200MB per file • XLSX
    </div>
    """, unsafe_allow_html=True)
    
    if uploaded_file is not None:
        # Validate file
        is_valid, message = validate_uploaded_file(uploaded_file)
        
        if not is_valid:
            st.markdown(f"""
            <div class="error-message">
                ❌ {message}
            </div>
            """, unsafe_allow_html=True)
            return
        
        # Show file info
        st.markdown(f"""
        <div class="process-card">
            <strong>📄 File uploaded:</strong> {uploaded_file.name}<br>
            <strong>📊 Size:</strong> {uploaded_file.size / (1024*1024):.1f} MB
        </div>
        """, unsafe_allow_html=True)
        
        # Process button
        if st.button("🚀 Start Conversion", type="primary"):
            
            # Progress tracking
            progress_placeholder = st.empty()
            status_placeholder = st.empty()
            
            # Process the pipeline
            with st.spinner("Processing your file..."):
                result_file, error = process_pipeline(uploaded_file, progress_placeholder, status_placeholder)
            
            if error:
                st.markdown(f"""
                <div class="error-message">
                    ❌ <strong>Processing Failed</strong><br>
                    {error}
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="success-message">
                    ✅ <strong>Conversion completed successfully!</strong><br>
                    Your file has been processed through all 8 steps.
                </div>
                """, unsafe_allow_html=True)
                
                # Download section
                if result_file and os.path.exists(result_file):
                    with open(result_file, "rb") as f:
                        result_data = f.read()
                    
                    # Generate download filename
                    base_name = uploaded_file.name.rsplit('.', 1)[0]
                    download_filename = f"{base_name}-Converted.xlsx"
                    
                    st.markdown("""
                    <div class="download-section">
                        <h3>📥 Download Your Converted File</h3>
                        <p>Standard Internal TSS format ready for use</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.download_button(
                        label="📥 Download Converted File",
                        data=result_data,
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
    
    # Footer
    st.markdown("""
    <div style="text-align: center; margin-top: 3rem; padding-top: 2rem; border-top: 1px solid #e5e7eb; color: #9ca3af;">
        <p>🔧 Powered by SEDO TSS Converter Pipeline • Built with ❤️ using Streamlit</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()