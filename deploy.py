#!/usr/bin/env python3
"""
Deployment helper script for Streamlit app
Run this to test the app locally before deploying
"""

import subprocess
import sys
import os
from pathlib import Path

def check_dependencies():
    """Check if all required dependencies are installed"""
    print("🔍 Checking dependencies...")
    
    try:
        import streamlit
        print(f"✅ Streamlit {streamlit.__version__}")
    except ImportError:
        print("❌ Streamlit not installed")
        print("💡 Install with: pip install streamlit>=1.28.0")
        return False
    
    try:
        import openpyxl
        print(f"✅ openpyxl {openpyxl.__version__}")
    except ImportError:
        print("❌ openpyxl not installed") 
        print("💡 Install with: pip install openpyxl>=3.1.0")
        return False
    
    try:
        import pandas as pd
        print(f"✅ pandas {pd.__version__}")
    except ImportError:
        print("❌ pandas not installed")
        print("💡 Install with: pip install pandas>=2.1.0")
        return False
    
    return True

def test_imports():
    """Test all application imports"""
    print("\n🧪 Testing application imports...")
    
    try:
        from validation_utils import ValidationError
        from pipeline_validator import validate_before_pipeline
        print("✅ Validation modules")
    except ImportError as e:
        print(f"❌ Validation import error: {e}")
        return False
    
    try:
        from step1_unmerge_standalone import ExcelUnmerger
        from step2_header_processing import HeaderProcessor
        from step3_template_creation import TemplateCreator
        from step4_article_filling import ArticleFiller
        from step5_data_transformation import DataTransformer
        from step6_sd_processing import SDProcessor
        from step7_finished_product import FinishedProductProcessor
        from step8_document_processing import DocumentProcessor
        print("✅ Pipeline modules")
    except ImportError as e:
        print(f"❌ Pipeline import error: {e}")
        return False
    
    return True

def run_local_server():
    """Run the Streamlit app locally"""
    print("\n🚀 Starting Streamlit app...")
    print("📱 App will be available at: http://localhost:8501")
    print("🛑 Press Ctrl+C to stop the server")
    
    try:
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", "app.py",
            "--server.headless", "false",
            "--browser.gatherUsageStats", "false"
        ], check=True)
    except KeyboardInterrupt:
        print("\n👋 Server stopped by user")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to start server: {e}")
        return False
    
    return True

def show_deployment_info():
    """Show deployment information"""
    print("\n" + "="*50)
    print("🌐 STREAMLIT CLOUD DEPLOYMENT")
    print("="*50)
    print("1. 📚 Push code to GitHub:")
    print("   git add .")
    print("   git commit -m 'Add Streamlit web app'")
    print("   git push origin main")
    print()
    print("2. 🚀 Deploy to Streamlit Cloud:")
    print("   • Go to https://share.streamlit.io")
    print("   • Connect your GitHub account")
    print("   • Select your repository")
    print("   • Set main file: app.py") 
    print("   • Click Deploy")
    print()
    print("3. 📋 Required files for deployment:")
    print("   ✅ app.py (main application)")
    print("   ✅ requirements.txt (dependencies)")
    print("   ✅ .streamlit/config.toml (configuration)")
    print("   ✅ All step*.py files (pipeline)")
    print("   ✅ validation_utils.py & pipeline_validator.py")
    print()
    print("4. 🔧 Configuration:")
    print("   • Max file size: 200MB")
    print("   • Supported formats: .xlsx, .xls, .xlsm")
    print("   • Processing: 8-step pipeline")
    print("   • UI: Modern, responsive design")
    print()
    print("5. 📱 Features:")
    print("   • Drag & drop file upload")
    print("   • Real-time progress tracking") 
    print("   • Instant file download")
    print("   • Mobile-friendly interface")
    print("="*50)

def main():
    """Main deployment function"""
    print("🌟 SEDO TSS Converter - Streamlit Deployment Helper")
    print("="*50)
    
    # Check dependencies
    if not check_dependencies():
        print("\n❌ Dependency check failed. Please install missing packages.")
        return False
    
    # Test imports
    if not test_imports():
        print("\n❌ Import test failed. Check your Python path and modules.")
        return False
    
    print("\n✅ All checks passed! App is ready to deploy.")
    
    # Ask what to do
    print("\n🤔 What would you like to do?")
    print("1. 🖥️  Run locally for testing")
    print("2. 📋 Show deployment instructions")
    print("3. 🚪 Exit")
    
    while True:
        choice = input("\nEnter choice (1-3): ").strip()
        
        if choice == "1":
            run_local_server()
            break
        elif choice == "2":
            show_deployment_info()
            break
        elif choice == "3":
            print("👋 Goodbye!")
            break
        else:
            print("❌ Invalid choice. Please enter 1, 2, or 3.")

if __name__ == "__main__":
    main()