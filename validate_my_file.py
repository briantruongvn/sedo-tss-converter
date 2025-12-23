#!/usr/bin/env python3
"""
SEDO TSS Converter - User File Validator
Simple validation script for users to check their files before submission
"""

import sys
import os
from pathlib import Path

# Add current directory to path to import validation modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from pipeline_validator import validate_before_pipeline
except ImportError:
    print("❌ Error: Cannot import validation modules")
    print("   Please ensure you're running this script from the SEDO TSS Converter directory")
    sys.exit(1)

def main():
    """Simple user-friendly file validator"""
    
    print("🔍 SEDO TSS Converter - File Validator")
    print("=" * 50)
    
    # Get file path from user
    if len(sys.argv) < 2:
        print("📁 Please provide your Excel file path:")
        print("   Usage: python validate_my_file.py \"path/to/your/file.xlsx\"")
        print("   Example: python validate_my_file.py \"data/input/my_file.xlsx\"")
        sys.exit(1)
    
    file_path = sys.argv[1]
    
    print(f"📄 Validating file: {file_path}")
    print("-" * 50)
    
    try:
        # Run comprehensive validation
        success = validate_before_pipeline(file_path, verbose=True)
        
        if success:
            print("\n" + "=" * 50)
            print("🎉 VALIDATION SUCCESSFUL! 🎉")
            print("=" * 50)
            print("✅ Your file is ready for the SEDO TSS Converter pipeline")
            print("🚀 Next step: Run the pipeline with:")
            print(f'   python step1_unmerge_standalone.py "{file_path}"')
            print("=" * 50)
            return True
        else:
            print("\n" + "=" * 50)
            print("❌ VALIDATION FAILED")
            print("=" * 50)
            print("🔧 Please fix the issues above and run validation again")
            print("📋 For detailed help, check: INPUT_REQUIREMENTS.md")
            print("=" * 50)
            return False
            
    except KeyboardInterrupt:
        print("\n❌ Validation cancelled by user")
        return False
    except Exception as e:
        print(f"\n🚨 Unexpected error during validation: {e}")
        print("📞 Please contact support with this error message")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)