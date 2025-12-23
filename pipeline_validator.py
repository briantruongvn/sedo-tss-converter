#!/usr/bin/env python3
"""
Pipeline Pre-flight Validator for SEDO TSS Converter
Comprehensive validation before starting any processing steps
"""

import openpyxl
import logging
from pathlib import Path
from typing import Union, Optional, List, Dict, Tuple
import sys

from validation_utils import ValidationError, FileValidator, HeaderDetector, ErrorHandler

logger = logging.getLogger(__name__)

class PipelineValidator:
    """
    Comprehensive pipeline validation before execution starts
    
    Validates entire pipeline requirements upfront to prevent mid-pipeline failures
    """
    
    def __init__(self, base_dir: Optional[str] = None):
        self.base_dir = Path(base_dir) if base_dir else Path.cwd()
        self.validation_results = []
        
    def validate_complete_pipeline(self, input_file: Union[str, Path]) -> Dict[str, any]:
        """
        Comprehensive pre-flight validation for entire 8-step pipeline
        
        Args:
            input_file: Input Excel file path
            
        Returns:
            Validation report dict
            
        Raises:
            ValidationError: If critical validation fails
        """
        logger.info("🔍 Starting comprehensive pipeline validation...")
        
        input_path = Path(input_file)
        
        # Stage 1: Basic file validation
        self._validate_input_file(input_path)
        
        # Stage 2: Excel structure validation  
        excel_stats = self._validate_excel_structure(input_path)
        
        # Stage 3: Step-specific requirements validation (reuse loaded workbook)
        self._validate_step_requirements(input_path, excel_stats)
        
        # Stage 4: System resources validation
        self._validate_system_resources()
        
        # Generate validation report
        report = {
            'status': 'PASSED',
            'input_file': str(input_path),
            'excel_stats': excel_stats,
            'validation_results': self.validation_results,
            'timestamp': __import__('datetime').datetime.now().isoformat()
        }
        
        logger.info("✅ Pipeline validation completed successfully")
        return report
    
    def _validate_input_file(self, input_path: Path):
        """Stage 1: Basic file validation"""
        logger.debug("Stage 1: Basic file validation")
        
        try:
            FileValidator.validate_input_file(input_path)
            self.validation_results.append({
                'stage': 'file_validation',
                'status': 'PASSED',
                'message': f'File validation passed: {input_path.name}'
            })
        except Exception as e:
            raise ValidationError(
                f"Critical file validation failed: {str(e)}",
                error_code="FILE_VALIDATION_FAILED",
                severity=ValidationError.CRITICAL,
                category=ValidationError.FILE_ERROR,
                suggestions=[
                    "Check if the file path is correct",
                    "Ensure the file is a valid Excel format (.xlsx, .xls, .xlsm)",
                    "Verify file is not corrupted or locked by another application"
                ],
                step="Pre-flight"
            )
    
    def _validate_excel_structure(self, input_path: Path) -> Dict[str, any]:
        """Stage 2: Excel structure validation"""
        logger.debug("Stage 2: Excel structure validation")
        
        try:
            stats = FileValidator.validate_excel_structure(input_path)
            
            # Additional structure checks
            if stats['max_row'] < 15:
                raise ValidationError(
                    f"File too small: {stats['max_row']} rows (minimum 15 required)",
                    error_code="FILE_TOO_SMALL",
                    severity=ValidationError.CRITICAL,
                    category=ValidationError.STRUCTURE_ERROR,
                    suggestions=[
                        "Ensure this is the correct input file",
                        "Check if file contains the expected data structure",
                        "SEDO files typically have headers around row 15 + data rows"
                    ],
                    step="Pre-flight"
                )
            
            if stats['max_col'] < 10:
                raise ValidationError(
                    f"File too narrow: {stats['max_col']} columns (minimum 10 required)",
                    error_code="FILE_TOO_NARROW", 
                    severity=ValidationError.CRITICAL,
                    category=ValidationError.STRUCTURE_ERROR,
                    suggestions=[
                        "Ensure this is the complete input file",
                        "Check if columns are hidden or data is missing",
                        "Verify file structure matches expected format"
                    ],
                    step="Pre-flight"
                )
            
            self.validation_results.append({
                'stage': 'excel_structure',
                'status': 'PASSED',
                'message': f'Excel structure valid: {stats["max_row"]} rows, {stats["max_col"]} cols',
                'details': stats
            })
            
            return stats
            
        except ValidationError:
            raise  # Re-raise our custom validation errors
        except Exception as e:
            raise ValidationError(
                f"Excel structure validation failed: {str(e)}",
                error_code="EXCEL_STRUCTURE_FAILED",
                severity=ValidationError.CRITICAL,
                category=ValidationError.STRUCTURE_ERROR,
                suggestions=[
                    "Check if file is corrupted",
                    "Try opening the file in Excel to verify it works",
                    "Ensure file is not password protected"
                ],
                step="Pre-flight"
            )
    
    def _validate_step_requirements(self, input_path: Path, excel_stats: Dict[str, any] = None):
        """Stage 3: Step-specific requirements validation"""
        logger.debug("Stage 3: Step-specific requirements validation")
        
        try:
            # Load workbook for header checks (not read-only to access merged_cells)
            # Note: This is separate from the read-only validation in Stage 2
            wb = openpyxl.load_workbook(str(input_path), read_only=False)
            ws = wb.active
            
            # Step 1 requirements: Check for merged cells (informational)
            merged_ranges = len(list(ws.merged_cells.ranges))
            if merged_ranges == 0:
                self.validation_results.append({
                    'stage': 'step1_check',
                    'status': 'INFO',
                    'message': 'No merged cells found - Step 1 will complete quickly'
                })
            else:
                self.validation_results.append({
                    'stage': 'step1_check', 
                    'status': 'PASSED',
                    'message': f'Found {merged_ranges} merged ranges for processing'
                })
            
            # Step 2 requirements: Check for General Type header
            header_result = HeaderDetector.find_general_type_header(ws)
            if header_result is None:
                raise ValidationError(
                    "Required 'General Type/Sub-Type in Connect' header not found",
                    error_code="GENERAL_TYPE_HEADER_MISSING",
                    severity=ValidationError.CRITICAL,
                    category=ValidationError.HEADER_ERROR,
                    suggestions=[
                        "Check if your file contains 'General Type/Sub-Type in Connect' header",
                        "Look for similar headers like 'General Type of Material in Connect'", 
                        "Ensure headers are in the first 50 rows",
                        "Verify file is not missing header rows"
                    ],
                    step="Step 2"
                )
            else:
                row, col, matched_text = header_result
                self.validation_results.append({
                    'stage': 'step2_check',
                    'status': 'PASSED',
                    'message': f'General Type header found: "{matched_text}" at row {row}'
                })
            
            # Step 4 requirements: Check for Article headers  
            article_headers = HeaderDetector.find_article_headers(ws)
            if article_headers is None:
                self.validation_results.append({
                    'stage': 'step4_check',
                    'status': 'WARNING',
                    'message': 'Article Name/Number headers not found - will create template without articles'
                })
            else:
                name_col, no_col, header_row = article_headers
                self.validation_results.append({
                    'stage': 'step4_check',
                    'status': 'PASSED',
                    'message': f'Article headers found at row {header_row}: Name col {name_col}, Number col {no_col}'
                })
            
            wb.close()
            
        except ValidationError:
            raise  # Re-raise our validation errors
        except Exception as e:
            raise ValidationError(
                f"Step requirements validation failed: {str(e)}",
                error_code="STEP_REQUIREMENTS_FAILED",
                severity=ValidationError.CRITICAL,
                category=ValidationError.HEADER_ERROR,
                suggestions=[
                    "Check if file can be opened in Excel",
                    "Verify file is not corrupted",
                    "Ensure file contains expected header structure"
                ],
                step="Pre-flight"
            )
    
    def _validate_system_resources(self):
        """Stage 4: System resources validation"""
        logger.debug("Stage 4: System resources validation")
        
        # Check output directory
        output_dir = self.base_dir / "data" / "output"
        if not output_dir.exists():
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
                self.validation_results.append({
                    'stage': 'system_resources',
                    'status': 'INFO',
                    'message': f'Created output directory: {output_dir}'
                })
            except Exception as e:
                raise ValidationError(
                    f"Cannot create output directory: {output_dir}",
                    error_code="OUTPUT_DIR_CREATION_FAILED",
                    severity=ValidationError.CRITICAL,
                    category=ValidationError.FILE_ERROR,
                    suggestions=[
                        "Check if you have write permissions in the project directory",
                        "Ensure disk has sufficient space",
                        "Try running as administrator if necessary"
                    ],
                    step="Pre-flight"
                )
        
        # Check disk space (warn if < 100MB free)
        try:
            import shutil
            free_space = shutil.disk_usage(output_dir).free / (1024 * 1024)  # MB
            if free_space < 100:
                self.validation_results.append({
                    'stage': 'system_resources',
                    'status': 'WARNING', 
                    'message': f'Low disk space: {free_space:.1f}MB free (recommended: >100MB)'
                })
            else:
                self.validation_results.append({
                    'stage': 'system_resources',
                    'status': 'PASSED',
                    'message': f'Sufficient disk space: {free_space:.1f}MB free'
                })
        except Exception:
            # Disk space check failed, but not critical
            self.validation_results.append({
                'stage': 'system_resources',
                'status': 'INFO',
                'message': 'Could not check disk space - proceeding anyway'
            })
    
    def print_validation_report(self, report: Dict[str, any]):
        """Print formatted validation report"""
        print("\n" + "="*60)
        print("📋 PIPELINE VALIDATION REPORT")
        print("="*60)
        print(f"📁 Input File: {report['input_file']}")
        print(f"📊 File Stats: {report['excel_stats']['max_row']} rows, {report['excel_stats']['max_col']} cols")
        print(f"⏰ Validation Time: {report['timestamp']}")
        print("\n🔍 Validation Results:")
        
        for result in report['validation_results']:
            status_icon = {
                'PASSED': '✅',
                'WARNING': '⚠️',
                'INFO': 'ℹ️',
                'FAILED': '❌'
            }.get(result['status'], '•')
            
            print(f"  {status_icon} {result['stage']}: {result['message']}")
        
        print(f"\n🎯 Overall Status: {status_icon} {report['status']}")
        print("="*60)

def validate_before_pipeline(input_file: Union[str, Path], verbose: bool = False) -> bool:
    """
    Convenience function for pipeline validation
    
    Args:
        input_file: Input file to validate
        verbose: Show detailed validation report
        
    Returns:
        True if validation passes, False otherwise (also raises ValidationError)
    """
    try:
        validator = PipelineValidator()
        report = validator.validate_complete_pipeline(input_file)
        
        if verbose:
            validator.print_validation_report(report)
        
        return True
        
    except ValidationError as e:
        print(e.get_formatted_error())
        if verbose:
            print(f"\n💥 Validation failed - pipeline cannot continue")
            print("🔧 Please fix the issues above and try again")
        return False
    except Exception as e:
        print(f"🚨 Unexpected validation error: {str(e)}")
        return False

if __name__ == "__main__":
    """CLI interface for standalone validation"""
    import argparse
    
    parser = argparse.ArgumentParser(description='SEDO TSS Pipeline Validator')
    parser.add_argument('input_file', help='Input Excel file to validate')
    parser.add_argument('-v', '--verbose', action='store_true', help='Show detailed validation report')
    
    args = parser.parse_args()
    
    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(asctime)s - %(levelname)s - %(message)s')
    
    success = validate_before_pipeline(args.input_file, args.verbose)
    sys.exit(0 if success else 1)