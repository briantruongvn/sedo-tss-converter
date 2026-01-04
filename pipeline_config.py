#!/usr/bin/env python3
"""
Pipeline Configuration - Single Point of Truth
SEDO TSS Converter Pipeline Configuration

This module defines the complete 8-step pipeline configuration with metadata,
step descriptions, and processing classes. This serves as the single source 
of truth for both CLI and Streamlit interfaces.

DESIGN PRINCIPLES:
- Single Point of Truth: All pipeline metadata centralized here
- Adaptive Configuration: Easy to add/modify/remove steps  
- Consistent Interface: Standardized step class requirements
- Flexible Execution: Support both CLI and web interfaces
"""

from dataclasses import dataclass
from typing import Type, Optional, List, Dict, Any
from pathlib import Path

@dataclass
class StepMetadata:
    """
    Metadata for a single pipeline step
    
    Attributes:
        step_number: Step number in pipeline (1-8)
        name: Internal step name (e.g., "unmerge_cells")
        display_name: User-friendly display name
        description: Detailed description of what step does
        class_name: Python class that implements this step
        module_name: Python module containing the class
        requires_original_input: Whether step needs original input file
        depends_on: List of step numbers this step depends on
        cli_script: CLI script filename for standalone execution
        estimated_duration_seconds: Estimated processing time
    """
    step_number: int
    name: str
    display_name: str 
    description: str
    class_name: str
    module_name: str
    requires_original_input: bool = False
    depends_on: List[int] = None
    cli_script: Optional[str] = None
    estimated_duration_seconds: int = 5

    def __post_init__(self):
        if self.depends_on is None:
            self.depends_on = [self.step_number - 1] if self.step_number > 1 else []

class PipelineConfig:
    """
    Complete pipeline configuration and metadata management
    
    This class serves as the single point of truth for all pipeline steps,
    their metadata, dependencies, and execution requirements.
    """
    
    # Define all 8 pipeline steps with complete metadata
    STEPS = [
        StepMetadata(
            step_number=1,
            name="unmerge_cells",
            display_name="Unmerging cells",
            description="Unmerge all merged cell ranges and preserve data by filling empty cells",
            class_name="ExcelUnmerger",
            module_name="step1_unmerge_standalone",
            requires_original_input=True,
            depends_on=[],  # First step has no dependencies
            cli_script="step1_unmerge_standalone.py",
            estimated_duration_seconds=10
        ),
        StepMetadata(
            step_number=2, 
            name="process_headers",
            display_name="Processing headers",
            description="Process headers using 3-case logic for General Type/Sub-Type structure",
            class_name="HeaderProcessor",
            module_name="step2_header_processing",
            depends_on=[1],
            cli_script="step2_header_processing.py",
            estimated_duration_seconds=5
        ),
        StepMetadata(
            step_number=3,
            name="create_template",
            display_name="Creating template", 
            description="Create structured template with 17 standardized headers and formatting",
            class_name="TemplateCreator",
            module_name="step3_template_creation",
            depends_on=[2],
            cli_script="step3_template_creation.py",
            estimated_duration_seconds=3
        ),
        StepMetadata(
            step_number=4,
            name="fill_articles",
            display_name="Filling article information",
            description="Extract and fill article information with dynamic header detection",
            class_name="ArticleFiller", 
            module_name="step4_article_filling",
            requires_original_input=True,
            depends_on=[3],
            cli_script="step4_article_filling.py",
            estimated_duration_seconds=8
        ),
        StepMetadata(
            step_number=5,
            name="transform_data", 
            display_name="Transforming data",
            description="Transform and map data from Step 2 to Step 4 template structure",
            class_name="DataTransformer",
            module_name="step5_data_transformation", 
            depends_on=[2, 4],
            cli_script="step5_data_transformation.py",
            estimated_duration_seconds=7
        ),
        StepMetadata(
            step_number=6,
            name="process_sd_data",
            display_name="Processing SD data",
            description="Process SD data with H→P mapping, multi-line parsing and de-duplication",
            class_name="SDProcessor",
            module_name="step6_sd_processing",
            depends_on=[2, 5],
            cli_script="step6_sd_processing.py", 
            estimated_duration_seconds=12
        ),
        StepMetadata(
            step_number=7,
            name="validate_products", 
            display_name="Validating finished products",
            description="Match finished products with articles using fuzzy matching and 'All items' logic",
            class_name="FinishedProductProcessor",
            module_name="step7_finished_product",
            requires_original_input=True,
            depends_on=[6],
            cli_script="step7_finished_product.py",
            estimated_duration_seconds=6
        ),
        StepMetadata(
            step_number=8,
            name="process_document",
            display_name="Processing final document", 
            description="Final document processing with requirement source extraction and validation",
            class_name="DocumentProcessor",
            module_name="step8_document_processing",
            requires_original_input=True,
            depends_on=[7],
            cli_script="step8_document_processing.py",
            estimated_duration_seconds=8
        )
    ]
    
    @classmethod
    def get_all_steps(cls) -> List[StepMetadata]:
        """Get all pipeline steps in order"""
        return cls.STEPS
    
    @classmethod  
    def get_step(cls, step_number: int) -> Optional[StepMetadata]:
        """Get specific step by number"""
        for step in cls.STEPS:
            if step.step_number == step_number:
                return step
        return None
    
    @classmethod
    def get_step_by_name(cls, name: str) -> Optional[StepMetadata]:
        """Get specific step by name"""
        for step in cls.STEPS:
            if step.name == name:
                return step
        return None
    
    @classmethod
    def get_step_count(cls) -> int:
        """Get total number of steps"""
        return len(cls.STEPS)
    
    @classmethod
    def get_step_names(cls) -> List[str]:
        """Get list of all step display names"""
        return [step.display_name for step in cls.STEPS]
    
    @classmethod
    def get_step_descriptions(cls) -> Dict[int, str]:
        """Get mapping of step number to description"""
        return {step.step_number: step.description for step in cls.STEPS}
    
    @classmethod
    def get_dependencies(cls, step_number: int) -> List[int]:
        """Get list of step numbers that given step depends on"""
        step = cls.get_step(step_number)
        return step.depends_on if step else []
    
    @classmethod
    def validate_step_order(cls, steps_to_run: List[int]) -> bool:
        """Validate that steps can be run in given order based on dependencies"""
        completed_steps = set()
        
        for step_num in steps_to_run:
            step = cls.get_step(step_num)
            if not step:
                return False
                
            # Check if all dependencies are completed
            for dep in step.depends_on:
                if dep not in completed_steps and dep != 0:  # 0 = no dependency
                    return False
                    
            completed_steps.add(step_num)
            
        return True
    
    @classmethod
    def get_estimated_duration(cls, steps: List[int] = None) -> int:
        """Get estimated total duration for given steps (or all steps)"""
        if steps is None:
            steps = [step.step_number for step in cls.STEPS]
            
        total_duration = 0
        for step_num in steps:
            step = cls.get_step(step_num)
            if step:
                total_duration += step.estimated_duration_seconds
                
        return total_duration
    
    @classmethod
    def get_progress_messages(cls) -> List[str]:
        """Get list of progress messages for UI display"""
        messages = ["Validating input file"]  # Pre-validation step
        messages.extend([step.display_name for step in cls.STEPS])
        return messages
    
    @classmethod
    def requires_original_input(cls, step_number: int) -> bool:
        """Check if step requires original input file"""
        step = cls.get_step(step_number)
        return step.requires_original_input if step else False
    
    @classmethod
    def get_output_filename(cls, input_filename: str, step_number: int) -> str:
        """Generate standardized output filename for given step"""
        base_name = Path(input_filename).stem
        return f"{base_name}-Step{step_number}.xlsx"
    
    @classmethod
    def get_cli_command(cls, step_number: int, input_file: str, **kwargs) -> Optional[str]:
        """Generate CLI command for given step"""
        step = cls.get_step(step_number)
        if not step or not step.cli_script:
            return None
            
        cmd = f"python {step.cli_script} {input_file}"
        
        # Add step-specific arguments
        if step_number == 4:  # Article filling needs template
            template_file = kwargs.get('template_file')
            if template_file:
                cmd += f" {template_file}"
        elif step_number in [5, 6, 7, 8]:  # Steps that need previous step file
            prev_step_file = kwargs.get('prev_step_file')
            if prev_step_file:
                cmd += f" --step{step_number-1}-file {prev_step_file}"
                
        return cmd

# Validation constants for pipeline
class PipelineConstants:
    """Constants used throughout the pipeline"""
    
    # File size limits
    MAX_FILE_SIZE_MB = 200
    MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024
    
    # Supported file formats
    SUPPORTED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm']
    
    # Directory structure
    INPUT_DIR = "data/input"
    OUTPUT_DIR = "data/output"
    
    # Key headers to detect in input files
    REQUIRED_HEADERS = [
        "General Type/Sub-Type in Connect",
        "Article Name",
        "Article No."
    ]
    
    # Progress tracking
    VALIDATION_STEP_INDEX = 0  # Pre-validation is step 0
    TOTAL_STEPS_WITH_VALIDATION = len(PipelineConfig.STEPS) + 1
    
    # Error handling
    MAX_RETRIES = 3
    RETRY_DELAY_SECONDS = 1

# Export main classes for easy importing
__all__ = [
    'StepMetadata', 
    'PipelineConfig', 
    'PipelineConstants'
]