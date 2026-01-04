#!/usr/bin/env python3
"""
Pipeline Runner - Unified Step Execution and Orchestration
SEDO TSS Converter Pipeline Runner

This module provides unified pipeline execution for both CLI and Streamlit interfaces.
It uses the centralized pipeline configuration to orchestrate step execution,
handle dependencies, and provide consistent progress tracking.

DESIGN PRINCIPLES:
- Single Source of Truth: Uses PipelineConfig for all step metadata
- Flexible Execution: Supports both CLI and web interface requirements
- Error Handling: Comprehensive error handling and validation
- Progress Tracking: Consistent progress reporting across interfaces
- Resource Management: Proper file and memory management
"""

import importlib
import logging
import tempfile
import time
from pathlib import Path
from typing import Union, Optional, Callable, Dict, Any, List, Tuple
import traceback

from pipeline_config import PipelineConfig, StepMetadata, PipelineConstants
from validation_utils import ValidationError, handle_validation_error
from pipeline_validator import validate_before_pipeline

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PipelineExecutionResult:
    """Results from pipeline execution"""
    
    def __init__(self, success: bool, final_output: Optional[str] = None, 
                 error: Optional[str] = None, step_outputs: Optional[Dict[int, str]] = None):
        self.success = success
        self.final_output = final_output
        self.error = error
        self.step_outputs = step_outputs or {}
        self.execution_time = 0.0
        self.completed_steps = len(self.step_outputs)

class ProgressTracker:
    """Handles progress tracking for both CLI and Streamlit"""
    
    def __init__(self, progress_callback: Optional[Callable] = None, 
                 status_callback: Optional[Callable] = None):
        self.progress_callback = progress_callback
        self.status_callback = status_callback
        self.total_steps = PipelineConstants.TOTAL_STEPS_WITH_VALIDATION
        
    def update(self, current_step: int, status_message: str, verbose: bool = False):
        """Update progress with current step and status"""
        progress = current_step / self.total_steps
        
        # Call callbacks if provided (for Streamlit)
        if self.progress_callback:
            self.progress_callback(progress, current_step, self.total_steps, status_message)
        if self.status_callback:
            self.status_callback(current_step, status_message)
            
        # Log for CLI
        if verbose:
            logger.info(f"Step {current_step}/{self.total_steps}: {status_message}")

class PipelineRunner:
    """
    Unified pipeline execution engine
    
    Orchestrates the complete 8-step pipeline using centralized configuration.
    Supports both CLI and Streamlit execution with consistent behavior.
    """
    
    def __init__(self, base_dir: Optional[Union[str, Path]] = None, verbose: bool = False):
        self.base_dir = Path(base_dir) if base_dir else Path.cwd()
        self.output_dir = self.base_dir / PipelineConstants.OUTPUT_DIR
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.verbose = verbose
        self.step_instances = {}  # Cache for step instances
        
    def _get_step_class(self, step: StepMetadata):
        """Dynamically import and get step class"""
        try:
            module = importlib.import_module(step.module_name)
            step_class = getattr(module, step.class_name)
            return step_class
        except (ImportError, AttributeError) as e:
            raise ValidationError(f"Failed to import {step.class_name} from {step.module_name}: {e}")
    
    def _get_step_instance(self, step: StepMetadata):
        """Get or create step instance (with caching)"""
        if step.step_number not in self.step_instances:
            step_class = self._get_step_class(step)
            # Most step classes take base_dir as constructor parameter
            self.step_instances[step.step_number] = step_class(str(self.base_dir))
        return self.step_instances[step.step_number]
    
    def _validate_input_file(self, input_file: Union[str, Path]) -> bool:
        """Validate input file before starting pipeline"""
        try:
            return validate_before_pipeline(input_file, verbose=self.verbose)
        except Exception as e:
            if self.verbose:
                logger.error(f"Input validation failed: {e}")
            return False
    
    def _get_step_inputs(self, step: StepMetadata, input_file: Path, 
                        step_outputs: Dict[int, str]) -> List[Union[str, Path]]:
        """Get input arguments for a specific step"""
        inputs = []
        
        # Special handling for each step's input requirements
        if step.step_number == 1:
            # Step 1: Only needs original input
            inputs.append(input_file)
        elif step.step_number == 5:  
            # Step 5: DataTransformer needs step2 + step4
            step2_output = step_outputs.get(2)
            step4_output = step_outputs.get(4) 
            if step2_output and step4_output:
                inputs = [step2_output, step4_output]
            else:
                raise ValidationError(f"Step 5 requires outputs from Step 2 and Step 4")
        elif step.step_number == 6:  
            # Step 6: SDProcessor needs step2 as main input
            step2_output = step_outputs.get(2)
            if step2_output:
                inputs = [step2_output]  # step2 as main input
            else:
                raise ValidationError(f"Step 6 requires output from Step 2")
        elif step.step_number == 7:
            # Step 7: Only needs step6 output (no original input)
            step6_output = step_outputs.get(6)
            if step6_output:
                inputs = [step6_output]
            else:
                raise ValidationError(f"Step 7 requires output from Step 6")
        elif step.step_number == 8:
            # Step 8: Takes step7 output as input (despite parameter being named step6_file)
            step7_output = step_outputs.get(7)
            if step7_output:
                inputs = [step7_output]
            else:
                raise ValidationError(f"Step 8 requires output from Step 7")
        elif step.requires_original_input and step.step_number > 1:
            # Steps that need both original input and previous step output
            inputs = [input_file]
            prev_step_num = step.step_number - 1
            if prev_step_num in step_outputs:
                inputs.append(step_outputs[prev_step_num])
            else:
                raise ValidationError(f"Step {step.step_number} requires output from Step {prev_step_num}")
        else:
            # Default: Get output from previous step
            prev_step_num = step.step_number - 1
            if prev_step_num in step_outputs:
                inputs.append(step_outputs[prev_step_num])
            else:
                raise ValidationError(f"Step {step.step_number} requires output from Step {prev_step_num}")
        
        return inputs
    
    def _get_step_kwargs(self, step: StepMetadata, step_outputs: Dict[int, str]) -> Dict[str, Any]:
        """Get keyword arguments for a specific step"""
        kwargs = {}
        
        # Steps that need additional file parameters
        if step.step_number == 6 and 5 in step_outputs:  # SD processing
            kwargs['step4_file'] = step_outputs[5]  # step5 output goes to step4_file param 
            
        return kwargs
    
    def _execute_step(self, step: StepMetadata, input_file: Path, 
                     step_outputs: Dict[int, str], tracker: ProgressTracker) -> str:
        """Execute a single pipeline step"""
        
        # Update progress
        tracker.update(step.step_number, step.display_name, self.verbose)
        
        # Get step instance
        step_instance = self._get_step_instance(step)
        
        # Prepare inputs
        inputs = self._get_step_inputs(step, input_file, step_outputs)
        kwargs = self._get_step_kwargs(step, step_outputs)
        
        # Execute step based on its interface
        try:
            if step.step_number == 1:  # ExcelUnmerger
                output_file = step_instance.unmerge_file(*inputs)
            elif step.step_number == 2:  # HeaderProcessor 
                output_file = step_instance.process_file(*inputs)
            elif step.step_number == 3:  # TemplateCreator
                output_file = step_instance.create_template(*inputs)
            elif step.step_number == 4:  # ArticleFiller
                output_file = step_instance.fill_article_info(*inputs)
            elif step.step_number == 5:  # DataTransformer
                output_file = step_instance.transform_data(*inputs)
            elif step.step_number == 6:  # SDProcessor
                output_file = step_instance.process_sd_data(*inputs, **kwargs)
            elif step.step_number == 7:  # FinishedProductProcessor
                output_file = step_instance.process_finished_products(*inputs)  # No kwargs for step 7
            elif step.step_number == 8:  # DocumentProcessor
                output_file = step_instance.process_step7(*inputs)  # No kwargs for step 8
            else:
                raise ValidationError(f"Unknown step number: {step.step_number}")
                
            return str(output_file)
            
        except Exception as e:
            error_msg = f"Step {step.step_number} ({step.display_name}) failed: {str(e)}"
            if self.verbose:
                error_msg += f"\n\nDetails:\n{traceback.format_exc()}"
            raise ValidationError(error_msg)
    
    def run_pipeline(self, input_file: Union[str, Path], 
                    progress_callback: Optional[Callable] = None,
                    status_callback: Optional[Callable] = None,
                    steps_to_run: Optional[List[int]] = None) -> PipelineExecutionResult:
        """
        Run the complete pipeline or specific steps
        
        Args:
            input_file: Path to input Excel file
            progress_callback: Callback for progress updates (progress, current, total, status)
            status_callback: Callback for status updates (current, status)  
            steps_to_run: List of step numbers to run (default: all steps 1-8)
            
        Returns:
            PipelineExecutionResult with success status, outputs, and error info
        """
        
        input_file = Path(input_file)
        start_time = time.time()
        
        if steps_to_run is None:
            steps_to_run = [step.step_number for step in PipelineConfig.get_all_steps()]
        
        # Validate step order
        if not PipelineConfig.validate_step_order(steps_to_run):
            return PipelineExecutionResult(
                success=False, 
                error="Invalid step order - dependencies not satisfied"
            )
        
        tracker = ProgressTracker(progress_callback, status_callback)
        step_outputs = {}
        
        try:
            # Step 0: Pre-validation
            tracker.update(0, "Validating input file", self.verbose)
            
            if not self._validate_input_file(input_file):
                return PipelineExecutionResult(
                    success=False,
                    error="Input file validation failed"
                )
            
            # Execute each requested step
            for step_num in steps_to_run:
                step = PipelineConfig.get_step(step_num)
                if not step:
                    return PipelineExecutionResult(
                        success=False,
                        error=f"Unknown step number: {step_num}",
                        step_outputs=step_outputs
                    )
                
                try:
                    output_file = self._execute_step(step, input_file, step_outputs, tracker)
                    step_outputs[step_num] = output_file
                    
                    if self.verbose:
                        logger.info(f"Step {step_num} completed: {output_file}")
                        
                except ValidationError as e:
                    return PipelineExecutionResult(
                        success=False,
                        error=str(e),
                        step_outputs=step_outputs
                    )
            
            # Final progress update
            tracker.update(
                PipelineConstants.TOTAL_STEPS_WITH_VALIDATION, 
                "Processing complete!", 
                self.verbose
            )
            
            # Get final output (highest step number)
            final_step = max(steps_to_run)
            final_output = step_outputs.get(final_step)
            
            execution_time = time.time() - start_time
            
            result = PipelineExecutionResult(
                success=True,
                final_output=final_output,
                step_outputs=step_outputs
            )
            result.execution_time = execution_time
            
            if self.verbose:
                logger.info(f"Pipeline completed successfully in {execution_time:.2f}s")
                logger.info(f"Final output: {final_output}")
            
            return result
            
        except Exception as e:
            execution_time = time.time() - start_time
            error_msg = f"Pipeline execution failed: {str(e)}"
            
            if self.verbose:
                error_msg += f"\n\nDetails:\n{traceback.format_exc()}"
                logger.error(error_msg)
            
            result = PipelineExecutionResult(
                success=False,
                error=error_msg,
                step_outputs=step_outputs
            )
            result.execution_time = execution_time
            
            return result
    
    def run_single_step(self, step_number: int, input_file: Union[str, Path],
                       previous_outputs: Optional[Dict[int, str]] = None) -> PipelineExecutionResult:
        """
        Run a single pipeline step
        
        Args:
            step_number: Step to run (1-8)
            input_file: Input file for the step
            previous_outputs: Outputs from previous steps (if needed)
            
        Returns:
            PipelineExecutionResult for the single step
        """
        return self.run_pipeline(
            input_file=input_file,
            steps_to_run=[step_number]
        )
    
    def get_step_info(self, step_number: int) -> Optional[StepMetadata]:
        """Get metadata for a specific step"""
        return PipelineConfig.get_step(step_number)
    
    def validate_dependencies(self, steps_to_run: List[int]) -> Tuple[bool, List[str]]:
        """
        Validate that all step dependencies can be satisfied
        
        Returns:
            (is_valid, list_of_missing_dependencies)
        """
        missing_deps = []
        available_steps = set()
        
        for step_num in sorted(steps_to_run):
            step = PipelineConfig.get_step(step_num) 
            if not step:
                missing_deps.append(f"Unknown step: {step_num}")
                continue
                
            # Check dependencies
            for dep in step.depends_on:
                if dep != 0 and dep not in available_steps and dep not in steps_to_run:
                    missing_deps.append(f"Step {step_num} requires Step {dep}")
            
            available_steps.add(step_num)
        
        return len(missing_deps) == 0, missing_deps

# Convenience functions for backward compatibility
def run_complete_pipeline(input_file: Union[str, Path], base_dir: Optional[str] = None,
                         verbose: bool = False) -> PipelineExecutionResult:
    """Run the complete 8-step pipeline"""
    runner = PipelineRunner(base_dir=base_dir, verbose=verbose)
    return runner.run_pipeline(input_file)

def run_pipeline_with_progress(input_file: Union[str, Path], 
                              progress_callback: Callable,
                              status_callback: Optional[Callable] = None,
                              base_dir: Optional[str] = None) -> PipelineExecutionResult:
    """Run pipeline with Streamlit-style progress callbacks"""
    runner = PipelineRunner(base_dir=base_dir, verbose=False)
    return runner.run_pipeline(
        input_file=input_file,
        progress_callback=progress_callback,
        status_callback=status_callback
    )

# Export main classes
__all__ = [
    'PipelineRunner',
    'PipelineExecutionResult', 
    'ProgressTracker',
    'run_complete_pipeline',
    'run_pipeline_with_progress'
]