# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**SEDO TSS Converter** - Transforms complex Excel compliance test summary files (with merged cells) into clean, structured, database-ready format through an 8-step processing pipeline.

**Core Principle**: ADAPTIVE, not HARDCODED. Use dynamic header detection and fuzzy matching, never assume fixed column positions or file structure.

## Quick Commands

### Validation (ALWAYS run first)
```bash
# User-friendly validation
python validate_my_file.py "data/input/your-file.xlsx"

# Comprehensive pre-flight validation
python pipeline_validator.py "data/input/your-file.xlsx" -v
```

### Pipeline Execution

**Recommended: Centralized Pipeline Runner**
```python
from pipeline_runner import run_complete_pipeline

result = run_complete_pipeline("data/input/input-1.xlsx", verbose=True)
if result.success:
    print(f"✅ Output: {result.final_output}")
else:
    print(f"❌ Error: {result.error}")
```

**Web Interface**
```bash
streamlit run app.py
```

**Manual CLI Execution** (8 steps)
```bash
python step1_unmerge_standalone.py data/input/input-X.xlsx
python step2_header_processing.py data/output/output-X-Step1.xlsx
python step3_template_creation.py data/output/output-X-Step2.xlsx
python step4_article_filling.py data/input/input-X.xlsx data/output/output-X-Step3.xlsx
python step5_data_transformation.py data/output/output-X-Step2.xlsx data/output/output-X-Step4.xlsx
python step6_sd_processing.py data/output/output-X-Step2.xlsx --step4-file data/output/output-X-Step5.xlsx
python step7_finished_product.py data/input/input-X.xlsx --step6-file data/output/output-X-Step6.xlsx
python step8_document_processing.py data/input/input-X.xlsx --step7-file data/output/output-X-Step7.xlsx
```

### Debug Individual Steps
```bash
# Add -v flag for verbose output
python step1_unmerge_standalone.py data/input/input-X.xlsx -v
python step8_document_processing.py data/input/input-X.xlsx --step7-file data/output/output-X-Step7.xlsx -v
```

### Dependencies
```bash
pip install -r requirements.txt
```

## Architecture Overview

### Centralized Configuration System (v3.0.0+)

**Single Point of Truth**: All pipeline metadata lives in `pipeline_config.py`. Changes propagate automatically to CLI and Streamlit interfaces.

```
pipeline_config.py          # Step definitions, metadata, dependencies
├── StepMetadata           # Dataclass for step configuration
├── PipelineConfig         # Central configuration manager
└── PipelineConstants      # System-wide constants

pipeline_runner.py         # Unified execution engine
├── PipelineRunner         # Orchestrates step execution
├── ProgressTracker        # Progress reporting for CLI/web
└── PipelineExecutionResult # Result wrapper
```

**Key Design**: Each step can be executed via:
1. CLI script directly (`python stepX_*.py`)
2. Pipeline runner (`PipelineRunner.run_pipeline()`)
3. Streamlit interface (`app.py`)

All three methods use the same underlying step classes.

### Validation Layer

**Pre-flight validation prevents pipeline failures**

```
validation_utils.py
├── ValidationError        # Structured exception with severity/category
├── FileValidator          # Excel file validation
├── HeaderDetector         # Fuzzy header matching
└── ErrorHandler           # User-friendly error messages

pipeline_validator.py      # Comprehensive pre-validation
└── PipelineValidator     # Multi-stage validation workflow

validate_my_file.py        # User-facing validation script
```

### Processing Pipeline

**8 sequential steps with dependency management:**

```
Step 1: ExcelUnmerger          (step1_unmerge_standalone.py)
  Input: Raw Excel with merged cells
  Output: Unmerged Excel

Step 2: HeaderProcessor        (step2_header_processing.py)
  Input: Step 1 output
  Output: Processed headers (3-case logic)
  Critical: Finds "General Type/Sub-Type in Connect" header

Step 3: TemplateCreator        (step3_template_creation.py)
  Input: Step 2 output
  Output: 17-column structured template

Step 4: ArticleFiller          (step4_article_filling.py)
  Input: Original Excel + Step 3 template
  Output: Template with article info (R+ columns)
  Critical: Dynamic detection of "Article Name"/"Article No." headers

Step 5: DataTransformer        (step5_data_transformation.py)
  Input: Step 2 + Step 4 outputs
  Output: Transformed data (Step 2 → Step 4 structure)

Step 6: SDProcessor            (step6_sd_processing.py)
  Input: Step 2 + Step 5 outputs
  Output: SD processed data with de-duplication
  Critical: H→P column mapping, multi-line parsing

Step 7: FinishedProductProcessor (step7_finished_product.py)
  Input: Original Excel + Step 6 output
  Output: Article-matched products
  Features: Fuzzy matching, "All items" logic

Step 8: DocumentProcessor      (step8_document_processing.py)
  Input: Original Excel + Step 7 output
  Output: FINAL processed file
  Features: Requirement source extraction, document validation
```

**Step Dependencies:**
- Step 1: No dependencies
- Step 2: Requires Step 1
- Step 3: Requires Step 2
- Step 4: Requires Step 3 + original input
- Step 5: Requires Step 2 + Step 4
- Step 6: Requires Step 2 + Step 5
- Step 7: Requires Step 6 + original input
- Step 8: Requires Step 7 + original input

### Data Flow Pattern

```
Original Input (merged cells)
    ↓
[Step 1] → output-X-Step1.xlsx (unmerged)
    ↓
[Step 2] → output-X-Step2.xlsx (headers processed)
    ↓
[Step 3] → output-X-Step3.xlsx (template created)
    ↓
[Step 4] → output-X-Step4.xlsx (articles filled)
    ← Original Input
    ↓
[Step 5] → output-X-Step5.xlsx (data transformed)
    ← Step 2 output
    ↓
[Step 6] → output-X-Step6.xlsx (SD processed)
    ← Step 2 output
    ↓
[Step 7] → output-X-Step7.xlsx (products matched)
    ← Original Input
    ↓
[Step 8] → output-X-Step8.xlsx ✅ FINAL
    ← Original Input
```

## Critical Implementation Rules

### NEVER Do These:
- **NEVER** hardcode column positions (use `HeaderDetector.find_*` methods)
- **NEVER** assume fixed file structure (use fuzzy matching)
- **NEVER** skip pre-validation (run `validate_my_file.py` or `validate_before_pipeline()`)
- **NEVER** ignore ValidationError exceptions (they contain actionable messages)

### ALWAYS Do These:
- **ALWAYS** use dynamic header detection via `HeaderDetector`
- **ALWAYS** validate inputs before processing
- **ALWAYS** preserve data at every step (no information loss)
- **ALWAYS** use `ValidationError` with helpful suggestions
- **ALWAYS** test with real-world files from `data/input/`
- **ALWAYS** update `pipeline_config.py` when adding/modifying steps

## Adding New Pipeline Steps

**Example: Adding Step 9**

1. **Define metadata in `pipeline_config.py`:**
```python
StepMetadata(
    step_number=9,
    name="your_step_name",
    display_name="Processing your feature",
    description="What this step does",
    class_name="YourProcessor",
    module_name="step9_your_feature",
    requires_original_input=False,  # or True
    depends_on=[8],  # Which steps must run first
    cli_script="step9_your_feature.py",
    estimated_duration_seconds=5
)
```

2. **Create `step9_your_feature.py` with standard interface:**
```python
class YourProcessor:
    def __init__(self, base_dir: str = "."):
        self.base_dir = Path(base_dir)

    @classmethod
    def get_metadata(cls):
        return PipelineConfig.get_step(9)

    def process_file(self, input_file, output_file=None):
        # Validate input
        input_path = validate_pipeline_input(input_file, "Step 9")

        # Process data
        # ...

        # Return output path
        return str(output_path)
```

3. **Update `pipeline_runner.py` execution logic:**
Add handling in `_execute_step()` method for step 9.

4. **Changes automatically propagate to:**
- CLI help text
- Streamlit progress display
- Dependency validation
- Progress estimation

## Key Headers to Detect

The pipeline expects these headers in input files:

1. **"General Type/Sub-Type in Connect"** (Step 2)
   - Variants: "General Type of Material in Connect"
   - Search: First 50 rows
   - Method: `HeaderDetector.find_general_type_header()`

2. **"Article Name" / "Article No."** (Step 4)
   - Search: First 20 rows, above "General Type" header
   - Method: `HeaderDetector.find_article_headers()`
   - Must find both in same row

## Testing

Tested files located in `data/input/`:
- `Test1.xlsx` - Complete test case with all features
- `input-1.xlsx` - Single article, basic structure
- `input-4.xlsx` - Multiple articles
- `input-5.xlsx` - DRÖNA case study
- `input-6.xlsx` - Different column positions
- `Drona.xlsx` - Real-world example
- `Skubb.xlsx` - Multiple articles (6 articles)
- `frakta.xlsx` - SPARKA series (5 articles)

Test any changes against multiple files to ensure adaptive logic works.

## Troubleshooting Common Issues

**Input Validation Failures:**
- `"General Type header not found"` → Check header exists in first 50 rows
- File not found → Verify path and permissions
- Invalid Excel → Check file format (.xlsx, .xls, .xlsm)

**Step-Specific Issues:**
- **Step 1**: Merged cell detection → check Excel structure
- **Step 2**: Header processing → verify "General Type/Sub-Type in Connect" exists
- **Step 4**: Article extraction → check "Article Name"/"Article No." above "General Type"
- **Step 6**: Over-aggressive dedup → check H→P mapping logic
- **Step 7**: Article matching fails → verify article definitions and column P
- **Step 8**: Pattern extraction errors → check requirement source formatting

**Debug Commands:**
```bash
# Verbose validation
python validate_my_file.py "file.xlsx" -v
python pipeline_validator.py "file.xlsx" -v

# Verbose step execution
python step1_unmerge_standalone.py input.xlsx -v
```

## File Naming Convention

Pipeline generates standardized output names:
```
Input:  data/input/MyFile.xlsx
Step 1: data/output/MyFile-Step1.xlsx
Step 2: data/output/MyFile-Step2.xlsx
...
Step 8: data/output/MyFile-Step8.xlsx (FINAL)
```

## Version History

- **v3.0.0** (2026-01-04): Centralized configuration system, single point of truth
- **v2.x**: Complete 8-step pipeline with all features
- **v1.x**: Initial implementation

## Documentation Files

- `CLAUDE.md` - This file (developer guide)
- `README.md` - User-facing documentation
- `INPUT_REQUIREMENTS.md` - Detailed input file requirements
- `QUICK_CHECKLIST.md` - 5-minute validation checklist
- `EMAIL_TEMPLATE.md` - User communication template
- `DEPLOYMENT_GUIDE.md` - Streamlit deployment instructions
