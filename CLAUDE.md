# 📋 SEDO TSS Converter

## 🎯 Mục tiêu
Chuyển đổi file Excel compliance test summary từ format Input (phức tạp, nhiều merged cells) sang format Output (structured, database-ready).

**Key principle: ADAPTIVE, not HARDCODED!** 🔑

## ✅ Trước khi bắt đầu

### Kiểm tra file input
```bash
# Bước 1: Kiểm tra file trước khi chạy pipeline
python validate_my_file.py "data/input/your-file.xlsx"

# Bước 2: Chỉ tiếp tục nếu validation PASSED
```

**📋 Tài liệu hỗ trợ:**
- `INPUT_REQUIREMENTS.md` - Yêu cầu chi tiết file input
- `QUICK_CHECKLIST.md` - Checklist nhanh 5 phút  
- `EMAIL_TEMPLATE.md` - Template gửi cho users
- `pipeline_validator.py` - Comprehensive validation tool

## 🚀 Sử dụng nhanh

### Xử lý toàn bộ pipeline (8 steps)
```bash
# Complete pipeline: Input → Output final
python step1_unmerge_standalone.py data/input/input-1.xlsx
python step2_header_processing.py data/output/output-1-Step1.xlsx  
python step3_template_creation.py data/output/output-1-Step2.xlsx
python step4_article_filling.py data/input/input-1.xlsx data/output/output-1-Step3.xlsx
python step5_data_transformation.py data/output/output-1-Step2.xlsx data/output/output-1-Step4.xlsx
python step6_sd_processing.py data/output/output-1-Step2.xlsx --step4-file data/output/output-1-Step5.xlsx
python step7_finished_product.py data/input/input-1.xlsx --step6-file data/output/output-1-Step6.xlsx
python step8_document_processing.py data/input/input-1.xlsx --step7-file data/output/output-1-Step7.xlsx
```

### Xử lý 1 step riêng lẻ
```bash
# Example: Chạy riêng Step 1
python step1_unmerge_standalone.py data/input/input-1.xlsx -v

# Example: Chạy riêng Step 8 (final output)
python step8_document_processing.py data/input/input-1.xlsx --step7-file data/output/output-1-Step7.xlsx
```

## 📁 Cấu trúc project

```
/
├── 🔍 VALIDATION SYSTEM
│   ├── validate_my_file.py              # User-friendly file validator
│   ├── pipeline_validator.py            # Comprehensive pre-flight validation
│   └── validation_utils.py              # Core validation utilities & error handling
│
├── 🔄 PROCESSING PIPELINE (8 steps)
│   ├── step1_unmerge_standalone.py      # Step 1: Unmerge cells
│   ├── step2_header_processing.py       # Step 2: Process headers  
│   ├── step3_template_creation.py       # Step 3: Create template
│   ├── step4_article_filling.py         # Step 4: Fill article info
│   ├── step5_data_transformation.py     # Step 5: Transform data  
│   ├── step6_sd_processing.py           # Step 6: SD processing & de-duplication
│   ├── step7_finished_product.py        # Step 7: Article matching & validation
│   └── step8_document_processing.py     # Step 8: Final document processing
│
├── 📋 DOCUMENTATION
│   ├── INPUT_REQUIREMENTS.md            # Detailed file requirements
│   ├── QUICK_CHECKLIST.md              # 5-minute validation checklist
│   ├── EMAIL_TEMPLATE.md               # Template for users
│   └── CLAUDE.md                       # This file - developer guide
│
├── 📦 DEPENDENCIES
│   └── requirements.txt                 # Python dependencies
│
└── 📊 DATA
    ├── input/                          # Input files (Input-X.xlsx)
    └── output/                         # All outputs (Step1→Step8)
```

## 🔄 Pipeline Logic - 8 Steps Complete

### 🔍 **Pre-validation**
```bash
python validate_my_file.py "input.xlsx"  # ALWAYS run first!
```
- **Purpose**: Prevent pipeline failures by validating upfront
- **Checks**: File format, size, structure, required headers
- **Output**: PASS/FAIL with actionable error messages

### **Step 1: Unmerge Cells** 📊
```bash
python step1_unmerge_standalone.py data/input/input-X.xlsx
```
- **Input**: `data/input/Input-X.xlsx` (raw file with merged cells)
- **Output**: `data/output/output-X-Step1.xlsx`
- **Logic**: 
  - Detect all merged cell ranges
  - Preserve top-left cell values
  - Unmerge all ranges and fill empty cells
  - **Key**: Foundation step - makes data accessible

### **Step 2: Header Processing** 🎯
```bash
python step2_header_processing.py data/output/output-X-Step1.xlsx
```
- **Input**: `data/output/output-X-Step1.xlsx`
- **Output**: `data/output/output-X-Step2.xlsx`
- **Logic**: Find "General Type/Sub-Type in Connect" header → Process 3 rows below with 3-case logic:
  - **Case 1**: `val16==val17==val18` → empty, keep val17, empty
  - **Case 2**: `val16!=val17==val18` → keep val16, keep val17, empty  
  - **Case 3**: `val16!=val17!=val18` → keep val16, val17+" "+val18, empty

### **Step 3: Template Creation** 📋
```bash
python step3_template_creation.py data/output/output-X-Step2.xlsx
```
- **Input**: `data/output/output-X-Step2.xlsx`
- **Output**: `data/output/output-X-Step3.xlsx`
- **Logic**:
  - Create structured template with 17 standardized headers
  - Apply professional formatting (borders, colors, fonts)
  - Set column widths and cell alignment
  - **Purpose**: Clean, database-ready structure

### **Step 4: Article Filling** 🏷️
```bash
python step4_article_filling.py data/input/input-X.xlsx data/output/output-X-Step3.xlsx
```
- **Input**: Original input + Step3 template
- **Output**: `data/output/output-X-Step4.xlsx`
- **Logic**:
  - **Dynamic header detection**: Find "Article Name"/"Article No." headers (adaptive positioning)
  - **Multi-article extraction**: Extract multiple articles from original file
  - **Professional formatting**: Place in R+ columns, 90° rotation, orange background
  - **Boundary detection**: Only search above "General Type" header

### **Step 5: Data Transformation** 🔄
```bash
python step5_data_transformation.py data/output/output-X-Step2.xlsx data/output/output-X-Step4.xlsx
```
- **Input**: Step2 data + Step4 template
- **Output**: `data/output/output-X-Step5.xlsx`
- **Logic**:
  - **Intelligent mapping**: Map Step2 data → Step4 template structure
  - **Data preservation**: Ensure no information loss during transformation
  - **Format consistency**: Maintain template formatting while adding data

### **Step 6: SD Processing** 🔧
```bash
python step6_sd_processing.py data/output/output-X-Step2.xlsx --step4-file data/output/output-X-Step5.xlsx
```
- **Input**: Step2 + Step5
- **Output**: `data/output/output-X-Step6.xlsx`
- **Logic**:
  - **H→P column mapping**: Map H values to corresponding P column
  - **Multi-line parsing**: Handle complex SD data with line breaks
  - **Smart de-duplication**: Remove duplicates while preserving unique entries
  - **Data validation**: Ensure SD data integrity

### **Step 7: Finished Product Validation** ✅
```bash
python step7_finished_product.py data/input/input-X.xlsx --step6-file data/output/output-X-Step6.xlsx
```
- **Input**: Original input + Step6
- **Output**: `data/output/output-X-Step7.xlsx`
- **Logic**:
  - **Article matching**: Match finished products with article definitions
  - **Fuzzy matching**: Handle variations in article names (case, spacing)
  - **"All items" logic**: If P contains "All"/"All items"/"All products" → mark all articles
  - **Validation rules**: Ensure product-article consistency

### **Step 8: Document Processing** 📄
```bash
python step8_document_processing.py data/input/input-X.xlsx --step7-file data/output/output-X-Step7.xlsx
```
- **Input**: Original input + Step7
- **Output**: `data/output/output-X-Step8.xlsx` ✅ **FINAL RESULT**
- **Logic**:
  - **Requirement source extraction**: Parse complex requirement patterns (IOS, MAT, EN, etc.)
  - **Advanced pattern matching**: Handle separators (&, ,, ;) and nested requirements
  - **Document validation**: Ensure all requirements properly categorized
  - **Final quality check**: Comprehensive output validation

## 🎯 Success Criteria

Pipeline được coi là thành công khi:
1. ✅ **Pre-validation PASSED** - File input đúng format và structure
2. ✅ **All 8 steps execute** - Không có step nào fail
3. ✅ **Data integrity** - Không mất thông tin qua các step
4. ✅ **Output quality** - Step8 file đúng format, đủ data
5. ✅ **Performance** - Xử lý file 1000 rows trong <10 seconds
6. ✅ **Error handling** - Clear error messages khi có issues

## 🔧 Debug & Troubleshooting

### Validation trước khi chạy
```bash
# ALWAYS validate input first
python validate_my_file.py "data/input/your-file.xlsx" -v

# Advanced validation
python pipeline_validator.py "data/input/your-file.xlsx" -v
```

### Debug từng step
```bash
# Debug Step 1
python step1_unmerge_standalone.py data/input/input-X.xlsx -v

# Debug Step 2  
python step2_header_processing.py data/output/output-X-Step1.xlsx -v

# Debug Step 8 (final)
python step8_document_processing.py data/input/input-X.xlsx --step7-file data/output/output-X-Step7.xlsx -v
```

### Common Issues & Solutions

#### **🚨 Input Validation Failures**
- **Issue**: `"General Type header not found"`
- **Solution**: Verify "General Type/Sub-Type in Connect" exists in first 50 rows
- **Fix**: Check exact text matching, case insensitive OK

#### **🚨 Pipeline Step Failures**
- **Step 1**: Merge detection problems → check Excel file structure
- **Step 2**: Header not found → verify "General Type/Sub-Type in Connect" exists  
- **Step 4**: Article headers missing → check "Article Name"/"Article No." headers above "General Type"
- **Step 6**: Over-aggressive de-duplication → check empty columns in H→P mapping
- **Step 7**: Article matching fails → verify article definitions in original file
- **Step 8**: Pattern extraction errors → check requirement source formatting

#### **🚨 Performance Issues**
- **Large files (>50MB)**: Consider splitting into smaller chunks
- **Many merged cells (>1000)**: Step 1 may take longer, normal behavior
- **Complex SD data**: Step 6 processing time increases with data complexity

## 📊 Test Files

Đã test với các files:
- `Test1.xlsx`: Complete test case with all features
- `input-1.xlsx`: Single article, basic structure
- `input-4.xlsx`: Multiple articles  
- `input-5.xlsx`: DRÖNA case study
- `input-6.xlsx`: Different column positions
- `Drona.xlsx`: Real-world example
- `Skubb.xlsx`: Multiple articles (6 articles)
- `frakta.xlsx`: SPARKA series (5 articles)

## 🎯 Key Features

### **🔍 Validation System**
- **Pre-flight validation**: Comprehensive file checking before processing
- **Early termination**: Stop on invalid input with clear error messages
- **User guidance**: Detailed requirements documentation and tools

### **🔧 Processing Pipeline**
- **Adaptive logic**: Dynamic header detection, không hardcode positions
- **Robust unmerging**: Handles complex merged cell patterns
- **Multi-article support**: Extract multiple articles automatically
- **Smart de-duplication**: Intelligent duplicate removal
- **Advanced matching**: Fuzzy article matching with "All items" logic
- **Pattern recognition**: Complex requirement source extraction

### **🛠️ Development Features**
- **Error handling**: Structured ValidationError with actionable messages
- **Standalone tools**: Mỗi step có thể chạy độc lập
- **Comprehensive logging**: Detailed progress tracking
- **Clean architecture**: Modular, maintainable code structure

---

# 👨‍💻 Developer Guide

## 🏗️ Architecture Overview

### **Validation Layer**
```python
validation_utils.py       # Core validation classes & utilities
├── ValidationError      # Structured error handling
├── FileValidator       # Excel file validation  
├── HeaderDetector      # Dynamic header detection
└── ErrorHandler        # User-friendly error messages

pipeline_validator.py    # Comprehensive pre-flight validation
└── PipelineValidator   # Multi-stage validation workflow
```

### **Processing Layer**
```
step1_unmerge_standalone.py    → ExcelUnmerger
step2_header_processing.py     → HeaderProcessor  
step3_template_creation.py     → TemplateCreator
step4_article_filling.py       → ArticleFiller
step5_data_transformation.py   → DataTransformer
step6_sd_processing.py         → SDProcessor
step7_finished_product.py      → FinishedProductProcessor
step8_document_processing.py   → DocumentProcessor
```

### **User Interface Layer**
```
validate_my_file.py           # User-friendly validation script
INPUT_REQUIREMENTS.md         # Detailed requirements
QUICK_CHECKLIST.md           # Quick reference
EMAIL_TEMPLATE.md            # Communication template
```

## 🔧 Adding New Features

### **Adding New Validation Rules**
1. **Edit `validation_utils.py`**:
```python
class FileValidator:
    @classmethod
    def validate_new_requirement(cls, file_path: Path) -> bool:
        # Your validation logic here
        pass
```

2. **Update `pipeline_validator.py`**:
```python
def _validate_step_requirements(self, input_path: Path):
    # Add your new validation call
    if not FileValidator.validate_new_requirement(input_path):
        raise ValidationError("Your error message")
```

### **Adding New Processing Step**
1. **Create `step9_your_feature.py`**:
```python
class YourProcessor:
    def process_file(self, input_file, output_file=None):
        # Pre-flight validation
        if not validate_before_pipeline(input_file, verbose=True):
            raise ValidationError("Validation failed")
        
        # Your processing logic
        # ...
        
        return str(output_file)
```

2. **Update CLAUDE.md** với step mới
3. **Add to documentation** và test files

### **Modifying Header Detection**
Edit `validation_utils.py`:
```python
class HeaderDetector:
    @classmethod
    def find_your_header(cls, worksheet) -> Optional[Tuple[int, int, str]]:
        patterns = ["Your Header Pattern", "Alternative Pattern"]
        return cls.find_header_fuzzy(worksheet, patterns)
```

## 📦 Dependencies

```bash
pip install openpyxl  # Excel file processing
```

## 📋 Code Principles

### **🚫 DON'Ts**
- **NEVER** hardcode column positions (always use dynamic detection)
- **NEVER** assume fixed file structure (use adaptive logic)
- **NEVER** ignore errors (always handle gracefully)
- **NEVER** skip validation (pre-flight check everything)

### **✅ DOs**
- **ALWAYS** use dynamic header detection  
- **PREFER** adaptive logic over fixed patterns
- **ENSURE** data preservation at every step
- **VALIDATE** inputs before processing
- **PROVIDE** actionable error messages
- **TEST** with real-world files
- **DOCUMENT** logic changes in CLAUDE.md

### **🔄 Update Workflow**
1. **Modify code** với new feature/fix
2. **Test thoroughly** với existing test files
3. **Update CLAUDE.md** với logic changes
4. **Update documentation** nếu cần (INPUT_REQUIREMENTS.md, etc.)
5. **Commit changes** với clear message
6. **Update version** trong requirements.txt nếu cần

---

**📝 Last Updated**: 2025-12-23  
**🔧 Version**: 2.0.0 (8-step pipeline with validation system)  
**👨‍💻 Maintainer**: Check git log for contributors