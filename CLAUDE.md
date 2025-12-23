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

## 🚀 Sử dụng nhanh

### Xử lý 1 file
```bash
# Input: data/input/input-X.xlsx → Output: data/output/output-X-Step6.xlsx
python step6_article_filling.py data/input/input-1.xlsx --step5-file data/output/output-1-Step5.xlsx
```

### Xử lý toàn bộ pipeline
```bash
# Chạy lần lượt từ Step 1 → Step 6
python step1_unmerge_standalone.py data/input/input-1.xlsx
python step2_header_processing.py data/output/output-1-Step1.xlsx  
python step3_template_creation.py data/output/output-1-Step2.xlsx
python step4_data_transformation.py data/output/output-1-Step2.xlsx data/output/output-1-Step3.xlsx
python step5_sd_processing.py data/output/output-1-Step2.xlsx --step4-file data/output/output-1-Step4.xlsx
python step6_article_filling.py data/input/input-1.xlsx --step5-file data/output/output-1-Step5.xlsx
```

## 📁 Cấu trúc project

```
/
├── step1_unmerge_standalone.py      # Step 1: Unmerge cells
├── step2_header_processing.py       # Step 2: Process headers  
├── step3_template_creation.py       # Step 3: Create template
├── step4_data_transformation.py     # Step 4: Transform data
├── step5_sd_processing.py           # Step 5: SD processing & de-duplication
├── step6_article_filling.py         # Step 6: Article name/number filling
├── requirements.txt                 # Dependencies
└── data/
    ├── input/                      # Input files (Input-X.xlsx)
    └── output/                     # All outputs (Step1→Step6)
```

## 🔄 Pipeline hoàn chỉnh

Converter thực hiện 6 bước tuần tự:

### Step 1: Unmerge Cells
- **Input**: `data/input/Input-X.xlsx` 
- **Output**: `data/output/output-X-Step1.xlsx`
- **Logic**: Unmerge tất cả merged cells, preserve data

### Step 2: Header Processing  
- **Input**: `data/output/output-X-Step1.xlsx`
- **Output**: `data/output/output-X-Step2.xlsx`
- **Logic**: Xử lý header với 3-case logic sau "General Type/Sub-Type in Connect"

### Step 3: Template Creation
- **Input**: `data/output/output-X-Step2.xlsx`
- **Output**: `data/output/output-X-Step3.xlsx`  
- **Logic**: Tạo structured template với 17 headers có formatting

### Step 4: Data Transformation
- **Input**: Step2 + Step3
- **Output**: `data/output/output-X-Step4.xlsx`
- **Logic**: Transform data từ Step2 vào template Step3

### Step 5: SD Processing
- **Input**: Step2 + Step4
- **Output**: `data/output/output-X-Step5.xlsx`
- **Logic**: Xử lý SD data, multi-line parsing, de-duplication

### Step 6: Article Filling
- **Input**: Original input + Step5
- **Output**: `data/output/output-X-Step6.xlsx` ✅ **FINAL**
- **Logic**: Dynamic header detection, extract article name/number

## 🎯 Success Criteria

Pipeline được coi là thành công khi:
1. ✅ 100% test cases pass
2. ✅ Output đúng format, đủ data, không miss information
3. ✅ Performance: xử lý file 1000 rows trong <5 seconds
4. ✅ Error messages rõ ràng, actionable
5. ✅ Code clean, documented, maintainable

## 🔧 Debug & Troubleshooting

### Kiểm tra từng step
Nếu pipeline fail, check từng bước:

```bash
# Debug Step 1
python step1_unmerge_standalone.py data/input/input-X.xlsx -v

# Debug Step 2  
python step2_header_processing.py data/output/output-X-Step1.xlsx -v

# Debug Step 3
python step3_template_creation.py data/output/output-X-Step2.xlsx -v

# etc...
```

### Common issues
- **Step 1**: Merge detection problems → check Excel file structure
- **Step 2**: Header not found → verify "General Type/Sub-Type in Connect" exists
- **Step 5**: Over-aggressive de-duplication → check empty columns
- **Step 6**: Article headers not found → verify "Article Name"/"Article No." headers

## 📊 Test Files

Đã test với các files:
- `input-1.xlsx`: Single article
- `input-4.xlsx`: Multiple articles  
- `input-5.xlsx`: DRÖNA case
- `input-6.xlsx`: Different column positions
- `Drona.xlsx`: Real-world example
- `Skubb.xlsx`: Multiple articles (6 articles)
- `frakta.xlsx`: SPARKA series (5 articles)

## 🎯 Key Features

- **Adaptive logic**: Dynamic header detection, không hardcode positions
- **Robust unmerging**: Handles complex merged cell patterns
- **Multi-article support**: Extract multiple articles automatically
- **De-duplication**: Smart duplicate removal
- **Error handling**: Clear error messages và recovery options
- **Standalone tools**: Mỗi step có thể chạy độc lập

---

# Developer Notes

## Dependencies
```bash
pip install openpyxl
```

## Code principles
- **NEVER** hardcode column positions
- **ALWAYS** use dynamic header detection  
- **PREFER** adaptive logic over fixed patterns
- **ENSURE** data preservation at every step