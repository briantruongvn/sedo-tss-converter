# 📋 SEDO TSS Converter

**Excel Compliance Test Summary Converter** - Transforms complex Excel files with merged cells into clean, structured, database-ready format.

## 🎯 Overview

Chuyển đổi file Excel compliance test summary từ format Input (phức tạp, nhiều merged cells) sang format Output (structured, database-ready).

**Key principle: ADAPTIVE, not HARDCODED!** 🔑

## 🚀 Quick Start

### Single File Processing
```bash
# Complete pipeline for one file
python step1_unmerge_standalone.py data/input/Test1.xlsx
python step2_header_processing.py data/output/Test1-Step1.xlsx  
python step3_template_creation.py data/output/Test1-Step2.xlsx
python step4_article_filling.py data/input/Test1.xlsx --step3-file data/output/Test1-Step3.xlsx
python step5_data_transformation.py data/output/Test1-Step2.xlsx data/output/Test1-Step3.xlsx
python step6_sd_processing.py data/output/Test1-Step2.xlsx --step5-file data/output/Test1-Step5.xlsx
python step7_finished_product.py data/output/Test1-Step6.xlsx
python step8_document_processing.py data/output/Test1-Step7.xlsx
```

### Result
- **Input**: `data/input/Test1.xlsx` (complex Excel with merged cells)
- **Output**: `data/output/Test1-Step8.xlsx` (clean structured data)

## 📁 Project Structure

```
SEDO Internal TSS Converter/
├── step1_unmerge_standalone.py      # Step 1: Unmerge cells
├── step2_header_processing.py       # Step 2: Process headers  
├── step3_template_creation.py       # Step 3: Create template
├── step4_article_filling.py         # Step 4: Article extraction
├── step5_data_transformation.py     # Step 5: Data transformation
├── step6_sd_processing.py           # Step 6: SD processing & de-duplication
├── step7_finished_product.py        # Step 7: Finished product processing
├── step8_document_processing.py     # Step 8: Document processing & cleanup
├── requirements.txt                 # Dependencies
├── CLAUDE.md                        # Detailed documentation
└── data/
    ├── input/                      # Input files
    │   ├── Test1.xlsx              # Sample test file
    │   └── Test Summary of CIRKUSTÄLT*.xlsx
    └── output/                     # Generated output files
        └── .gitkeep               # Keep directory in git
```

## 🔄 8-Step Pipeline

The converter processes files through 8 sequential steps:

| Step | Function | Input | Output | Purpose |
|------|----------|-------|---------|----------|
| **1** | Cell Unmerging | Raw Excel | Unmerged Excel | Remove merged cells, preserve data |
| **2** | Header Processing | Step 1 | Processed headers | Apply 3-case logic to headers |
| **3** | Template Creation | Step 2 | Structured template | Create 17-column template |
| **4** | Article Filling | Original + Step 3 | Article info | Extract article names/numbers |
| **5** | Data Transformation | Step 2 + Step 3 | Transformed data | H→P mapping, data population |
| **6** | SD Processing | Step 2 + Step 5 | Deduplicated data | SD processing, remove duplicates |
| **7** | Finished Product | Step 6 | Article matched | Process finished products, article matching |
| **8** | Document Processing | Step 7 | **FINAL OUTPUT** | Document specs, cleanup column P |

## ✨ Key Features

- **🔄 Adaptive Logic**: Dynamic header detection, no hardcoded positions
- **🛡️ Robust Unmerging**: Handles complex merged cell patterns
- **📝 Multi-Article Support**: Extract multiple articles automatically
- **🧹 Smart De-duplication**: Intelligent duplicate removal
- **🎯 Article Matching**: Supports "All", "All items", "All products" patterns
- **📊 H→P Mapping**: Consistent data transformation across steps
- **🧼 Data Cleanup**: Document type/requirement source extraction
- **⚙️ Standalone Tools**: Each step can run independently

## 📊 Tested Files

Successfully processed:
- ✅ **Test1.xlsx**: 2 articles, 410 final rows
- ✅ **CIRKUSTÄLT files**: 4 articles, 414 final rows
- ✅ Various edge cases and formats

## 🔧 Requirements

```bash
pip install openpyxl
```

## 🐛 Troubleshooting

### Common Issues
- **Step 2**: Header not found → verify "General Type/Sub-Type in Connect" exists
- **Step 4**: Article headers not found → check "Article Name"/"Article No." headers
- **Step 6**: Over-aggressive de-duplication → check for empty columns
- **Step 7**: Article matching issues → verify column P content and article headers

### Debug Mode
Add `-v` flag to any step for verbose logging:
```bash
python step1_unmerge_standalone.py data/input/Test1.xlsx -v
```

## 🔄 Development

Current branch: `robustness-improvements` 
- Adding enhanced error handling
- Improving input validation
- Better user experience

Main branch: `main`
- Stable working pipeline
- Tested with multiple file formats

## 📄 License

Internal tool for SEDO TSS processing.

---

**🤖 Generated with Claude Code**