# 📋 SEDO TSS Converter - Input File Requirements

## 🎯 Mục đích
Tài liệu này mô tả **yêu cầu chi tiết** cho file input để đảm bảo pipeline chạy thành công. Vui lòng đọc kỹ và kiểm tra file của bạn trước khi chạy converter.

---

## ✅ Yêu cầu cơ bản

### 📁 **Định dạng file**
- **Format**: Excel (.xlsx, .xls, .xlsm)
- **Kích thước**: Tối đa 100MB (khuyến nghị < 50MB)
- **Encoding**: UTF-8 hoặc Excel standard
- **Trạng thái**: File không bị khóa, không password protected

### 📊 **Kích thước tối thiểu**
- **Số dòng**: Tối thiểu 15 rows (khuyến nghị > 30)
- **Số cột**: Tối thiểu 10 columns (khuyến nghị > 15) 
- **Dữ liệu**: Phải có dữ liệu thực tế trong 10 dòng/cột đầu tiên

---

## 🔍 Yêu cầu cấu trúc bắt buộc

### 1️⃣ **Header "General Type/Sub-Type in Connect"** ⭐ **BẮT BUỘC**
```
✅ Đúng: "General Type/Sub-Type in Connect"
✅ Được chấp nhận: "General Type of Material in Connect"
❌ Sai: "General Type", "Material Type", "Connect Type"
```

- **Vị trí**: Trong 50 dòng đầu tiên
- **Định dạng**: Text chính xác, không viết tắt
- **Lưu ý**: Header này là **điểm mốc quan trọng** cho pipeline

### 2️⃣ **Article Headers** (Tùy chọn nhưng khuyến nghị)
```
✅ Article Name: "Article Name", "article name", "Product Name"
✅ Article No.: "Article No.", "Article No", "Product No", "Art No"
```

- **Vị trí**: Trên header "General Type" (tránh xung đột)
- **Cặp đôi**: Nếu có Article Name thì nên có Article Number
- **Khoảng cách**: Cùng dòng hoặc gần nhau

---

## 📋 Checklist trước khi chạy

### ✅ **File Validation Checklist**
Hãy kiểm tra các mục sau trước khi submit:

- [ ] **File tồn tại** và có thể mở được trong Excel
- [ ] **Không có lỗi** khi mở file (không corrupted)
- [ ] **File size < 100MB** (kiểm tra thuộc tính file)
- [ ] **Có dữ liệu thực tế** (không phải file rỗng hoặc template)

### ✅ **Structure Validation Checklist**
- [ ] **Header "General Type/Sub-Type in Connect"** có mặt
- [ ] **Header nằm trong 50 dòng đầu**
- [ ] **File có > 15 dòng dữ liệu**
- [ ] **File có > 10 cột dữ liệu**
- [ ] **Có merged cells** (bình thường cho input files)

### ✅ **Content Validation Checklist**
- [ ] **Article information** được điền (nếu có)
- [ ] **Không có special characters** gây lỗi encoding
- [ ] **Các cells quan trọng không bị ẩn**
- [ ] **File không bị password protection**

---

## 🚀 Cách kiểm tra nhanh

### **Method 1: Sử dụng Pipeline Validator**
```bash
python pipeline_validator.py "path/to/your/file.xlsx" -v
```

**Output mong đợi:**
```
✅ file_validation: File validation passed
✅ excel_structure: Excel structure valid  
✅ step2_check: General Type header found
✅ step4_check: Article headers found (optional)
✅ system_resources: Sufficient disk space
🎯 Overall Status: ✅ PASSED
```

### **Method 2: Manual Check in Excel**
1. **Mở file trong Excel**
2. **Tìm header "General Type/Sub-Type in Connect"** (Ctrl+F)
3. **Kiểm tra file size** (File Properties)
4. **Đếm số dòng/cột có dữ liệu**

---

## ⚠️ Common Issues & Solutions

### ❌ **"General Type header not found"**
**Nguyên nhân:**
- Header text không chính xác
- Header nằm quá sâu (> 50 dòng)
- Header bị merge với cell khác

**Giải pháp:**
- Kiểm tra chính tả: `"General Type/Sub-Type in Connect"`
- Di chuyển header lên trên (< 50 dòng)
- Unmerge cells chứa header

### ❌ **"File too small" errors**
**Nguyên nhân:**
- File chỉ có template, không có dữ liệu
- Dữ liệu bị ẩn hoặc trong sheets khác

**Giải pháp:**
- Đảm bảo có > 15 dòng dữ liệu thực tế
- Unhide các dòng/cột bị ẩn
- Chuyển sang sheet chính có dữ liệu

### ❌ **"Invalid Excel file" errors**
**Nguyên nhân:**
- File bị corrupted
- Sai format (không phải Excel)
- File đang mở trong ứng dụng khác

**Giải pháp:**
- Re-save file trong Excel (.xlsx format)
- Đóng file trong tất cả ứng dụng
- Kiểm tra file integrity

### ❌ **"Permission denied" errors**
**Nguyên nhân:**
- File đang mở trong Excel
- Không có quyền đọc file
- File nằm trong thư mục protected

**Giải pháp:**
- Đóng Excel trước khi chạy
- Copy file sang thư mục khác
- Chạy với quyền administrator

---

## 📊 File Examples

### ✅ **Good Example Structure**
```
Row 1-5:   [Company info, dates, etc.]
Row 6:     Article name | xxx | Article No. | xxx | ...
Row 7-10:  [Article data rows]
...
Row 15:    General Type/Sub-Type in Connect | xxx | ...
Row 16+:   [Main data with merged cells]
```

### ❌ **Bad Example Structure**
```
Row 1:     Just headers without context
Row 2:     General Type/Material Connect  ← Sai text
Row 3:     [Empty rows]
Row 60:    General Type/Sub-Type in Connect  ← Quá sâu
```

---

## 🆘 Troubleshooting Workflow

### **Step 1: Pre-validation**
1. Run `python pipeline_validator.py "your_file.xlsx" -v`
2. Nếu PASS → Proceed to pipeline
3. Nếu FAIL → Xem error messages và fix

### **Step 2: Fix Issues**
1. **File errors** → Check file format, size, permissions
2. **Structure errors** → Check dimensions, headers
3. **Header errors** → Verify required headers exist
4. **System errors** → Check disk space, permissions

### **Step 3: Re-validate**
1. Fix issues theo suggestions
2. Run validator lại
3. Lặp lại cho đến khi PASS

### **Step 4: Run Pipeline**
```bash
python step1_unmerge_standalone.py "your_file.xlsx"
```

---

## 📞 Support

### **Nếu vẫn gặp lỗi:**
1. **Gửi error message đầy đủ** (copy từ console)
2. **Gửi file sample** (nếu không sensitive)
3. **Mô tả workflow** bạn đã thực hiện

### **Information cần cung cấp:**
- Tên file và size
- Error message từ validator
- Excel version đang sử dụng
- Operating system (Windows/Mac)

---

## 🎯 Success Criteria

File của bạn **sẵn sàng** khi:
- ✅ Pipeline validator báo "PASSED"
- ✅ Không có CRITICAL errors
- ✅ Tất cả required headers được tìm thấy
- ✅ File structure hợp lệ

**→ Sau đó có thể chạy pipeline an toàn!**

---

*📝 Cập nhật lần cuối: 2025-12-23*
*🔧 Version: 1.0.0*