# 📧 Email Template for Users

## Subject: SEDO TSS Converter - Input File Requirements

---

Chào [TÊN USER],

Để đảm bảo **SEDO TSS Converter** chạy thành công với file Excel của bạn, vui lòng **kiểm tra các yêu cầu sau** trước khi submit:

## ✅ **Quick Checklist (5 phút)**

### 📁 **File Requirements:**
- File format: `.xlsx`, `.xls`, hoặc `.xlsm`
- File size: < 100MB  
- File mở được trong Excel
- Không bị password protected

### 📊 **Content Requirements:**
- **> 15 dòng** dữ liệu thực tế
- **> 10 cột** dữ liệu
- **Header "General Type/Sub-Type in Connect"** có mặt ⭐ **BẮT BUỘC**
- Header nằm trong **50 dòng đầu tiên**

## 🧪 **Self-Test Command**
Trước khi gửi file, hãy chạy lệnh test này:

```bash
python pipeline_validator.py "path/to/your/file.xlsx" -v
```

**Kết quả mong đợi:** `🎯 Overall Status: ✅ PASSED`

## 📋 **Documents đính kèm:**
- `INPUT_REQUIREMENTS.md` - Hướng dẫn chi tiết đầy đủ
- `QUICK_CHECKLIST.md` - Checklist nhanh để in ra

## 🚨 **Important Notes:**
1. **Chạy validator trước** - Điều này sẽ tiết kiệm thời gian cho cả hai bên
2. **Nếu có lỗi** - Đọc kỹ error message và follow instructions  
3. **Gửi file chỉ khi** validator báo "PASSED"

## 📞 **Support:**
Nếu gặp vấn đề với validation:
- Gửi **full error message** từ validator
- Attach **sample file** (nếu không sensitive)
- Mô tả **workflow** bạn đã thử

---

**Cảm ơn bạn đã dành thời gian kiểm tra! Điều này giúp pipeline chạy smooth hơn rất nhiều.** 🙏

Best regards,  
[TÊN BẠN]

---

### 📎 Attachments:
- INPUT_REQUIREMENTS.md
- QUICK_CHECKLIST.md