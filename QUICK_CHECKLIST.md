# ✅ SEDO TSS Converter - Quick Input Checklist

## 🚀 Pre-Flight Checklist (5 phút)

### 📁 **File Basics**
- [ ] File format: `.xlsx`, `.xls`, or `.xlsm`
- [ ] File size: < 100MB
- [ ] File opens correctly in Excel
- [ ] Not password protected

### 📊 **File Content**
- [ ] **> 15 rows** of data
- [ ] **> 10 columns** of data  
- [ ] Has real data (not empty template)
- [ ] Contains merged cells (normal)

### 🔍 **Required Headers**
- [ ] **"General Type/Sub-Type in Connect"** exists ⭐ **REQUIRED**
- [ ] Header is in first 50 rows
- [ ] Text matches exactly (case insensitive OK)

### 📋 **Optional but Recommended**
- [ ] "Article Name" header exists
- [ ] "Article No." header exists
- [ ] Article headers are above "General Type" header

## 🧪 **Quick Test**
```bash
python pipeline_validator.py "your_file.xlsx" -v
```

**Expected result:** `🎯 Overall Status: ✅ PASSED`

## 🆘 **If Validation Fails**
1. **Read error message carefully**
2. **Follow suggestions provided**
3. **Check INPUT_REQUIREMENTS.md for details**
4. **Re-run validator until PASSED**

---

### 📞 **Need Help?**
- Check `INPUT_REQUIREMENTS.md` for detailed guide
- Include full error message when asking for support
- Test with validator before submitting issues

---

*🚀 Ready? Run: `python step1_unmerge_standalone.py "your_file.xlsx"`*