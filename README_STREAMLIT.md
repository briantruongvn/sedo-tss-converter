# 🌐 Ngoc Son Internal TSS Converter - Web Application

## 🚀 Live Demo
[Access the web application here](https://your-streamlit-app-url.streamlit.app)

## 📋 Features

### 🎯 **User-Friendly Interface**
- **Drag & Drop Upload**: Easy file upload with validation
- **Real-time Progress**: Track processing through 8 steps
- **Instant Download**: Download converted file immediately
- **Responsive Design**: Works on desktop and mobile

### 🔧 **Complete Pipeline Integration**
- ✅ **Pre-flight Validation**: File checking before processing
- ✅ **8-Step Processing**: Complete SEDO TSS conversion pipeline
- ✅ **Error Handling**: Clear error messages with solutions
- ✅ **Progress Tracking**: Visual progress bar and step indicators

### 📊 **File Support**
- **Formats**: .xlsx, .xls, .xlsm
- **Size Limit**: 200MB maximum
- **Validation**: Comprehensive file structure checking

## 🖥️ Local Development

### Prerequisites
```bash
pip install -r requirements.txt
```

### Run Locally
```bash
streamlit run app.py
```

The app will be available at: http://localhost:8501

## 🌐 Streamlit Cloud Deployment

### 1. **Push to GitHub**
```bash
git add .
git commit -m "Add Streamlit web application"
git push origin main
```

### 2. **Deploy to Streamlit Cloud**
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Connect your GitHub account
3. Select repository: `your-repo/SEDO Internal TSS Converter`
4. Set main file: `app.py`
5. Click "Deploy"

### 3. **Configuration**
The app includes:
- `.streamlit/config.toml` - Streamlit configuration
- `requirements.txt` - Python dependencies
- Custom CSS styling for professional appearance

## 🎨 UI Components

### **Header Section**
- Professional title with chart icon
- Clean subtitle explaining functionality
- Sans-serif typography (Inter font)

### **Upload Section** 
- Large drag-and-drop area
- File format and size constraints
- Visual feedback on hover

### **Processing Section**
- Step-by-step progress indicator
- Real-time status updates
- Progress bar with gradient styling

### **Results Section**
- Success/error messages
- File download button
- Processing summary

## 🔧 Customization

### **Styling**
Edit the CSS in `app.py` to customize:
- Colors and gradients
- Font families and sizes
- Layout and spacing
- Button and component styling

### **Processing Logic**
The app integrates the complete 8-step pipeline:
1. Input validation
2. Cell unmerging
3. Header processing
4. Template creation
5. Article filling
6. Data transformation
7. SD processing
8. Finished product validation
9. Document processing

### **Error Handling**
- File validation before processing
- Graceful error messages
- Processing failure recovery
- User-friendly error descriptions

## 📱 Mobile Responsive

The app is designed to work on:
- ✅ Desktop browsers
- ✅ Tablet devices
- ✅ Mobile phones
- ✅ Touch interfaces

## 🔒 Security Features

- File size limits (200MB)
- File type validation
- Temporary file processing
- No data persistence
- Secure file handling

## 📈 Performance

- Efficient pipeline processing
- Progress tracking for user feedback
- Memory management for large files
- Optimized for Streamlit Cloud

## 🆘 Troubleshooting

### **Common Issues**
1. **File Upload Fails**
   - Check file format (.xlsx, .xls, .xlsm)
   - Verify file size < 200MB
   - Ensure file is not corrupted

2. **Processing Errors**
   - File may be missing required headers
   - Check INPUT_REQUIREMENTS.md for file structure
   - Verify file contains "General Type/Sub-Type in Connect" header

3. **Deployment Issues**
   - Check requirements.txt for missing dependencies
   - Verify GitHub repository is public
   - Ensure app.py is in repository root

## 📞 Support

- **Documentation**: See INPUT_REQUIREMENTS.md
- **Quick Check**: Use validate_my_file.py locally
- **Issues**: Check error messages for specific guidance

---

**🔧 Built with Streamlit • 📊 Powered by SEDO TSS Pipeline • ❤️ Made with Love**