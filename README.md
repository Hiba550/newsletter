# 📰 Professional Newsletter Generator

An automated, professional newsletter/magazine generator for colleges and departments. Transform your Excel data and images into beautifully formatted Word documents with just a few clicks!

## ✨ Features

### 🎯 Two Template Options

#### 1. **Basic Template**
- Simple event listing by department
- Event descriptions, dates, and images
- Professional formatting with borders
- Quick and easy to use

#### 2. **Enhanced Template** (⭐ Recommended)
Complete magazine-style newsletter with:
- � **Cover Page** - Newsletter title, department info, editorial board
- 🎯 **Vision & Mission** - Department vision and mission statements
- � **Program Objectives (PEOs)** - Educational objectives
- 🎓 **Program Outcomes (PSOs)** - Learning outcomes
- 📑 **Table of Contents** - Professional index page
- 🎪 **Event Sections** - Multiple categorized event sections
- � **Contact Page** - Department contact information

### 🎨 Professional Styling
- ✅ Thin black border exactly 1.5cm from page edge (MS Word style)
- ✅ Times New Roman font (12pt body, 11pt headers)
- ✅ Dark blue theme (#1E3A8A) for headers
- ✅ Page numbers at bottom center
- ✅ Custom headers with newsletter info
- ✅ Professional layout with images and tables

## Quick Start

### Prerequisites
- Python 3.7+
- pip (Python package installer)

### Installation

1. **Clone or download this project**
   ```bash
   cd newsletter
   ```

2. **Install required packages**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python app.py
   ```

4. **Open your browser**
   ```
   http://localhost:5000
   ```

## How to Use

### Step 1: Download Template
- Click "Download Sample Template" to get the Excel format
- The template includes sample data to understand the structure

### Step 2: Prepare Your Data
- Fill the Excel file with your event information:
  - **event_title**: Name of the event
  - **event_description**: Detailed description
  - **event_date**: Event date (MM/DD/YYYY format)
  - **department**: Department name
  - **image_reference**: Image filename reference (e.g., "1", "2", "3")

### Step 3: Prepare Images
- Name your images as: `1.png`, `2.png`, `3.png`, etc.
- Reference them in Excel as: "1", "2", "3"
- Supported formats: PNG, JPG, JPEG, GIF

### Step 4: Upload and Generate
- Upload your Excel file and images
- Click "Generate Magazine"
- Download both PDF and DOC formats

## Excel Template Format

| Column | Description | Example |
|--------|-------------|---------|
| event_title | Event name | "Annual Tech Fest 2025" |
| event_description | Event details | "Technical festival with competitions..." |
| event_date | Event date | "2025-03-15" |
| department | Department name | "Computer Science" |
| image_reference | Image file reference | "1" (refers to 1.png) |

## Sample Data

The application comes with:
- **Sample Excel template** with 6 example events
- **Sample images** (1.png, 2.png, 3.png)
- **College logo** placeholder

## Customization

### College Information
Edit the `magazine_generator.py` file:
```python
self.college_name = "Your College Name"
self.college_logo = "static/images/your_logo.png"
```

### Styling
- Modify HTML templates in `templates/` folder
- Update CSS styles in template files
- Change colors in the CSS variables

### Image Logo
Replace `static/images/college_logo.png` with your college logo.

## File Structure
```
newsletter/
├── app.py                 # Main Flask application
├── magazine_generator.py  # Magazine generation logic
├── requirements.txt       # Python dependencies
├── templates/            # HTML templates
│   ├── base.html
│   ├── index.html
│   └── download.html
├── static/               # Static files
│   ├── images/          # Images and logo
│   └── sample_template.xlsx
├── uploads/             # Uploaded files (temporary)
└── generated/           # Generated magazines (temporary)
```

## Output Features

### PDF Format
- Professional layout with borders
- College logo on every page
- Page numbers
- Department-wise sections
- Integrated images
- Print-ready quality

### DOC Format
- Editable Word document
- Professional formatting
- Embedded images
- Department sections
- Easy to customize further

## Troubleshooting

### Common Issues

1. **"Module not found" errors**
   ```bash
   pip install -r requirements.txt
   ```

2. **Images not showing**
   - Check image file names match Excel references
   - Ensure images are uploaded along with Excel file

3. **PDF generation fails**
   - The app falls back to simple PDF generation
   - Check if all required packages are installed

4. **Excel file not reading**
   - Ensure file is in .xlsx or .xls format
   - Check if required columns exist

### Dependencies
- Flask: Web framework
- pandas: Excel file processing
- python-docx: DOC file generation
- weasyprint: PDF generation
- Pillow: Image processing
- openpyxl: Excel file reading

## Security Notes
- Uploaded files are stored temporarily
- Generated files are automatically cleaned up
- No sensitive data is permanently stored

## Support
For issues or questions, check:
1. Sample template format
2. Image naming convention
3. Required columns in Excel file

## Future Enhancements
- [ ] Multiple college templates
- [ ] Batch processing
- [ ] Email delivery
- [ ] Cloud storage integration
- [ ] Advanced styling options

---

**Built with ❤️ for educational institutions**"# newsletter" 
