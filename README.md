# PDF Hymnal and Book Extractor

![Python](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Status](https://img.shields.io/badge/status-active-brightgreen.svg)

A powerful Python tool for extracting hymns and songs from PDF files, with intelligent structure detection and multiple output formats.

## 🎵 Features

- **Smart Text Extraction**: Uses both basic PDF text extraction and advanced OCR when needed
- **Musical Notation Cleaning**: Automatically removes musical notation patterns from extracted text
- **Intelligent Structure Detection**: Recognizes hymn numbers, titles, verses, and choruses
- **Multiple Output Formats**: Exports to Excel, JSON, and Word documents
- **Batch Processing**: Processes multiple PDF files at once
- **Comprehensive Reporting**: Generates detailed reports of extraction results

## 📋 Requirements

### Required Dependencies
```bash
pip install pdfplumber
pip install openpyxl
```

### Optional Dependencies (for Advanced OCR)
```bash
pip install opencv-python
pip install numpy
pip install Pillow
pip install pytesseract
pip install PyMuPDF
```

### Optional Dependencies (for Word Export)
```bash
pip install python-docx
```

## 🚀 Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/pdf-hymnal-extractor.git
cd pdf-hymnal-extractor
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. (Optional) Install Tesseract OCR for advanced text recognition:
   - **Windows**: Download from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki)
   - **macOS**: `brew install tesseract`
   - **Linux**: `sudo apt-get install tesseract-ocr`

## 🎯 Usage

1. **Prepare your files**: Place your PDF hymnal files in the `pdf_books` directory (created automatically on first run)

2. **Run the extractor**:
```bash
python hymnal_extractor.py
```

3. **Check results**: The tool will create multiple output files for each processed PDF:
   - `filename_extracted_TIMESTAMP.xlsx` - Excel spreadsheet
   - `filename_extracted_TIMESTAMP.json` - JSON structured data
   - `filename_extracted_TIMESTAMP.docx` - Word document (if available)

## 📊 Output Formats

### Excel Output
- **Columns**: Number, Title, Type, Verse, Content
- **Types**: MAIN TITLE, TITLE, VERSE, CHORUS
- **Auto-sized columns** for better readability

### JSON Output
```json
{
  "metadata": {
    "extraction_date": "2024-01-01T12:00:00",
    "total_songs": 150,
    "title": "Hymnal Book Title"
  },
  "content": {
    "songs": [
      {
        "number": "1",
        "title": "Amazing Grace",
        "verses": {
          "stanza_1": {
            "number": "1",
            "lines": ["Amazing grace how sweet the sound..."]
          }
        },
        "chorus": ["How sweet the sound..."]
      }
    ]
  }
}
```

### Word Output
- **Formatted document** with proper headings
- **Structured layout** with clear verse separation
- **Page breaks** between songs

## 🔧 How It Works

### 1. Text Extraction
- **Primary method**: Uses `pdfplumber` for PDFs with embedded text
- **Fallback method**: Advanced OCR using `pytesseract` and `OpenCV` for scanned PDFs
- **Image preprocessing**: Applies filters and noise reduction for better OCR accuracy

### 2. Musical Notation Cleaning
Automatically removes common musical notation patterns:
- Chord symbols and notation marks
- Rhythm patterns
- Musical directions
- Staff notation remnants

### 3. Structure Detection
Intelligently identifies:
- **Hymn numbers and titles** (e.g., "1 Amazing Grace")
- **Verse numbers** and content
- **Chorus/refrain sections**
- **Main book titles**

### 4. Quality Filtering
- Removes lines with predominantly musical symbols
- Validates verse content length and structure
- Ensures meaningful text extraction

## 📁 Directory Structure

```
pdf-hymnal-extractor/
├── hymnal_extractor.py      # Main extractor script
├── pdf_books/               # Input directory for PDF files
├── requirements.txt         # Python dependencies
├── README.md               # This file
└── output/                 # Generated files appear here
    ├── *.xlsx              # Excel files
    ├── *.json              # JSON files
    ├── *.docx              # Word files
    └── report_*.json       # Processing reports
```

## ⚙️ Configuration

The extractor can be customized by modifying parameters in the `BookExtractor` class:

- **Directory**: Change the input directory path
- **OCR Settings**: Adjust OCR confidence and language settings
- **Pattern Matching**: Modify musical notation cleaning patterns
- **Structure Detection**: Customize verse and chorus detection rules

## 🐛 Troubleshooting

### Common Issues

1. **No text extracted**:
   - Ensure PDF contains text or install OCR dependencies
   - Check if PDF is password-protected

2. **OCR not working**:
   - Install Tesseract OCR system dependency
   - Verify image processing libraries are installed

3. **Poor structure detection**:
   - Check if hymnal follows standard numbering conventions
   - Manual adjustment of detection patterns may be needed

### Debug Information

The tool provides verbose output showing:
- Dependency availability
- Processing progress
- Text extraction statistics
- Structure detection results

## 📝 Example Output

```
Processing [1/3]: church_hymnal.pdf
Extracting text with pdfplumber...
   Processing page 1/250
   ...
Removing musical notations...
   Removed 45 lines of musical notation
Detecting text structure...
   Found 127 songs/hymns
     1. Amazing Grace - 4 verses, 8 chorus lines
     2. How Great Thou Art - 3 verses, 4 chorus lines
     ...
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [pdfplumber](https://github.com/jsvine/pdfplumber) for PDF text extraction
- [pytesseract](https://github.com/madmaze/pytesseract) for OCR capabilities
- [OpenCV](https://opencv.org/) for image processing
- [openpyxl](https://openpyxl.readthedocs.io/) for Excel file generation

## 📞 Support

If you encounter any issues or have questions, please:
1. Check the troubleshooting section above
2. Open an issue on GitHub
3. Provide sample PDF files (if possible) for debugging

---

**Note**: This tool works best with well-formatted hymnals that follow standard numbering conventions. Results may vary depending on the original PDF quality and structure.
