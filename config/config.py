import sys

# Dependency availability flags
ADVANCED_OCR_AVAILABLE = False
WORD_AVAILABLE = False

def check_dependencies():
    """Check and import required dependencies."""
    global ADVANCED_OCR_AVAILABLE, WORD_AVAILABLE
    
    # Essential dependencies
    try:
        import pdfplumber
        print("pdfplumber imported successfully")
    except ImportError:
        print("pdfplumber not found. Install with: pip install pdfplumber")
        sys.exit(1)

    try:
        from openpyxl import Workbook
        print("openpyxl imported successfully")
    except ImportError:
        print("openpyxl not found. Install with: pip install openpyxl")
        sys.exit(1)
    
    # Optional advanced OCR dependencies
    try:
        import cv2
        import numpy as np
        from PIL import Image
        import pytesseract
        import fitz  # PyMuPDF
        ADVANCED_OCR_AVAILABLE = True
        print("Advanced OCR libraries available")
    except ImportError as e:
        ADVANCED_OCR_AVAILABLE = False
        print(f"Advanced OCR not available: {e}")
        print("   The extractor will work only with PDFs that already contain text")

    # Optional Word export dependency
    try:
        from docx import Document
        WORD_AVAILABLE = True
        print("python-docx available")
    except ImportError:
        WORD_AVAILABLE = False
        print("python-docx not available. Word files will not be created")

# Musical notation patterns for cleaning
MUSICAL_PATTERNS = [
    r'[:]+[-—.:]+[|]*', 
    r'[smtdfrl]+[,\']*\s*[:—.-]+\s*[smtdfrl,\']*',
    r'[|]+\s*[:.—-]+\s*[|]*',
    r'^[sd,\':\-—|.\s]+$',
    r'[drmfslti]+[,\']*\s*[:\-—.]+\s*[drmfslti,\']*', 
    r'\b[a-z][,\']+\s*[:\-—.]+', 
    r'[:\-—.]{3,}', 
    r'^[|]+.*[|]+$', 
    r'^\s*[:.—-]+\s*$', 
    r'\b[smtdfrl]+[,\']*[:\-—.]+[smtdfrl,\']*\b',  
]

# Chorus keywords
CHORUS_KEYWORDS = ['CHORUS', 'CORO', 'REFRAIN', 'CÔRO']

# OCR configuration
OCR_CONFIG = r'--oem 3 --psm 6'
OCR_LANGUAGES = 'por+eng'

# File processing settings
DEFAULT_DIRECTORY = 'pdf_books'
MAX_VERSE_LINES = 6
MAX_CHORUS_LINES = 30
MUSICAL_CHAR_THRESHOLD = 0.6