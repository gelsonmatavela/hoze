"""Text extraction from PDF files using various methods."""

import pdfplumber
from config import ADVANCED_OCR_AVAILABLE, OCR_CONFIG, OCR_LANGUAGES
from text_cleaner import TextCleaner
from image_processor import ImageProcessor

if ADVANCED_OCR_AVAILABLE:
    import cv2
    import numpy as np
    import pytesseract
    import fitz  # PyMuPDF


class TextExtractor:
    """Handles text extraction from PDF files."""
    
    def __init__(self):
        self.text_cleaner = TextCleaner()
        self.image_processor = ImageProcessor()
        self.advanced_ocr_available = ADVANCED_OCR_AVAILABLE
    
    def extract_text_basic(self, pdf_path: str) -> str:
        """Extract text using pdfplumber."""
        print("Extracting text with pdfplumber...")
        all_text = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    print(f"   Processing page {i+1}/{len(pdf.pages)}")
                    text = page.extract_text()
                    if text and text.strip():
                        cleaned_text = self.text_cleaner.clean_musical_notation(text)
                        if cleaned_text.strip():
                            all_text.append(cleaned_text)
        except Exception as e:
            print(f"Error with pdfplumber: {e}")
            
        return '\n'.join(all_text)
    
    def extract_text_advanced(self, pdf_path: str) -> str:
        """Extract text using advanced OCR methods."""
        if not self.advanced_ocr_available:
            return self.extract_text_basic(pdf_path)
            
        print("Extracting text with advanced OCR...")
        all_text = []
        
        # First try basic method
        basic_text = self.extract_text_basic(pdf_path)
        if basic_text and len(basic_text.strip()) > 100:
            return basic_text
        
        print("   Insufficient text, applying OCR...")
        try:
            all_text = self._extract_with_ocr(pdf_path)
        except Exception as e:
            print(f"Error with OCR: {e}")
        
        return '\n'.join(all_text)
    
    def _extract_with_ocr(self, pdf_path: str) -> list:
        """Extract text using OCR on PDF images."""
        all_text = []
        
        doc = fitz.open(pdf_path)
        for page_num in range(doc.page_count):
            print(f"   OCR page {page_num+1}/{doc.page_count}")
            page = doc[page_num]
            
            # Extract as image
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            
            # Convert to numpy
            img_array = np.frombuffer(img_data, dtype=np.uint8)
            img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
            
            # Preprocess
            processed_img = self.image_processor.preprocess_image_for_ocr(img)
            
            # OCR
            text = pytesseract.image_to_string(
                processed_img, config=OCR_CONFIG, lang=OCR_LANGUAGES
            )
            
            if text.strip():
                cleaned_text = self.text_cleaner.clean_musical_notation(text)
                if cleaned_text.strip():
                    all_text.append(cleaned_text)
                    
        doc.close()
        return all_text