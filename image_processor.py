"""Image processing utilities for OCR enhancement."""

from config import ADVANCED_OCR_AVAILABLE

if ADVANCED_OCR_AVAILABLE:
    import cv2
    import numpy as np


class ImageProcessor:
    """Handles image preprocessing for better OCR results."""
    
    def __init__(self):
        self.available = ADVANCED_OCR_AVAILABLE
    
    def preprocess_image_for_ocr(self, image_array) -> any:
        """Preprocess image for better OCR results."""
        if not self.available:
            return image_array
            
        # Convert to grayscale
        if len(image_array.shape) == 3:
            gray = cv2.cvtColor(image_array, cv2.COLOR_RGB2GRAY)
        else:
            gray = image_array
            
        # Bilateral filter
        bilateral = cv2.bilateralFilter(gray, 9, 75, 75)
        
        # Adaptive threshold
        thresh = cv2.adaptiveThreshold(
            bilateral, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
            cv2.THRESH_BINARY, 11, 2
        )
        
        # Morphological operations
        kernel = np.ones((1, 1), np.uint8)
        opening = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)
        closing = cv2.morphologyEx(opening, cv2.MORPH_CLOSE, kernel)
        
        # Denoising
        denoised = cv2.fastNlMeansDenoising(closing)
        
        return denoised