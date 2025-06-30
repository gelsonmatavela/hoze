"""Main book extractor class that orchestrates the extraction process."""

import os
from datetime import datetime
from typing import Dict
from config import DEFAULT_DIRECTORY, ADVANCED_OCR_AVAILABLE
from text_extrator import TextExtractor
from struture_detector import StructureDetector
from file_exporters import FileExporter
from report_generator import ReportGenerator


class BookExtractor:
    """Main class for extracting and processing PDF hymnal books."""
    
    def __init__(self, directory: str = DEFAULT_DIRECTORY):
        self.directory = directory
        self.text_extractor = TextExtractor()
        self.structure_detector = StructureDetector()
        self.file_exporter = FileExporter()
        self.report_generator = ReportGenerator()
        
    def process_files(self):
        """Process all PDF files in the directory."""
        print(f"Starting processing in: {self.directory}")
        
        if not self._setup_directory():
            return
        
        files = self._get_pdf_files()
        if not files:
            return
        
        print(f"Found {len(files)} PDF file(s)")
        
        for i, file in enumerate(files, 1):
            self._process_single_file(file, i, len(files))
        
        self.report_generator.generate_report()
    
    def _setup_directory(self) -> bool:
        """Setup processing directory."""
        if not os.path.exists(self.directory):
            os.makedirs(self.directory)
            print(f"Directory created: {self.directory}")
            print("   Place your PDF files in this directory and run again.")
            return False
        return True
    
    def _get_pdf_files(self) -> list:
        """Get list of PDF files in directory."""
        files = [f for f in os.listdir(self.directory) 
                 if f.lower().endswith('.pdf')]
        
        if not files:
            print("No PDF files found!")
            print(f"   Place PDF files in: {os.path.abspath(self.directory)}")
        
        return files
    
    def _process_single_file(self, file: str, current: int, total: int):
        """Process a single PDF file."""
        print(f"\n{'='*60}")
        print(f"Processing [{current}/{total}]: {file}")
        print('='*60)
        
        file_path = os.path.join(self.directory, file)
        
        try:
            # Extract text
            text = self._extract_text(file_path)
            if not text.strip():
                print("No text extracted")
                self._add_failed_result(file, "No text extracted")
                return
            
            print(f"Text extracted: {len(text)} characters")
            
            # Detect structure
            book_structure = self.structure_detector.detect_structure(text)
            
            # Generate output files
            self._create_output_files(file, book_structure)
            
            print("Processing completed!")
            
        except Exception as e:
            print(f"Error: {e}")
            self._add_failed_result(file, str(e))
    
    def _extract_text(self, file_path: str) -> str:
        """Extract text from PDF file."""
        if ADVANCED_OCR_AVAILABLE:
            return self.text_extractor.extract_text_advanced(file_path)
        else:
            return self.text_extractor.extract_text_basic(file_path)
    
    def _create_output_files(self, file: str, book_structure: Dict):
        """Create output files for extracted data in a dedicated folder."""
        base_name = os.path.splitext(file)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create output folder for this file
        output_folder = f"{base_name}_extracted_{timestamp}"
        output_path = os.path.join(self.directory, output_folder)
        
        try:
            os.makedirs(output_path, exist_ok=True)
            print(f"Created output folder: {output_folder}")
        except Exception as e:
            print(f"Warning: Could not create output folder: {e}")
            # Fallback to current directory
            output_path = self.directory
            output_folder = ""
        
        # Define file names (without folder path for cleaner names)
        excel_file = f"{base_name}_extracted.xlsx"
        json_file = f"{base_name}_extracted.json"
        word_file = f"{base_name}_extracted.docx"
        
        # Define full file paths
        excel_path = os.path.join(output_path, excel_file)
        json_path = os.path.join(output_path, json_file)
        word_path = os.path.join(output_path, word_file)
        
        # Create files in the output folder
        try:
            self.file_exporter.create_excel_file(book_structure, excel_path)
            self.file_exporter.create_json_file(book_structure, json_path)
            self.file_exporter.create_word_file(book_structure, word_path)
            
            print(f"Files created in folder: {output_folder}")
            print(f"  - {excel_file}")
            print(f"  - {json_file}")
            if self.file_exporter.word_available:
                print(f"  - {word_file}")
            
            # Add successful result with folder information
            self._add_successful_result(
                file, output_folder, excel_file, json_file, word_file, 
                len(book_structure.get('songs', []))
            )
            
        except Exception as e:
            print(f"Error creating files: {e}")
            self._add_failed_result(file, f"File creation error: {str(e)}")
    
    def _add_successful_result(self, original_file: str, output_folder: str,
        excel_file: str, json_file: str, word_file: str, songs_found: int):
        """Add a successful processing result."""
        result = {
            'original_file': original_file,
            'output_folder': output_folder,
            'excel_file': excel_file,
            'json_file': json_file,
            'word_file': word_file if self.file_exporter.word_available else 'N/A',
            'songs_found': songs_found,
            'status': 'SUCCESS'
        }
        self.report_generator.add_result(result)
    
    def _add_failed_result(self, original_file: str, error_message: str):
        """Add a failed processing result."""
        result = {
            'original_file': original_file,
            'status': f'ERROR: {error_message}'
        }
        self.report_generator.add_result(result)