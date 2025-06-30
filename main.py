from config import check_dependencies, ADVANCED_OCR_AVAILABLE, WORD_AVAILABLE
from book_extrator import BookExtractor


def print_banner():
    """Print application banner and version info."""
    print("="*60)
    print("PDF HYMNAL AND BOOK EXTRACTOR")
    print("="*60)
    print("Version 2.1 - Modular Architecture")
    print()


def print_dependency_status():
    """Print status of optional dependencies."""
    print("Checking dependencies...")
    
    if ADVANCED_OCR_AVAILABLE:
        print("   Advanced OCR: Available")
    else:
        print("   Advanced OCR: Not available")
        print("     (Works only with PDFs that already contain text)")
    
    if WORD_AVAILABLE:
        print("   Word Export: Available")
    else:
        print("   Word Export: Not available")
    
    print()


def main():
    """Main application entry point."""
    print_banner()
    
    # Check and import dependencies
    check_dependencies()
    print_dependency_status()
    
    # Create and run extractor
    extractor = BookExtractor('pdf_books')
    extractor.process_files()
    
    print("\nProcess finished!")
    print("Check the generated files in the current directory")


if __name__ == '__main__':
    main()