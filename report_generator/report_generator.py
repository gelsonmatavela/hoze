"""Report generation utilities for processing results."""

import json
from datetime import datetime
from typing import List, Dict


class ReportGenerator:
    """Generates processing reports and summaries."""
    
    def __init__(self):
        self.extracted_books = []
    
    def add_result(self, result: Dict):
        """Add a processing result to the report."""
        self.extracted_books.append(result)
    
    def generate_report(self):
        """Generate and display final processing report."""
        print(f"\n{'='*60}")
        print("FINAL REPORT")
        print('='*60)
        
        successful = [b for b in self.extracted_books if b.get('status') == 'SUCCESS']
        failed = [b for b in self.extracted_books if 'ERROR' in b.get('status', '')]
        
        print(f"Total files: {len(self.extracted_books)}")
        print(f"Successful: {len(successful)}")
        print(f"Failed: {len(failed)}")
        
        if successful:
            self._print_successful_results(successful)
        
        if failed:
            self._print_failed_results(failed)
        
        # Save detailed report
        report_file = self._save_detailed_report()
        print(f"\nDetailed report saved in: {report_file}")
    
    def _print_successful_results(self, successful: List[Dict]):
        """Print information about successful extractions."""
        total_songs = sum(b.get('songs_found', 0) for b in successful)
        print(f"Total songs/hymns extracted: {total_songs}")
        
        print("\nFiles created:")
        for book in successful:
            print(f"   {book['original_file']}:")
            print(f"     - {book['excel_file']}")
            print(f"     - {book['json_file']}")
            if book.get('word_file') != 'N/A':
                print(f"     - {book['word_file']}")
            print(f"     - Songs: {book['songs_found']}")
    
    def _print_failed_results(self, failed: List[Dict]):
        """Print information about failed extractions."""
        print("\nFiles with errors:")
        for book in failed:
            print(f"   {book['original_file']}: {book['status']}")
    
    def _save_detailed_report(self) -> str:
        """Save detailed report to JSON file."""
        successful = [b for b in self.extracted_books if b.get('status') == 'SUCCESS']
        failed = [b for b in self.extracted_books if 'ERROR' in b.get('status', '')]
        
        report_file = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        report_data = {
            'timestamp': datetime.now().isoformat(),
            'summary': {
                'total': len(self.extracted_books),
                'successful': len(successful),
                'failed': len(failed)
            },
            'files': self.extracted_books
        }
        
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)
        
        return report_file