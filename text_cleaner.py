import re
from config import MUSICAL_PATTERNS, MUSICAL_CHAR_THRESHOLD


class TextCleaner:
    """Handles cleaning of musical notation from extracted text."""
    
    def __init__(self):
        self.musical_patterns = MUSICAL_PATTERNS
        self.musical_char_threshold = MUSICAL_CHAR_THRESHOLD
    
    def clean_musical_notation(self, text: str) -> str:
        """Remove musical notation patterns from text."""
        print("Removing musical notations...")
        
        lines = text.split('\n')
        cleaned_lines = []
        removed_count = 0
        
        for line in lines:
            original_line = line.strip()
            
            if not original_line:
                continue
            
            # Check if it's predominantly musical notation
            if self._is_musical_notation(original_line):
                removed_count += 1
                continue
            
            # Apply pattern cleaning
            cleaned_line = self._apply_pattern_cleaning(original_line)
            
            if self._has_valid_content(cleaned_line):
                cleaned_lines.append(cleaned_line)
        
        print(f"   Removed {removed_count} lines of musical notation")
        return '\n'.join(cleaned_lines)
    
    def _is_musical_notation(self, line: str) -> bool:
        """Check if a line is predominantly musical notation."""
        musical_chars = len(re.findall(r'[:\-—.|,\'smdftlri]', line))
        total_chars = len(line.replace(' ', ''))
        
        if total_chars > 0 and (musical_chars / total_chars) > self.musical_char_threshold:
            return True
        return False
    
    def _apply_pattern_cleaning(self, line: str) -> str:
        """Apply pattern-based cleaning to a line."""
        cleaned_line = line
        for pattern in self.musical_patterns:
            cleaned_line = re.sub(pattern, '', cleaned_line, flags=re.IGNORECASE)
        
        return ' '.join(cleaned_line.split())
    
    def _has_valid_content(self, line: str) -> bool:
        """Check if a line has valid content after cleaning."""
        if not line or len(line.strip()) <= 2:
            return False
        
        return bool(re.search(r'[a-zA-Z]{3,}', line))