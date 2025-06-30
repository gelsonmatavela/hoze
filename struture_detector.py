"""Text structure detection for identifying songs, verses, and choruses."""

import re
from datetime import datetime
from typing import Dict, List
from config import CHORUS_KEYWORDS, MAX_VERSE_LINES, MAX_CHORUS_LINES


class StructureDetector:
    """Detects the structure of hymnal text (songs, verses, choruses)."""
    
    def __init__(self):
        self.chorus_keywords = CHORUS_KEYWORDS
        self.max_verse_lines = MAX_VERSE_LINES
        self.max_chorus_lines = MAX_CHORUS_LINES
    
    def detect_structure(self, text: str) -> Dict:
        """Detect and organize text structure into songs, verses, and choruses."""
        print("Detecting text structure...")
        
        lines = text.split('\n')
        structure = {
            'title': '',
            'songs': [],
            'metadata': {
                'total_lines': len([l for l in lines if l.strip()]),
                'extraction_date': datetime.now().isoformat()
            }
        }
        
        current_song = None
        i = 0
        
        while i < len(lines):
            line = lines[i].strip()
            if not line:
                i += 1
                continue
            
            # Detect song/hymn numbers with titles
            song_match = re.match(r'^(\d+)\s+(.+)', line)
            if song_match and len(song_match.group(2)) > 3:
                if self._is_song_title(lines, i):
                    current_song = self._create_new_song(
                        song_match, structure, current_song
                    )
                    i += 1
                    continue
            
            # Detect main title
            if self._is_main_title(line, structure, current_song):
                structure['title'] = line
                i += 1
                continue
            
            # Detect chorus/refrain
            if self._is_chorus_line(line):
                if current_song:
                    chorus_lines, next_i = self._extract_chorus(lines, i)
                    current_song['chorus'] = chorus_lines
                    i = next_i
                    continue
            
            # Detect numbered verses
            verse_match = re.match(r'^(\d+)\s+(.+)', line)
            if verse_match and current_song:
                verse_data, next_i = self._extract_numbered_verse(lines, i, verse_match)
                if verse_data:
                    current_song['verses'].append(verse_data)
                i = next_i
                continue
            
            # Detect unnumbered verses
            if current_song and self._is_potential_verse(line):
                verse_data, next_i = self._extract_unnumbered_verse(lines, i, current_song)
                if verse_data:
                    current_song['verses'].append(verse_data)
                i = next_i
                continue
            
            i += 1
        
        # Add last song
        if current_song:
            structure['songs'].append(current_song)
        
        # Filter valid songs
        structure['songs'] = self._filter_valid_songs(structure['songs'])
        
        self._print_structure_summary(structure)
        return structure
    
    def _is_song_title(self, lines: List[str], i: int) -> bool:
        """Check if a numbered line is a song title or verse."""
        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if re.match(r'^\d+\s', next_line):
                # Count consecutive numbered lines
                numbered_sequence = 0
                for j in range(i, min(i + 5, len(lines))):
                    if re.match(r'^\d+\s', lines[j].strip()):
                        numbered_sequence += 1
                # If 3+ numbered lines in sequence, likely verses
                if numbered_sequence >= 3:
                    return False
        return True
    
    def _create_new_song(self, song_match, structure: Dict, current_song: Dict) -> Dict:
        """Create a new song entry."""
        # Save previous song
        if current_song:
            structure['songs'].append(current_song)
        
        # Create new song
        return {
            'number': song_match.group(1),
            'title': song_match.group(2).strip(),
            'verses': [],
            'chorus': []
        }
    
    def _is_main_title(self, line: str, structure: Dict, current_song: Dict) -> bool:
        """Check if line is the main title of the book."""
        return (line.isupper() and len(line) > 10 and 
                not structure['title'] and not current_song)
    
    def _is_chorus_line(self, line: str) -> bool:
        """Check if line indicates start of chorus."""
        return any(keyword in line.upper() for keyword in self.chorus_keywords)
    
    def _extract_chorus(self, lines: List[str], i: int) -> tuple:
        """Extract chorus lines starting from position i."""
        chorus_lines = []
        j = i + 1
        
        while j < len(lines) and j < i + self.max_chorus_lines:
            chorus_line = lines[j].strip()
            if not chorus_line:
                j += 1
                continue
            
            # Stop conditions
            if re.match(r'^\d+\s', chorus_line) and len(chorus_line.split()) > 3:
                break
            if self._is_chorus_line(chorus_line):
                break
            
            chorus_lines.append(chorus_line)
            j += 1
        
        return chorus_lines, j
    
    def _extract_numbered_verse(self, lines: List[str], i: int, verse_match) -> tuple:
        """Extract a numbered verse."""
        verse_num = verse_match.group(1)
        verse_text = verse_match.group(2)
        
        if len(verse_text.split()) <= 15:  # Likely start of verse
            verse_lines = [verse_text]
            j = i + 1
            
            while j < len(lines) and len(verse_lines) < self.max_verse_lines:
                next_line = lines[j].strip()
                if not next_line:
                    j += 1
                    continue
                
                # Stop conditions
                if re.match(r'^\d+\s', next_line):
                    break
                if self._is_chorus_line(next_line):
                    break
                
                verse_lines.append(next_line)
                j += 1
            
            # Return verse if it has valid content
            if len(' '.join(verse_lines).strip()) > 10:
                return {
                    'number': verse_num,
                    'lines': verse_lines
                }, j
        
        return None, i + 1
    
    def _is_potential_verse(self, line: str) -> bool:
        """Check if line could start an unnumbered verse."""
        words = line.split()
        return (len(line) > 5 and not re.match(r'^\d+', line) and 
                len(words) > 1 and not line.isupper())
    
    def _extract_unnumbered_verse(self, lines: List[str], i: int, current_song: Dict) -> tuple:
        """Extract an unnumbered verse."""
        verse_lines = [lines[i].strip()]
        j = i + 1
        
        while j < len(lines) and len(verse_lines) < 4:
            next_line = lines[j].strip()
            if not next_line:
                j += 1
                continue
            
            # Stop conditions
            if re.match(r'^\d+\s', next_line) and len(next_line.split()) > 3:
                break
            if self._is_chorus_line(next_line):
                break
            
            verse_lines.append(next_line)
            j += 1
        
        # Add verse if it's the first or has substantial content
        if (len(current_song['verses']) == 0 or 
            len(' '.join(verse_lines)) > 15):
            return {
                'number': str(len(current_song['verses']) + 1),
                'lines': verse_lines
            }, j
        
        return None, j
    
    def _filter_valid_songs(self, songs: List[Dict]) -> List[Dict]:
        """Filter out songs without valid content."""
        return [song for song in songs if song.get('verses') or song.get('chorus')]
    
    def _print_structure_summary(self, structure: Dict):
        """Print summary of detected structure."""
        print(f"   Found {len(structure['songs'])} songs/hymns")
        for song in structure['songs'][:3]:  # Show first 3 for debugging
            title = song.get('title', 'No title')[:30]
            verses_count = len(song.get('verses', []))
            chorus_count = len(song.get('chorus', []))
            print(f"     {song.get('number', 'N/A')}. {title}... - "
                  f"{verses_count} verses, {chorus_count} chorus lines")