#!/usr/bin/env python3
"""
Test new Step 8 requirement source extraction patterns
"""

import re

def _extract_requirement_sources(text: str):
    """Test version of the updated extraction logic"""
    requirement_sources = []
    
    # Split text by common separators to handle multiple entries
    segments = re.split(r'[&,;]+', text)
    
    for segment in segments:
        segment = segment.strip()
        if not segment:
            continue
        
        # Look for complex IOS pattern first (to avoid double matching with MAT)
        # IOS with prefix: IOS-MAT-0010, IOS-PRG-0272, IOS- PRG-0273 (with spaces)
        ios_complex_pattern = r'\bIOS-\s*[A-Z]{2,4}-\d+'
        ios_complex_matches = re.findall(ios_complex_pattern, segment, re.IGNORECASE)
        
        # Look for simple IOS pattern: IOS-0123
        ios_simple_pattern = r'\bIOS-\d+'
        ios_simple_matches = re.findall(ios_simple_pattern, segment, re.IGNORECASE)
        
        # Look for MAT pattern only if no IOS-MAT found
        segment_temp = segment
        for ios_match in ios_complex_matches:
            segment_temp = segment_temp.replace(ios_match, '')
        
        mat_pattern = r'\bMAT[-]?\d+(?=[\s:]|$)'
        mat_matches = re.findall(mat_pattern, segment_temp, re.IGNORECASE)
        
        # Add all matches from this segment
        for match in ios_complex_matches + ios_simple_matches + mat_matches:
            requirement_sources.append(match.upper())
    
    # Remove duplicates while preserving order
    seen = set()
    unique_sources = []
    for source in requirement_sources:
        if source not in seen:
            seen.add(source)
            unique_sources.append(source)
    
    return unique_sources

# Test cases from user requirements
test_cases = [
    "SD MAT0250: Jiangsu Reborn",
    "SD IOS-PRG-0272 & IOS-PRG-0273", 
    "SD IOS-PRG-0272 & IOS- PRG-0273",  # User's problem case with space
    "SD IOS-MAT-0010 & IOS-MAT-0254",
    "SD IOS-MAT-0010 & IOS- MAT-0254",  # Space variant
    "MSDS of PU + FR",
    "GRS direct supplier: Wujiang Hongan",
    "SD MAT-0010 test",
    "IOS-0123 simple",
    "IOS- PRG-0273 standalone",  # Space variant standalone
    "MAT0250 MAT-0300 duplicate test",
]

print("🧪 Testing Step 7 Requirement Source Patterns:")
print()

for test_text in test_cases:
    result = _extract_requirement_sources(test_text)
    result_str = " & ".join(result) if result else "(empty)"
    print(f"Input:  '{test_text}'")
    print(f"Output: '{result_str}'")
    print()