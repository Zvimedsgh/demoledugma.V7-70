with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

subs_to_clean = [
    'ExportTotalChart', 
    'ExportCompCharts', 
    'BuildTitleSlide', 
    'BuildTotalSlideFromImage', 
    'BuildChartSlide', 
    'BuildTableSlide'
]

for sub in subs_to_clean:
    # Find the sub block
    match = re.search(fr'(Sub {sub}\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
    if match:
        sub_text = match.group(1)
        # Remove any ThisWorkbook.Protect calls
        clean_text = re.sub(r'^\s*\d*\s*ThisWorkbook\.Protect\s+"Z961814r"\s*$', '', sub_text, flags=re.MULTILINE | re.IGNORECASE)
        # Handle cases where it's not the only thing on the line, though unlikely
        # Actually just replace literal
        clean_text = re.sub(r'\bThisWorkbook\.Protect\s+"Z961814r"', '', clean_text, flags=re.IGNORECASE)
        
        text = text.replace(sub_text, clean_text)
        print(f"Cleaned {sub}")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Removed Protects from presentation helpers!")
