with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'(?:Sub|Function)\s+([A-Za-z0-9_]+)[^\n]*\n(.*?)End (?:Sub|Function)', text, re.DOTALL | re.IGNORECASE)

for m in matches:
    name = m.group(1)
    sub_text = m.group(2)
    
    # check for xlSheetVisible or xlSheetVeryHidden or xlSheetHidden
    if 'xlSheet' in sub_text and '.Visible' in sub_text:
        print(f"--- {name} ---")
        for i, line in enumerate(sub_text.split('\n')):
            if 'xlSheet' in line and '.Visible' in line:
                print(f"  {line.strip()}")
