with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
matches = re.finditer(r'(?:Sub|Function)\s+([A-Za-z0-9_]+)[^\n]*\n(.*?)End (?:Sub|Function)', text, re.DOTALL | re.IGNORECASE)

for m in matches:
    name = m.group(1)
    sub_text = m.group(2)
    
    if '.Visible' in sub_text:
        has_unprotect = 'Unprotect' in sub_text
        has_resume_next = 'Resume Next' in sub_text
        print(f"Sub: {name} | Unprotect: {has_unprotect} | Resume Next: {has_resume_next}")
