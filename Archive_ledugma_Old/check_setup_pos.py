with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Public Sub A00_SetupMainSheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for i, line in enumerate(sub_text.split('\n')):
        if 'shp.Left' in line or 'shp.Top' in line or 'shpOpInst' in line or 'shpInstallInst' in line or 'msgOp' in line or 'msgInst' in line:
            print(f"Line {i+1}: {line.strip()}")
