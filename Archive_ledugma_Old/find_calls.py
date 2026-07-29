with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Public Sub ApplyCorrectionsAndBuildReports\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for line in sub_text.split('\n'):
        if re.search(r'^\s*([A-Za-z0-9_]+)\s', line) and not line.strip().startswith("'") and not 'If ' in line and not 'End ' in line:
            # simple heuristic to find subroutine calls
            print(line.strip())
