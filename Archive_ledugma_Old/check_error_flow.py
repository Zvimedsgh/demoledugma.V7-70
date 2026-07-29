with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Public Sub ApplyCorrectionsAndBuildReports\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for i, line in enumerate(sub_text.split('\n')):
        if 'BUILD_COMPANIES' in line or 'BUILD_MONTHS' in line or 'BUILD_SUMMARY' in line or 'Resume Next' in line or 'GoTo ERR_HANDLER' in line:
            print(f"Line {i+1}: {line.strip()}")
