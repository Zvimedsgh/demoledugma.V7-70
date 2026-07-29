with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Public Sub ApplyCorrectionsAndBuildReports\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    lines = sub_text.split('\n')
    start_print = False
    for i, line in enumerate(lines):
        if 'debugStep = "BUILD_COMPANIES"' in line:
            start_print = True
        if start_print:
            print(f"Line {i+1}: {line.strip()}")
        if 'debugStep = "BUILD_BRANCH"' in line:
            # print a few more lines
            for j in range(1, 10):
                print(f"Line {i+1+j}: {lines[i+j].strip()}")
            break
