with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Public Sub A00_SetupMainSheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    for line in sub_text.split('\n')[310:761]:
        if re.match(r'^\s*([A-Za-z0-9_]+)\s', line) and not 'If ' in line and not 'End ' in line and not 'Dim ' in line and not "'" in line and not 'On ' in line and not 'Set ' in line:
            # print calls
            print(line.strip())
