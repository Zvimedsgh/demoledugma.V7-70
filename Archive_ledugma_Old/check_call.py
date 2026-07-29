with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
match = re.search(r'(Sub A00_SetupMainSheet\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
if match:
    sub_text = match.group(1)
    if 'SetupInstructionsSheet' in sub_text:
        print("A00_SetupMainSheet calls SetupInstructionsSheet!")
    else:
        print("No, it doesn't.")
