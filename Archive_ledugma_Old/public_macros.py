with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Make OpenManual and OpenInstructions Public
text = re.sub(r'\bSub OpenManual\(\)', 'Public Sub OpenManual()', text)
text = re.sub(r'\bSub OpenInstructions\(\)', 'Public Sub OpenInstructions()', text)
text = re.sub(r'\bSub HideInstructions\(\)', 'Public Sub HideInstructions()', text)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Made instructions macros Public!")
