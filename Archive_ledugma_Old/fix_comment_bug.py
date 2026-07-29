with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Remove the line continuation from the comment
text = text.replace("Then ' _\n", "Then\n")
# Also fix any other instances where ' _ or '  _ might be used
text = re.sub(r"Then\s*'\s*_\s*\n", "Then\n", text)
text = text.replace("' _", "")
text = text.replace("'  ", "")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Fixed the comment line continuation bug!")
