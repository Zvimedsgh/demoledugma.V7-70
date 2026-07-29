with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Add Unprotect
text = re.sub(r'(Sub A00_SetupMainSheet\(\).*?On Error GoTo ERR_HANDLER)', r'\1\nThisWorkbook.Unprotect "Z961814r"\nwsMain.Unprotect "Z961814r"\nwsMgmt.Unprotect "Z961814r"', text, flags=re.DOTALL | re.IGNORECASE)

# Add Protect before ERR_HANDLER label
text = re.sub(r'(Exit Sub\s*ERR_HANDLER:)', r'ThisWorkbook.Protect "Z961814r"\n\1', text, flags=re.DOTALL)

# Fix error message
text = text.replace('MsgBoxU "Error in V2.079! (If you don\'t see this, you are running old code!)"', 'MsgBoxU "Error in V2.222!"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated successfully.")
