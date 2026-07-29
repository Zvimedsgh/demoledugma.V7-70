with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('ThisWorkbook.Unprotect "Z961814r"\nwsMain.Unprotect "Z961814r"\nwsMgmt.Unprotect "Z961814r"', 'ThisWorkbook.Unprotect "Z961814r"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated successfully.")
