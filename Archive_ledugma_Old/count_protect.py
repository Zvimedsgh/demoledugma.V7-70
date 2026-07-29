with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

count = text.count('ThisWorkbook.Protect "Z961814r"') + text.count('ThisWorkbook.Protect Password:="Z961814r"')
print(f"Number of times ThisWorkbook.Protect is called: {count}")
