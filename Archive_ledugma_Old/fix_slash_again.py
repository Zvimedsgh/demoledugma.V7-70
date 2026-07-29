with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

# Fix the bad backslash
text = text.replace("RGB(255, 153, 153) \\'", "RGB(255, 153, 153) '")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Removed all bad backslashes.")
