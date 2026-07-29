with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.57.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix the literal \n strings that got injected
content = content.replace(r'\n', '\n')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_57"', 'Attribute VB_Name = "Goren_Claude_V3_58"')

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.58.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print("V3.58 created with fixed syntax!")
