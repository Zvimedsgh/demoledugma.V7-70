import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.226.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# careful, don't replace 1550 everywhere, only in A00_SetupMainSheet
# Actually, wait, let's just do exact string replacements
content = content.replace('1550 wsMain.Range("G3").Value = 2024', '1550 If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2024')
content = content.replace('1560 If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2026', '1560 If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2025')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.226.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Fixed A00_SetupMainSheet years in 226')
