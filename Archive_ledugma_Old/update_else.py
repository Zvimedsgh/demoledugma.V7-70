with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.226.bas', 'r', encoding='utf-8') as f:
    content = f.read()

old_else = """    Else
        ' Full Mode Restorations
        wsMain.Range("G3:G4").Interior.ColorIndex = xlNone"""
        
new_else = """    Else
        ' Full Mode Restorations
        If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2024
        If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2025
        wsMain.Range("G3:G4").Interior.ColorIndex = xlNone"""

content = content.replace(old_else, new_else)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.226.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Updated 226 else block')
