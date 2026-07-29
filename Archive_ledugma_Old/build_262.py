import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.261.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.261"', 'APP_VERSION As String = "2.262"')
content = content.replace('VERSION: V2.261', 'VERSION: V2.262')
content = content.replace('Error in V2.261!', 'Error in V2.262!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_261"', 'Attribute VB_Name = "Goren_Claude_V2_262"')

# Update Changelog
changelog = """' CHANGES IN 2.262:
'   - BUGFIX: Complete fix for Error 1004 on G3:G4 Validation by applying MergeArea to Validation objects.
"""
content = content.replace("' CHANGES IN 2.261:", changelog + "' CHANGES IN 2.261:")

# Fix the With blocks
content = content.replace('With wsMain.Range("G3").Validation', 'With wsMain.Range("G3").MergeArea.Validation')
content = content.replace('With wsMain.Range("G4").Validation', 'With wsMain.Range("G4").MergeArea.Validation')

# Fix the delete blocks
content = content.replace('wsMain.Range("G3").Validation.Delete', 'wsMain.Range("G3").MergeArea.Validation.Delete')
content = content.replace('wsMain.Range("G4").Validation.Delete', 'wsMain.Range("G4").MergeArea.Validation.Delete')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.262.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 262')
