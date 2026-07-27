import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.021.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.022.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update version
content = content.replace("Goren_Claude_V2_021", "Goren_Claude_V2_022")
content = content.replace("V2.021", "V2.022")
content = content.replace('APP_VERSION As String = "2.021"', 'APP_VERSION As String = "2.022"')

# 2. Fix Agency Title Position
old_title = '''Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("A2").Left, wsMain.Range("A2").Top, 600, 40)'''
new_title = '''Set shpTitle = wsMain.Shapes.AddTextbox(1, 100, 5, 800, 40)'''
content = content.replace(old_title, new_title)

# 3. Force Years in Demo Mode
old_years = """    ' Lock years if demo mode
    If isDemoMode Then
        With wsMain.Range("G3:G4").Validation"""
        
new_years = """    ' Lock years if demo mode
    If isDemoMode Then
        wsMain.Range("G3").Value = 2024
        wsMain.Range("G4").Value = 2025
        With wsMain.Range("G3:G4").Validation"""

content = content.replace(old_years, new_years)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.022 successfully")
