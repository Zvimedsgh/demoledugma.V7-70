import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Add cleanup to BuildPresentation
target_pres_exit = """Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
1060    Exit Sub"""

pres_cleanup = """' Clean up presentation sheets automatically
Application.DisplayAlerts = False
DeleteSheetIfExists SHEET_COMPANIES()
DeleteSheetIfExists SHEET_BRANCH()
DeleteSheetIfExists SHEET_MAINBRANCH()
DeleteSheetIfExists SHEET_TELLERS()
DeleteSheetIfExists SHEET_AGENTS()
DeleteSheetIfExists SHEET_MONTHS()
DeleteSheetIfExists SHEET_SUMMARY()
Application.DisplayAlerts = True

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
1060    Exit Sub"""

if target_pres_exit in content:
    content = content.replace(target_pres_exit, pres_cleanup)
else:
    print("Could not find BuildPresentation exit block!")

# Update Version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_138"', 'Attribute VB_Name = "Goren_Claude_V2_139"')
content = content.replace('VERSION: V2.138', 'VERSION: V2.139')
content = content.replace('APP_VERSION As String = "2.138"', 'APP_VERSION As String = "2.139"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.139.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.139 created.")

