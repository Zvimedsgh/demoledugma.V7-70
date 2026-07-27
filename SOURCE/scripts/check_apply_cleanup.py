import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add cleanup to BuildPresentation
pres_cleanup = """
    ' Clean up presentation sheets
    Application.DisplayAlerts = False
    DeleteSheetIfExists SHEET_COMPANIES()
    DeleteSheetIfExists SHEET_BRANCH()
    DeleteSheetIfExists SHEET_MAINBRANCH()
    DeleteSheetIfExists SHEET_TELLERS()
    DeleteSheetIfExists SHEET_AGENTS()
    DeleteSheetIfExists SHEET_MONTHS()
    DeleteSheetIfExists SHEET_SUMMARY()
    Application.DisplayAlerts = True
"""

# Find the end of BuildPresentation
# It ends with:
#     Application.ScreenUpdating = True
#     Application.DisplayAlerts = True
#     MsgBoxU ...
# End Sub

target_pres_end = """    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    1080    MsgBoxU"""

new_pres_end = pres_cleanup + "\n" + target_pres_end

if target_pres_end in content:
    content = content.replace(target_pres_end, new_pres_end)
else:
    print("Could not find end of BuildPresentation")

# 2. Add cleanup to SaveReportsToFolder
# After saving the new workbook, delete the client sheets from ThisWorkbook
# Wait, let's look at the end of SaveReportsToFolder
