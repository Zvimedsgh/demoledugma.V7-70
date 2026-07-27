import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.70.bas', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Update Version
text = re.sub(r'Attribute VB_Name = "Goren_Claude1_69"', 'Attribute VB_Name = "Goren_Claude1_70"', text)
text = re.sub(r'VERSION: V1\.69', 'VERSION: V1.70', text)

# 2. Add version changes log
changelog = '''' CHANGES IN 1.70:
'   - BUGFIX: Fixed exit button Hebrew spelling from Yeshia to Yetzia.
'   - BUGFIX: Added Unprotect/Protect to ResetHomeDefaults.
' CHANGES IN 1.69:'''
text = re.sub(r'''\' CHANGES IN 1\.69:''', changelog, text)

# 3. Fix Exit System button Hebrew
text = re.sub(r'ChrW\(1497\) & ChrW\(1513\) & ChrW\(1497\) & ChrW\(1488\) & ChrW\(1492\)', 'ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1488) & ChrW(1492)', text)

# 4. Modify ResetHomeDefaults to Unprotect and Protect
reset_func_old = '''Public Sub ResetHomeDefaults()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    Application.EnableEvents = True
    Application.EnableEvents = False'''
    
reset_func_new = '''Public Sub ResetHomeDefaults()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    On Error Resume Next
    wsMain.Unprotect "Z961814r"
    On Error GoTo 0
    
    Application.EnableEvents = True
    Application.EnableEvents = False'''
text = text.replace(reset_func_old, reset_func_new)

reset_end_old = '''    wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)
    
    Application.EnableEvents = True
    
    Application.Goto wsMain.Range("A1")
End Sub'''

reset_end_new = '''    wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)
    
    Application.EnableEvents = True
    
    On Error Resume Next
    wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True
    On Error GoTo 0
    
    Application.Goto wsMain.Range("A1")
End Sub'''
text = text.replace(reset_end_old, reset_end_new)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.70.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Patched successfully")
