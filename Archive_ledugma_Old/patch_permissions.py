import os

src_path = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.10.bas'
out_path = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.11.bas'

with open(src_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version header
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_10"', 'Attribute VB_Name = "Goren_Claude_V3_11"')
content = content.replace("' VERSION: V3.10", "' VERSION: V3.11\n' CHANGES IN V3.11:\n'   - PERMISSIONS: Demo Mode now grants universal Full Access automatically, bypassing Username checks.")

# Update CheckUserPermissions
old_logic = """    ' If user not found in list, default to limited
    If Not found Then accessLevel = limitedAccess"""

new_logic = """    ' --- DEMO MODE UNIVERSAL ACCESS ---
    ' Check if we are in DEMO Mode
    Dim isSysDemo As Boolean
    isSysDemo = FORCE_DEMO_MODE
    On Error Resume Next
    Dim demoParamStr As String
    demoParamStr = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If Not FORCE_DEMO_MODE Then
        If demoParamStr = ChrW(1499) & ChrW(1503) Or demoParamStr = "YES" Then isSysDemo = True
        If demoParamStr = ChrW(1500) & ChrW(1488) Or demoParamStr = "NO" Then isSysDemo = False
    End If
    On Error GoTo PERM_ERR
    
    ' Universal Access for Demo: If user not found but system is in DEMO, grant full access
    If isSysDemo And Not found Then
        accessLevel = fullAccess
    ElseIf Not found Then
        accessLevel = limitedAccess
    End If
    ' ----------------------------------"""

content = content.replace(old_logic, new_logic)

with open(out_path, 'w', encoding='utf-8') as f:
    f.write(content)

print('V3.11 generated successfully')
