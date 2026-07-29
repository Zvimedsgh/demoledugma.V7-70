import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.00.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Speed up PowerPoint by minimizing it, and use AppActivate after MsgBox
pp_init_old = """    615     ppApp.Visible = True
    620     Set ppPres = ppApp.Presentations.Add"""

pp_init_new = """    615     ppApp.Visible = True
    On Error Resume Next
    ppApp.WindowState = 2 ' Minimized for speed
    On Error GoTo ERR_HANDLER
    620     Set ppPres = ppApp.Presentations.Add"""

content = content.replace(pp_init_old, pp_init_new)

msgbox_old = """    MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation
    
1060    Exit Sub"""

msgbox_new = """    MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation
    
    ' Bring PowerPoint to the front after user clicks OK
    On Error Resume Next
    AppActivate "PowerPoint"
    On Error GoTo ERR_HANDLER
    
1060    Exit Sub"""

content = content.replace(msgbox_old, msgbox_new)

# 2. Add Yes/No Validation to H_SET_PARAMS B7 in A00_SetupMainSheet
validation_code = """    ' Add Yes/No Dropdown in Settings B7
    On Error Resume Next
    Dim wsParamsSetup As Worksheet
    Set wsParamsSetup = ThisWorkbook.Worksheets(H_SET_PARAMS())
    If Not wsParamsSetup Is Nothing Then
        wsParamsSetup.Unprotect Password:="1234"
        With wsParamsSetup.Range("B7").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=ChrW(1499) & ChrW(1503) & "," & ChrW(1500) & ChrW(1488)
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        wsParamsSetup.Protect Password:="1234"
    End If
    On Error GoTo 0

    ' Auto-detect screen width and set zoom"""

content = content.replace("    ' Auto-detect screen width and set zoom", validation_code)


content = content.replace('APP_VERSION As String = "3.00"', 'APP_VERSION As String = "3.01"')
content = content.replace('VERSION: V3.00', 'VERSION: V3.01')
content = content.replace('Error in V3.00!', 'Error in V3.01!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_00"', 'Attribute VB_Name = "Goren_Claude_V3_01"')

changelog = """' CHANGES IN 3.01:
'   - UI: Sped up PowerPoint creation by minimizing it during generation, and added AppActivate to bring it to front after the MsgBox.
'   - UI: Added Yes/No Data Validation dropdown to cell B7 in the Parameters sheet during setup.
"""
content = content.replace("' CHANGES IN 3.00:", changelog + "' CHANGES IN 3.00:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.01.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.01')
