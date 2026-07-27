import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.024.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.025.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace("Goren_Claude_V2_024", "Goren_Claude_V2_025")
content = content.replace("V2.024", "V2.025")
content = content.replace('APP_VERSION As String = "2.024"', 'APP_VERSION As String = "2.025"')

# Add ApplyDemoLockOnOpen Sub
new_sub = """' ============================================================================
' MACRO: ApplyDemoLockOnOpen
' ============================================================================
Public Sub ApplyDemoLockOnOpen()
    Dim wsMain As Worksheet
    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then Exit Sub
    
    Dim isDemoMode As Boolean
    Dim demoParam As String
    isDemoMode = False
    demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
    
    If isDemoMode Then
        wsMain.Range("G3").Value = 2024
        wsMain.Range("G4").Value = 2025
        With wsMain.Range("G3:G4").Validation
            .Delete
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
            .ErrorMessage = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497) & " " & ChrW(1488) & ChrW(1508) & ChrW(1513) & ChrW(1512) & ChrW(1497) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492)
            .ShowError = True
        End With
        wsMain.Range("G3:G4").Interior.ColorIndex = 15 ' Grey
    Else
        ' Remove validation if not demo
        wsMain.Range("G3:G4").Validation.Delete
        wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
    End If
End Sub
"""

# Find AssignButtonMacros to insert before it
content = content.replace("Public Sub AssignButtonMacros()", new_sub + "\nPublic Sub AssignButtonMacros()")

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.025 successfully")
