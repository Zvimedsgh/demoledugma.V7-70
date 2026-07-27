import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.204.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.205.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Add a tiny sub
new_sub = """Public Sub ApplyDemoModeValidation()
    On Error Resume Next
    ThisWorkbook.Unprotect "Z961814r"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(H_SET_PARAMS())
    If ws Is Nothing Then Exit Sub
    
    Dim pr As Long
    For pr = 2 To 100
        If UCase$(Trim$(CStr(ws.Cells(pr, 1).Value2))) = "DEMO_MODE" Then
            ws.Cells(pr, 2).Validation.Delete
            ws.Cells(pr, 2).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=ChrW(1499) & ChrW(1503) & "," & ChrW(1500) & ChrW(1488)
            MsgBoxU ChrW(1492) & ChrW(1493) & ChrW(1505) & ChrW(1507) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492), vbInformation
            Exit Sub
        End If
    Next pr
End Sub
"""

# Insert after InitRawColumns
pattern = r'(Private Sub InitRawColumns\(\)[\s\S]*?End Sub\n)'
content = __import__('re').sub(pattern, r'\1\n' + new_sub + '\n', content)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_204"', 'Attribute VB_Name = "Goren_Claude_V2_205"')
content = content.replace('VERSION: V2.204', 'VERSION: V2.205')
content = content.replace('APP_VERSION As String = "2.204"', 'APP_VERSION As String = "2.205"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.205")
