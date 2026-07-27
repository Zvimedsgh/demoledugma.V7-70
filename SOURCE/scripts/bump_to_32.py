import sys
import shutil

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.031.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.032.bas'

shutil.copy(filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("Goren_Claude_V2_031", "Goren_Claude_V2_032")
content = content.replace("V2.031", "V2.032")
content = content.replace('APP_VERSION As String = "2.031"', 'APP_VERSION As String = "2.032"')

# 1. Add IsParameterSheet function
func_is_param = """Private Function IsParameterSheet(ByVal sName As String) As Boolean
    Dim shtNames As Variant
    shtNames = Array(H_SET_FIELDMAP(), H_SET_BRANCHES(), H_SET_PARAMS(), H_SET_PERIODS(), H_SET_MESSAGES(), H_SET_PERMISSIONS(), H_SET_REASONS(), H_SET_CLIENTS())
    Dim i As Long
    For i = LBound(shtNames) To UBound(shtNames)
        If StrComp(sName, shtNames(i), vbTextCompare) = 0 Then
            IsParameterSheet = True
            Exit Function
        End If
    Next i
    IsParameterSheet = False
End Function

Private Sub SettingsGoToSheet"""

content = content.replace("Private Sub SettingsGoToSheet", func_is_param)

# 2. Modify SettingsGoToSheet
old_goto = """Private Sub SettingsGoToSheet(shtName As String)
On Error Resume Next
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(shtName)
If Not ws Is Nothing Then
ws.Activate
Application.Goto ws.Range("A1")
End If
End Sub"""

new_goto = """Private Sub SettingsGoToSheet(shtName As String)
On Error Resume Next
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(shtName)
If Not ws Is Nothing Then
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
ws.Activate
Application.Goto ws.Range("A1")
End If
End Sub"""

content = content.replace(old_goto, new_goto)

# 3. Modify NavSettings_Menu
old_nav = """Public Sub NavSettings_Menu()
10  Dim wsMgmt As Worksheet
20  On Error Resume Next
30  Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
40  On Error GoTo 0
50  If wsMgmt Is Nothing Then Exit Sub

60  wsMgmt.Activate
70  Application.GoTo Reference:=wsMgmt.Range("A1"), Scroll:=True
End Sub"""

new_nav = """Public Sub NavSettings_Menu()
10  Dim wsMgmt As Worksheet
20  On Error Resume Next
30  Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
40  On Error GoTo 0
50  If wsMgmt Is Nothing Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If IsParameterSheet(ws.Name) Then
            If ws.Visible <> xlSheetHidden Then ws.Visible = xlSheetHidden
        End If
    Next ws
    On Error GoTo 0

60  wsMgmt.Activate
70  Application.GoTo Reference:=wsMgmt.Range("A1"), Scroll:=True
End Sub"""

content = content.replace(old_nav, new_nav)

# 4. Modify ShowHiddenSheets
old_show = """Public Sub ShowHiddenSheets()
' Shows all hidden sheets except MATACH
Dim ws As Worksheet
On Error Resume Next
For Each ws In ThisWorkbook.Worksheets
If ws.Name <> MATACH_SHEET_NAME() Then
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
End If
Next ws
On Error GoTo 0"""

new_show = """Public Sub ShowHiddenSheets()
' Shows all hidden sheets except MATACH
Dim ws As Worksheet
On Error Resume Next
For Each ws In ThisWorkbook.Worksheets
If ws.Name <> MATACH_SHEET_NAME() Then
    If IsParameterSheet(ws.Name) Then
        If ws.Visible <> xlSheetHidden Then ws.Visible = xlSheetHidden
    Else
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If
Next ws
On Error GoTo 0"""

content = content.replace(old_show, new_show)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.032 successfully")
