import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_fixed.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix SettingsGoToSheet visibility check
old_settings_goto = """Private Sub SettingsGoToSheet(shtName As String)
On Error Resume Next
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(shtName)
If Not ws Is Nothing Then
ws.Activate
Application.Goto ws.Range("A1")
End If
End Sub"""

new_settings_goto = """Private Sub SettingsGoToSheet(shtName As String)
On Error Resume Next
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(shtName)
If Not ws Is Nothing Then
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
ws.Activate
Application.Goto ws.Range("A1")
End If
End Sub"""

content = content.replace(old_settings_goto, new_settings_goto)

# Fix DEMO_MODE read logic in A00_SetupMainSheet
content = content.replace(
    'demoParam = UCase$(Trim$(GetStringParameter(wsMgmt, "DEMO_MODE")))',
    'demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))'
)

# Fix parameters writing logic in A00_SetupMainSheet
old_params_block = """' ---- rngFILES_FOLDER and rngREPORTS_FOLDER ----
On Error Resume Next
existingPath = Trim$(CStr(ThisWorkbook.Names("rngFILES_FOLDER").RefersToRange.Value2))
If existingPath = "" Then
ThisWorkbook.Names("rngFILES_FOLDER").Delete
Err.Clear
ThisWorkbook.Names.Add "rngFILES_FOLDER", wsMgmt.Range("B176")
End If
existingPath = Trim$(CStr(ThisWorkbook.Names("rngREPORTS_FOLDER").RefersToRange.Value2))
If existingPath = "" Then
ThisWorkbook.Names("rngREPORTS_FOLDER").Delete
Err.Clear
ThisWorkbook.Names.Add "rngREPORTS_FOLDER", wsMgmt.Range("B177")
End If
Err.Clear
On Error GoTo ERR_HANDLER"""

new_params_block = """' ---- rngFILES_FOLDER and rngREPORTS_FOLDER ----
Dim wsParams As Worksheet
On Error Resume Next
Set wsParams = ThisWorkbook.Worksheets(H_SET_PARAMS())
On Error GoTo ERR_HANDLER

On Error Resume Next
existingPath = Trim$(CStr(ThisWorkbook.Names("rngFILES_FOLDER").RefersToRange.Value2))
If existingPath = "" Then
ThisWorkbook.Names("rngFILES_FOLDER").Delete
Err.Clear
ThisWorkbook.Names.Add "rngFILES_FOLDER", wsParams.Range("B176")
End If
existingPath = Trim$(CStr(ThisWorkbook.Names("rngREPORTS_FOLDER").RefersToRange.Value2))
If existingPath = "" Then
ThisWorkbook.Names("rngREPORTS_FOLDER").Delete
Err.Clear
ThisWorkbook.Names.Add "rngREPORTS_FOLDER", wsParams.Range("B177")
End If
Err.Clear
On Error GoTo ERR_HANDLER"""

content = content.replace(old_params_block, new_params_block)

# Fix the rest of wsMgmt parameters
old_param_write_block = """' ---- Set Error_Email parameter if not exists ----
1960  paramStartRow = 1 + 1
' Find last param row (scan until EOD or empty)
1965  paramLastRow = paramStartRow
Do While UCase$(Trim$(CStr(wsMgmt.Cells(paramLastRow, COL_PARAM_NAME).Value2))) <> EOD_MARKER _
And Trim$(CStr(wsMgmt.Cells(paramLastRow, COL_PARAM_NAME).Value2)) <> ""
paramLastRow = paramLastRow + 1
Loop
paramLastRow = paramLastRow - 1  ' last actual data row
1970  foundEmail = False
1980  For pr = paramStartRow To paramLastRow
1990  If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "ERROR_EMAIL" Then foundEmail = True: Exit For
2000  Next pr
2010  If Not foundEmail Then
2020  wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "ERROR_EMAIL"
2030  wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "zvi@gorentech.co.il"
2040  paramLastRow = paramLastRow + 1
2050  End If

Dim foundDemo As Boolean
foundDemo = False
For pr = paramStartRow To paramLastRow
If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "DEMO_MODE" Then foundDemo = True: Exit For
Next pr
If Not foundDemo Then
wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "DEMO_MODE"
wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = ChrW(1500) & ChrW(1488) '
paramLastRow = paramLastRow + 1
End If

' ---- Set BACKUP_PATH parameter if not exists ----
2060  foundBackup = False
2070  For pr = paramStartRow To paramLastRow
2080  If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "BACKUP_PATH" Then foundBackup = True: Exit For
2090  Next pr
2100  If Not foundBackup Then
2110  wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "BACKUP_PATH"
2120  wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "C:\LEVAV PROJECT\BACKUPS"
2130  End If"""

new_param_write_block = old_param_write_block.replace("wsMgmt", "wsParams")
content = content.replace(old_param_write_block, new_param_write_block)

# Fix wsMgmt in BuildReview (replace with wsBranches)
old_br_setup = """Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
390     Dim dictHelper As Object"""

new_br_setup = """Dim wsBranches As Worksheet
Set wsBranches = ThisWorkbook.Worksheets(H_SET_BRANCHES())
Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
390     Dim dictHelper As Object"""
content = content.replace(old_br_setup, new_br_setup)

# The branch logic is lines 949-972. We need to replace wsMgmt with wsBranches ONLY in the branch mapping logic of BuildReview!
old_br_branch = """' Branch mapping logic
930                 brOrigName = Trim$(CStr(wsData.Cells(r, COL_BRANCH).Value2))
935                 brKey = ""
940                 If brOrigName <> "" Then
945                     Dim brEndRow As Long, brScan As Long
946                     On Error Resume Next
947                     brEndRow = 1 - 1
948                     On Error GoTo 0
949                     If brEndRow < 3 Then brEndRow = wsMgmt.Cells(wsMgmt.Rows.Count, 1).End(xlUp).Row
950                     brLastRow = 3
951                     While Trim$(CStr(wsMgmt.Cells(brLastRow, 1).Value2)) = "" And brLastRow > 3
952                         brLastRow = brLastRow - 1
953                     Wend
954                     If brLastRow < 3 Then brLastRow = 2
955                     brLastRow = brLastRow + 1
956                     
957                     For brScan = 3 To brEndRow
958                         If StrComp(Trim$(CStr(wsMgmt.Cells(brScan, 1).Value2)), brOrigName, vbTextCompare) = 0 Then
959                             brKey = Trim$(CStr(wsMgmt.Cells(brScan, 4).Value2))
960                             Exit For
961                         End If
962                     Next brScan
963                     
964                     If brKey = "" Then
965                         ' Add new unknown branch to end of management list
966                         wsMgmt.Rows(brEndRow).Insert Shift:=xlDown
967                         On Error Resume Next
968                         Err.Clear
969                         wsMgmt.Cells(brLastRow, 1).Value = brOrigName
970                         If Err.Number <> 0 Then
971                         End If
972                         wsMgmt.Cells(brLastRow, 3).Value = HebrewToKey(brOrigName)
973                         On Error GoTo ERR_HANDLER
974                         brKey = ""
975                     End If"""

new_br_branch = old_br_branch.replace("wsMgmt", "wsBranches")
content = content.replace(old_br_branch, new_br_branch)

# Fix wsMgmt in ApplyCorrectionsAndBuildReports
old_ac_setup = """580     Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
590     
600     debugStep = "CHECK_PARAMS\""""

new_ac_setup = """580     Dim wsBranches As Worksheet
Set wsBranches = ThisWorkbook.Worksheets(H_SET_BRANCHES())
Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
590     
600     debugStep = "CHECK_PARAMS\""""
content = content.replace(old_ac_setup, new_ac_setup)

content = content.replace(
    'LoadCorrectionsToDicts wsMgmt, yearVal, dictHelper, _',
    'LoadCorrectionsToDicts wsBranches, yearVal, dictHelper, _'
)
content = content.replace(
    'LoadCorrectionsToDicts wsMgmt, refYear, dictHelper, _',
    'LoadCorrectionsToDicts wsBranches, refYear, dictHelper, _'
)

# Fix LoadCorrectionsToDicts signature and body
old_load_sig = """Private Sub LoadCorrectionsToDicts(ByVal wsMgmt As Worksheet, ByVal yearStr As String, ByVal dictHelper As Object, _
ByRef dictCorrections As Object, ByRef dictIgnore As Object, _
ByRef dictHasFix As Object, ByRef dictHasIgnore As Object, ByRef dictHasUnhandled As Object, _
ByRef corrCount As Long, ByRef ignoreCount As Long, ByRef unhandledCount As Long, ByRef reviewCount As Long)"""

new_load_sig = """Private Sub LoadCorrectionsToDicts(ByVal wsBranches As Worksheet, ByVal yearStr As String, ByVal dictHelper As Object, _
ByRef dictCorrections As Object, ByRef dictIgnore As Object, _
ByRef dictHasFix As Object, ByRef dictHasIgnore As Object, ByRef dictHasUnhandled As Object, _
ByRef corrCount As Long, ByRef ignoreCount As Long, ByRef unhandledCount As Long, ByRef reviewCount As Long)"""
content = content.replace(old_load_sig, new_load_sig)

old_load_branch = """' Find the branch row in settings where col B is empty
Dim brFixRow As Long
Dim brFixLast As Long
On Error Resume Next
brFixLast = 1 - 1
On Error GoTo 0
If brFixLast < 3 Then brFixLast = wsMgmt.Cells(wsMgmt.Rows.Count, 1).End(xlUp).Row
For brFixRow = 3 To brFixLast
If Trim$(CStr(wsMgmt.Cells(brFixRow, 2).Value2)) = "" Then
' This is a branch without a main branch assigned
' Write the main branch
wsMgmt.Cells(brFixRow, 2).Value = fixText
' Find MAIN_BRANCH_KEY from existing rows with same main branch
Dim mbKeyRow As Long
For mbKeyRow = 3 To brFixLast
If StrComp(Trim$(CStr(wsMgmt.Cells(mbKeyRow, 2).Value2)), fixText, vbTextCompare) = 0 And mbKeyRow <> brFixRow Then
If Trim$(CStr(wsMgmt.Cells(mbKeyRow, 4).Value2)) <> "" Then
wsMgmt.Cells(brFixRow, 4).Value = wsMgmt.Cells(mbKeyRow, 4).Value2
Exit For
End If
End If
Next mbKeyRow
' If no existing key found, generate from fixText
If Trim$(CStr(wsMgmt.Cells(brFixRow, 4).Value2)) = "" Then
wsMgmt.Cells(brFixRow, 4).Value = HebrewToKey(fixText)
End If
Exit For
End If
Next brFixRow"""

new_load_branch = old_load_branch.replace("wsMgmt", "wsBranches")
content = content.replace(old_load_branch, new_load_branch)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Applied exact string replacements successfully.")
