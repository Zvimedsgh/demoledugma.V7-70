with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Fix ShowResultSheets
new_show = """Sub ShowResultSheets()
' Shows only the result sheets (companies, branches, agents, etc.)
Dim ws As Worksheet
Dim resNames As String
resNames = "|" & SHEET_COMPANIES() & "|" & SHEET_AGENTS() & "|" & SHEET_BRANCH() & "|" & SHEET_MAINBRANCH() & "|" & SHEET_TELLERS() & "|" & SHEET_MONTHS() & "|" & SHEET_SUMMARY() & "|"

On Error Resume Next
ThisWorkbook.Unprotect "Z961814r"

ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Visible = xlSheetVisible

For Each ws In ThisWorkbook.Worksheets
If InStr(1, resNames, "|" & ws.Name & "|", vbTextCompare) > 0 Then
ws.Visible = xlSheetVisible
End If
Next ws

' Go to the first result sheet if it exists
If SheetExists(SHEET_COMPANIES()) Then
ThisWorkbook.Worksheets(SHEET_COMPANIES()).Activate
End If
ThisWorkbook.Protect "Z961814r"
On Error GoTo 0
End Sub"""
text = re.sub(r'Sub ShowResultSheets\(\).*?End Sub', new_show, text, flags=re.DOTALL | re.IGNORECASE)

# Fix HideInstallationInstructions
new_hide_inst = """Sub HideInstallationInstructions()
Dim wsName1 As String, wsName2 As String
wsName1 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) ' _
wsName2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) '  

On Error Resume Next
ThisWorkbook.Unprotect "Z961814r"

Dim wsInst As Worksheet
Set wsInst = ThisWorkbook.Worksheets(wsName1)
If wsInst Is Nothing Then Set wsInst = ThisWorkbook.Worksheets(wsName2)
If wsInst Is Nothing Then Set wsInst = ThisWorkbook.Worksheets("Sheet2")

If Not wsInst Is Nothing Then
wsInst.Visible = xlSheetVeryHidden
End If

' Save state in AA1 of Main sheet
ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("AA1").Value = "YES"
ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
ThisWorkbook.Protect "Z961814r"
ThisWorkbook.Save
On Error GoTo 0
End Sub"""
text = re.sub(r'Sub HideInstallationInstructions\(\).*?End Sub', new_hide_inst, text, flags=re.DOTALL | re.IGNORECASE)

# Fix HideInstallInst
new_hide2 = """Sub HideInstallInst()
On Error Resume Next
ThisWorkbook.Unprotect "Z961814r"
ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("AA1").Value = "YES"
Dim sInstall As String
sInstall = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
ThisWorkbook.Worksheets(sInstall).Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
ThisWorkbook.Protect "Z961814r"
On Error GoTo 0
End Sub"""
text = re.sub(r'Sub HideInstallInst\(\).*?End Sub', new_hide2, text, flags=re.DOTALL | re.IGNORECASE)

# Fix HideOpInst
new_hide3 = """Sub HideOpInst()
On Error Resume Next
ThisWorkbook.Unprotect "Z961814r"
ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("AA2").Value = "YES"
Dim sOp As String
sOp = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500)
ThisWorkbook.Worksheets(sOp).Visible = xlSheetVeryHidden
ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
ThisWorkbook.Protect "Z961814r"
On Error GoTo 0
End Sub"""
text = re.sub(r'Sub HideOpInst\(\).*?End Sub', new_hide3, text, flags=re.DOTALL | re.IGNORECASE)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Fixed missing unprotects successfully!")
