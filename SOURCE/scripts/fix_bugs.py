import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.113.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix Bug 2: Add DATA and NEW_CLIENTS to hide loop in ApplyCorrectionsAndBuildReports
old_hide = """    For Each wsHide In ThisWorkbook.Worksheets
        If wsHide.Name = MANAGEMENT_SHEET_NAME() Then
            wsHide.Visible = xlSheetVeryHidden
        ElseIf Left$(wsHide.Name, Len(REVIEW_SHEET_NAME())) = REVIEW_SHEET_NAME() Then
            wsHide.Visible = xlSheetVeryHidden
        ElseIf Left$(wsHide.Name, 5) = H_BASE() & "_" Then
            wsHide.Visible = xlSheetVeryHidden
        ElseIf wsHide.Name = H_RESHIMOT() Then
            wsHide.Visible = xlSheetVeryHidden
        End If
    Next wsHide"""

new_hide = """    For Each wsHide In ThisWorkbook.Worksheets
        If wsHide.Name = MANAGEMENT_SHEET_NAME() Then
            wsHide.Visible = xlSheetVeryHidden
        ElseIf wsHide.Name = NEW_CLIENTS_SHEET_NAME() Then
            wsHide.Visible = xlSheetVeryHidden
        ElseIf Left$(wsHide.Name, Len(REVIEW_SHEET_NAME())) = REVIEW_SHEET_NAME() Then
            wsHide.Visible = xlSheetVeryHidden
        ElseIf Left$(wsHide.Name, 5) = H_BASE() & "_" Then
            wsHide.Visible = xlSheetVeryHidden
        ElseIf wsHide.Name = H_RESHIMOT() Then
            wsHide.Visible = xlSheetVeryHidden
        ElseIf Left$(wsHide.Name, 5) = H_DATA() & "_" Then
            wsHide.Visible = xlSheetVeryHidden
        End If
    Next wsHide"""

content = content.replace(old_hide, new_hide)

# Fix Bug 3: Only ask for password when SHOWING sheets
old_toggle = """Public Sub ToggleHiddenSheets()
    Dim pwd As String
    pwd = InputBox(ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & ":", ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492))
    If pwd <> "Z961814r" Then
        MsgBoxU ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1490) & ChrW(1493) & ChrW(1497) & ChrW(1492), vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim anyOtherVisible As Boolean
    anyOtherVisible = False"""

new_toggle = """Public Sub ToggleHiddenSheets()
    Dim ws As Worksheet
    Dim anyOtherVisible As Boolean
    anyOtherVisible = False
    
    ' First check if any sheets are currently visible
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case CONTROL_SHEET_NAME(), MANAGEMENT_SHEET_NAME(), SHEET_COMPANIES(), _
                 SHEET_AGENTS(), SHEET_BRANCH(), SHEET_MAINBRANCH(), _
                 SHEET_TELLERS(), SHEET_MONTHS(), SHEET_SUMMARY()
                ' These are always visible, so ignore them
            Case Else
                If ws.Visible = xlSheetVisible Then
                    anyOtherVisible = True
                    Exit For
                End If
        End Select
    Next ws
    
    ' If we are going to SHOW sheets, require password
    If Not anyOtherVisible Then
        Dim pwd As String
        pwd = InputBox(ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & ":", ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492))
        If pwd <> "Z961814r" Then
            MsgBoxU ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1490) & ChrW(1493) & ChrW(1497) & ChrW(1492), vbCritical
            Exit Sub
        End If
    End If
    
    ' Now perform action
"""

# replace old toggle block (note: I need to replace up to anyOtherVisible = False, but the original code had the loop AFTER anyOtherVisible = False. I should just replace the whole function)
old_func_pattern = r'Public Sub ToggleHiddenSheets\(\).*?End Sub'
new_func_code = """Public Sub ToggleHiddenSheets()
    Dim ws As Worksheet
    Dim anyOtherVisible As Boolean
    anyOtherVisible = False
    
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case CONTROL_SHEET_NAME(), MANAGEMENT_SHEET_NAME(), SHEET_COMPANIES(), _
                 SHEET_AGENTS(), SHEET_BRANCH(), SHEET_MAINBRANCH(), _
                 SHEET_TELLERS(), SHEET_MONTHS(), SHEET_SUMMARY()
                ' These are always visible, so ignore them
            Case Else
                If ws.Visible = xlSheetVisible Then
                    anyOtherVisible = True
                    Exit For
                End If
        End Select
    Next ws
    
    If anyOtherVisible Then
        ' Hiding sheets - no password required
        HideWorkSheets
    Else
        ' Showing sheets - require password
        Dim pwd As String
        pwd = InputBox(ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & ":", ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492))
        If pwd <> "Z961814r" Then
            MsgBoxU ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1490) & ChrW(1493) & ChrW(1497) & ChrW(1492), vbCritical
            Exit Sub
        End If
        ShowHiddenSheets
    End If
End Sub"""

import re
content = re.sub(r'Public Sub ToggleHiddenSheets\(\).*?End Sub', new_func_code, content, flags=re.DOTALL)


# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_113"', 'Attribute VB_Name = "Goren_Claude_V2_114"')
content = content.replace('VERSION: V2.113', 'VERSION: V2.114')
content = content.replace('APP_VERSION As String = "2.113"', 'APP_VERSION As String = "2.114"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.114.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.114 created. Fixed 3 bugs.")
