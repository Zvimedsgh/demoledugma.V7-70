import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.018.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''        Dim ftVal As String
        ftVal = Trim$(CStr(wsCheck.Range("rngFilterType").Value2))
        If ftVal <> "" And ftVal <> selText Then
            Dim fvVal As String
            fvVal = Trim$(CStr(wsCheck.Range("rngFilterValue").Value2))
            If fvVal = "" Or fvVal = selText Then
                MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1512) & " " & ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1505) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1503) & " " & ChrW(1489) & ChrW(1514) & ChrW(1488) & " G10", vbExclamation
                Exit Sub
            End If
        End If
        ' --- End validation ---'''

new_code = '''        Dim filterCol As Long
        Dim filterValue As String
        filterCol = 0
        filterValue = ""

        Dim ftVal As String
        ftVal = Trim$(CStr(wsCheck.Range("rngFilterType").Value2))
        If ftVal <> "" And ftVal <> selText Then
            Dim fvVal As String
            fvVal = Trim$(CStr(wsCheck.Range("rngFilterValue").Value2))
            If fvVal = "" Or fvVal = selText Then
                MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1512) & " " & ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1505) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1503) & " " & ChrW(1489) & ChrW(1514) & ChrW(1488) & " G9", vbExclamation
                Exit Sub
            End If
            
            filterValue = fvVal
            If ftVal = H_COMPANY() Then
                filterCol = BASE_COL_COMPANY
            ElseIf ftVal = H_TELLER() Then
                filterCol = BASE_COL_TELLER
            ElseIf ftVal = H_AGENT() Then
                filterCol = BASE_COL_AGENTNAME
            ElseIf ftVal = H_BRANCH() Then
                filterCol = BASE_COL_BRANCHNAME
            ElseIf ftVal = H_BRANCH() & ChrW(32) & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494) Then
                filterCol = BASE_COL_MAINBRANCH
            End If
        End If
        ' --- End validation ---'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
