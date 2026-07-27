import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.192.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Fix Client Search to always use InStr (Contains)
search_old = """        Dim bWildcard As Boolean
        bWildcard = (InStr(1, searchText, "*") > 0 Or InStr(1, searchText, "?") > 0)
        
        For r = 2 To lastRow
            cName = Trim$(CStr(ws.Cells(r, 6).Value2))
            If cName <> "" Then
                Dim bMatch As Boolean
                bMatch = False
                If bWildcard Then
                    If UCase(cName) Like UCase(searchText) Then bMatch = True
                Else
                    ' Whole word match
                    Dim paddedName As String, paddedSearch As String
                    paddedName = " " & Replace(Replace(UCase(cName), "-", " "), "_", " ") & " "
                    paddedSearch = " " & Replace(Replace(UCase(searchText), "-", " "), "_", " ") & " "
                    If InStr(1, paddedName, paddedSearch, vbTextCompare) > 0 Then bMatch = True
                End If
                If bMatch Then"""

search_new = """        ' Always use partial match (Contains)
        For r = 2 To lastRow
            cName = Trim$(CStr(ws.Cells(r, 6).Value2))
            If cName <> "" Then
                Dim bMatch As Boolean
                bMatch = False
                Dim cleanSearch As String
                cleanSearch = Replace(Replace(searchText, "*", ""), "?", "")
                If InStr(1, UCase(cName), UCase(cleanSearch), vbTextCompare) > 0 Then bMatch = True
                If bMatch Then"""

content = content.replace(search_old, search_new)


# 2. Update Search screen instructions
search_inst_old = """    wsSearch.Cells(5, 5).Value = ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & ":"
    wsSearch.Cells(6, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1512) & ChrW(1511) & " 77"
    wsSearch.Cells(7, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " *77*: " & ChrW(1497) & _
                                 ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & _
                                 ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77"
    wsSearch.Cells(8, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77*: " & ChrW(1497) & _
                                 ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & _
                                 ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77\""""

search_inst_new = """    wsSearch.Cells(5, 5).Value = ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & ":"
    wsSearch.Cells(6, 5).Value = "- " & ChrW(1492) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & " " & ChrW(1495) & ChrW(1499) & ChrW(1501) & ":"
    wsSearch.Cells(7, 5).Value = ChrW(1488) & ChrW(1497) & ChrW(1503) & " " & ChrW(1510) & ChrW(1493) & ChrW(1512) & ChrW(1498) & " " & ChrW(1489) & ChrW(1499) & ChrW(1493) & ChrW(1499) & ChrW(1489) & ChrW(1497) & ChrW(1493) & ChrW(1514)
    wsSearch.Cells(8, 5).Value = ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77 " & ChrW(1493) & ChrW(1492) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1514מצא) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1513) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77\""""
search_inst_new = search_inst_new.replace("ChrW(1514מצא)", "ChrW(1514) & ChrW(1502) & ChrW(1510) & ChrW(1488)")
content = content.replace(search_inst_old, search_inst_new)


# 3. Fix CheckUserPermissions visibility for Operation Instructions
check_perm_old = """        If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
            If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then hideIt2 = False
        End If"""

check_perm_new = """        If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
            If UCase$(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("AA2").Value) <> "YES" Then 
                hideIt2 = False
                If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
            End If
        End If"""
content = content.replace(check_perm_old, check_perm_new)


# 4. Fix HideWorkSheets visibility for Operation Instructions
hide_old = """    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
        If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA2").Value) <> "YES" Then hideIt = False
    End If"""

hide_new = """    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
        If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA2").Value) <> "YES" Then 
            hideIt = False
            If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
        End If
    End If"""
content = content.replace(hide_old, hide_new)

# 5. Fix jumping on open
jump_old = """    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then
        On Error Resume Next
        ThisWorkbook.Worksheets(sOp2).Activate
        On Error GoTo 0
    End If"""

jump_new = """    If UCase$(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("AA2").Value) <> "YES" Then
        On Error Resume Next
        ThisWorkbook.Worksheets(sOp2).Activate
        On Error GoTo 0
    End If"""
content = content.replace(jump_old, jump_new)


content = content.replace('Attribute VB_Name = "Goren_Claude_V2_192"', 'Attribute VB_Name = "Goren_Claude_V2_193"')
content = content.replace('VERSION: V2.192', 'VERSION: V2.193')
content = content.replace('APP_VERSION As String = "2.192"', 'APP_VERSION As String = "2.193"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.193")
