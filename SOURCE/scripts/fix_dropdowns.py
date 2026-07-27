import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.043_20260702_1545.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_period = """90      If InStr(1, periodType, ChrW(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) > 0 Then
100         listName = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504) & ChrW(1492) & "," & ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1492)
' "riv'oni" = quarterly
110     ElseIf InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
120         listName = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497) & "," & ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497)
' "chodshi" = monthly
130     ElseIf InStr(1, periodType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
140         listName = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512) & "," & ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512) & "," & ChrW(1502) & ChrW(1512) & ChrW(1509) & "," & ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500) & "," & ChrW(1502) & ChrW(1488) & ChrW(1497) & "," & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497) & "," & ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497) & "," & ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496) & "," & ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512) & "," & ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512) & "," & ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512) & "," & ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
' "shnatit" = yearly -> no second dropdown needed, jump to DateType
150     ElseIf InStr(1, periodType, ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497), vbTextCompare) > 0 Then"""

new_period = """90      If InStr(1, periodType, ChrW(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) > 0 Then
100         listName = "=lst_half_year"
' "riv'oni" = quarterly
110     ElseIf InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
120         listName = "=lst_quarter"
' "chodshi" = monthly
130     ElseIf InStr(1, periodType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
140         listName = "=lst_month"
' "shnatit" = yearly -> no second dropdown needed, jump to DateType
150     ElseIf InStr(1, periodType, ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497), vbTextCompare) > 0 Then"""

if old_period in content:
    content = content.replace(old_period, new_period)
    print("Replaced UpdatePeriodDropdown logic")
else:
    print("FAILED to find old_period")

old_filter = """' Build comma-separated list
320     valList = ""
330     For i = 2 To lastR
340         If Trim$(CStr(wsLists.Cells(i, targetCol).Value2)) <> "" Then
350             If valList <> "" Then valList = valList & ","
360             valList = valList & Trim$(CStr(wsLists.Cells(i, targetCol).Value2))
370         End If
380     Next i

390     If valList = "" Then GoTo CLEAN_EXIT

' Add validation list to G10"""

new_filter = """' Create a Named Range for the validation to avoid 255 char limit
320     ThisWorkbook.Names.Add "lst_temp_filter", wsLists.Range(wsLists.Cells(2, targetCol), wsLists.Cells(lastR, targetCol))
330     valList = "=lst_temp_filter"

390     If valList = "" Then GoTo CLEAN_EXIT

' Add validation list to G10"""

if old_filter in content:
    content = content.replace(old_filter, new_filter)
    print("Replaced UpdateFilterValueDropdown logic")
else:
    print("FAILED to find old_filter")

# Upgrade version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_043"', 'Attribute VB_Name = "Goren_Claude_V2_044"')
content = content.replace('Private Const APP_VERSION As String = "2.043"', 'Private Const APP_VERSION As String = "2.044"')
content = content.replace('VERSION: V2.043', 'VERSION: V2.044')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.044.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.044")
