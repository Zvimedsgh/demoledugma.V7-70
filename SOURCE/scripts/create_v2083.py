import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.082.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.083.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_083"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.083\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.083"\n')
    elif "Public Sub UpdateClientList()" in line:
        new_lines.append(line)
        skip = True
    elif skip and "End Sub" in line:
        skip = False
        new_lines.append(line)
    elif skip:
        pass
    else:
        new_lines.append(line)

update_client_list = """
    Dim wsClients As Worksheet
    Dim wsSrc As Worksheet
    Dim shtNames As Variant
    Dim r As Long, iIdx As Integer
    Dim lastRow As Long
    Dim dict As Object
    Dim arrAll As Variant
    
    10 On Error Resume Next
    20 Set wsClients = ThisWorkbook.Worksheets(H_SET_CLIENTS())
    30 If wsClients Is Nothing Then Exit Sub
    40 On Error GoTo ERR_HANDLER
    
    50 Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    60 shtNames = Array("DATA_" & (Year(Date) - 1), "DATA_" & Year(Date))
    
    70 For iIdx = 0 To UBound(shtNames)
        80 On Error Resume Next
        90 Set wsSrc = ThisWorkbook.Worksheets(shtNames(iIdx))
        100 On Error GoTo ERR_HANDLER
        
        110 If Not wsSrc Is Nothing Then
            120 lastRow = 1
            130 On Error Resume Next
            140 lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            150 On Error GoTo ERR_HANDLER
            
            160 If lastRow > 1 Then
                Dim clientData As Variant
                170 clientData = wsSrc.Range(wsSrc.Cells(2, RAW_CLIENTNAME), wsSrc.Cells(lastRow, RAW_CLIENTNAME)).Value2
                Dim rowIdx As Long
                For rowIdx = 1 To UBound(clientData, 1)
                    Dim cName As String
                    cName = Trim$(CStr(clientData(rowIdx, 1)))
                    If Len(cName) > 0 Then
                        If Not dict.Exists(cName) Then
                            dict.Add cName, 1
                        End If
                    End If
                Next rowIdx
            220 End If
        230 End If
    240 Next iIdx
    
    250 If dict.Count = 0 Then Exit Sub
    
    270 arrAll = dict.keys
    
    ' Write to column A on H_SET_CLIENTS
    Dim startRow As Long
    350 startRow = 1
    360 wsClients.Cells(startRow, 1).Value = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    370 wsClients.Cells(startRow, 1).Font.Bold = True
    380 wsClients.Cells(startRow, 1).Font.Size = 12
    
    wsClients.Range("A2:A" & wsClients.Rows.Count).ClearContents
    
    If UBound(arrAll) >= 0 Then
        Dim outArr() As Variant
        ReDim outArr(0 To UBound(arrAll), 1 To 1)
        For r = 0 To UBound(arrAll)
            outArr(r, 1) = arrAll(r)
        Next r
        wsClients.Range(wsClients.Cells(startRow + 1, 1), wsClients.Cells(startRow + 1 + UBound(arrAll), 1)).Value2 = outArr
        
        wsClients.Range(wsClients.Cells(startRow + 1, 1), wsClients.Cells(startRow + 1 + UBound(arrAll), 1)).Sort _
            Key1:=wsClients.Cells(startRow + 1, 1), Order1:=xlAscending, Header:=xlNo
    End If
    
    420 wsClients.Cells(startRow + UBound(arrAll) + 2, 1).Value = EOD_MARKER
    
    450 On Error Resume Next
    460 ThisWorkbook.Names("lst_clients").Delete
    470 On Error GoTo 0
    480 If UBound(arrAll) >= 0 Then
    490     ThisWorkbook.Names.Add "lst_clients", wsClients.Range(wsClients.Cells(startRow + 1, 1), wsClients.Cells(startRow + 1 + UBound(arrAll), 1))
    500 End If
    Exit Sub
    
ERR_HANDLER:
    ' Just ignore on error
"""

idx = new_lines.index("Public Sub UpdateClientList()\n")
new_lines.insert(idx + 1, update_client_list)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.083 created with fast UpdateClientList.")

