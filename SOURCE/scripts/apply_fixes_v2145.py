import sys, re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_update = """    60 shtNames = Array("DATA_" & (Year(Date) - 1), "DATA_" & Year(Date))
    
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
    240 Next iIdx"""

new_update = """    ' Search through all base sheets (H_BASE)
    For Each wsSrc In ThisWorkbook.Worksheets
        If InStr(1, wsSrc.Name, H_BASE(), vbTextCompare) > 0 Then
            lastRow = 1
            On Error Resume Next
            lastRow = wsSrc.Cells(wsSrc.Rows.Count, 6).End(xlUp).Row
            On Error GoTo ERR_HANDLER
            
            If lastRow > 1 Then
                Dim clientData As Variant
                clientData = wsSrc.Range(wsSrc.Cells(2, 6), wsSrc.Cells(lastRow, 6)).Value2
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
            End If
        End If
    Next wsSrc"""

if target_update in content:
    content = content.replace(target_update, new_update)
    print("Fixed UpdateClientList.")
else:
    print("Could not find UpdateClientList target.")

# Update the title color for E3/Search Screen as requested
# We need to make the title red again if they wanted it. But they said "like the title that was in red and is now black". Let's change the color of the title in SearchClientName.
target_title = """wsSearch.Cells(1, 1).Font.Size = 13"""
new_title = """wsSearch.Cells(1, 1).Font.Size = 13\n    wsSearch.Cells(1, 1).Font.Color = RGB(255, 0, 0) ' Make it red again!"""

if target_title in content:
    content = content.replace(target_title, new_title)
    print("Fixed search title color.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_144"', 'Attribute VB_Name = "Goren_Claude_V2_145"')
content = content.replace('VERSION: V2.144', 'VERSION: V2.145')
content = content.replace('APP_VERSION As String = "2.144"', 'APP_VERSION As String = "2.145"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.145.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.145 created.")
