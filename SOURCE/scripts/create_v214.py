import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.213.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.214.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Add the new macro at the end of the file
new_macro = """
' MACRO: HideInstallationInstructions
Public Sub HideInstallationInstructions()
    Dim wsName1 As String, wsName2 As String
    wsName1 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) ' הוראות_התקנה
    wsName2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) ' הוראות התקנה

    On Error Resume Next
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
    ThisWorkbook.Save
    On Error GoTo 0
End Sub
"""

if "Public Sub HideInstallationInstructions" not in content:
    content += new_macro

# Modify HideWorkSheets to handle both space and underscore
old_hide = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' _
        If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then hideIt = False
    End If"""

new_hide = """Dim wsName1 As String, wsName2 As String
    wsName1 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
    wsName2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
    
    If ws.Name = wsName1 Or ws.Name = wsName2 Or ws.Name = "Sheet2" Then
        If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then hideIt = False
    End If"""

content = content.replace(old_hide, new_hide)


# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_213"', 'Attribute VB_Name = "Goren_Claude_V2_214"')
content = content.replace('VERSION: V2.213', 'VERSION: V2.214')
content = content.replace('APP_VERSION As String = "2.213"', 'APP_VERSION As String = "2.214"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.214")
