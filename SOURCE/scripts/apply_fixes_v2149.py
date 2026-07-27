import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.148.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = """    ' Write to G12 on home sheet
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    wsMain.Range("rngClientName").Value = selectedName
        On Error Resume Next
        wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = selectedName
        On Error GoTo ERR_HANDLER
    
    ' Delete search sheet and go back to home"""

new_code = """    ' Write to G12 on home sheet
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    wsMain.Range("rngClientName").Value = selectedName
        On Error Resume Next
        wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = selectedName
        On Error GoTo 0
    
    ' Delete search sheet and go back to home"""

if target in content:
    content = content.replace(target, new_code)
    print("Fixed ERR_HANDLER syntax error.")
else:
    print("Could not find the target string.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_148"', 'Attribute VB_Name = "Goren_Claude_V2_149"')
content = content.replace('VERSION: V2.148', 'VERSION: V2.149')
content = content.replace('APP_VERSION As String = "2.148"', 'APP_VERSION As String = "2.149"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.149.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.149 created.")
