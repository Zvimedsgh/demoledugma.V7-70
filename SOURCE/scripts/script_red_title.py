import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''    Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F11").Left + (wsMain.Range("F11").Width - 60) / 2, wsMain.Range("F11").Top + 5, 60, 25)'''
new_code = '''    ' Add Large Red Title for Agency Name
    On Error Resume Next
    wsMain.Shapes("lblAgencyTitle").Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    Dim shpTitle As Shape
    Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("A2").Left, wsMain.Range("A2").Top, 600, 40)
    shpTitle.Name = "lblAgencyTitle"
    shpTitle.TextFrame2.TextRange.Text = GetActiveAgencyName()
    shpTitle.TextFrame2.TextRange.Font.Size = 28
    shpTitle.TextFrame2.TextRange.Font.Bold = msoTrue
    shpTitle.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(200, 0, 0)
    shpTitle.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter
    shpTitle.Line.Visible = msoFalse
    shpTitle.Fill.Visible = msoFalse

    Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F11").Left + (wsMain.Range("F11").Width - 60) / 2, wsMain.Range("F11").Top + 5, 60, 25)'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
