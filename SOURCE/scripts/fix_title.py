import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = """    ' 4. Main Title
    Dim shpTitle As Shape
    On Error Resume Next
    wsMain.Shapes("txtMainTitle").Delete
    On Error GoTo ERR_HANDLER
    Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("E1").Left, wsMain.Range("E1").Top, 600, 40)
    shpTitle.Name = "txtMainTitle"
    shpTitle.TextFrame2.TextRange.Font.Size = 24
    shpTitle.TextFrame2.TextRange.Font.Bold = msoTrue
    shpTitle.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpTitle.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(180, 50, 50)
    shpTitle.TextFrame2.TextRange.Text = GetActiveAgencyName()
    shpTitle.Line.Visible = msoFalse
    shpTitle.Fill.Visible = msoFalse"""

new_str = """    ' 4. Main Title (Centered in row 1)
    wsMain.Range("C1:K1").Merge
    wsMain.Range("C1").Value = GetActiveAgencyName()
    wsMain.Range("C1").Font.Size = 24
    wsMain.Range("C1").Font.Bold = True
    wsMain.Range("C1").HorizontalAlignment = xlCenter
    wsMain.Range("C1").VerticalAlignment = xlCenter
    wsMain.Range("C1").Font.Color = RGB(180, 50, 50)"""

if old_str in content:
    content = content.replace(old_str, new_str)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Fixed Main Title")
else:
    print("Could not find Main Title string")
