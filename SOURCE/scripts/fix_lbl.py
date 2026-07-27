import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = """Dim shpTitle As Shape
Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("E1").Left, wsMain.Range("E1").Top, 600, 40)
shpTitle.Name = "lblAgencyTitle"
shpTitle.TextFrame2.TextRange.Text = GetActiveAgencyName()
shpTitle.TextFrame2.TextRange.Font.Size = 28
shpTitle.TextFrame2.TextRange.Font.Bold = msoTrue
shpTitle.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(200, 0, 0)
shpTitle.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter
shpTitle.Line.Visible = msoFalse
shpTitle.Fill.Visible = msoFalse"""

new_str = """' Write Agency Name to cell and center across selection
wsMain.Range("C1:K1").Merge
wsMain.Range("C1").Value = GetActiveAgencyName()
wsMain.Range("C1").Font.Size = 28
wsMain.Range("C1").Font.Bold = True
wsMain.Range("C1").Font.Color = RGB(200, 0, 0)
wsMain.Range("C1").HorizontalAlignment = xlCenter
wsMain.Range("C1").VerticalAlignment = xlCenter"""

if old_str in content:
    content = content.replace(old_str, new_str)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Fixed title shape to cell")
else:
    print("Could not find title string")
