with open(r'C:\ledugma\DEMO\modDemoReports_V9.41_CopyPaste.txt', 'r', encoding='windows-1255') as f:
    text = f.read()

text = text.replace('wsMain.Rows("1:1").RowHeight = 35', 'wsMain.Rows("1:1").RowHeight = 50')
text = text.replace('wsMain.Range("A1:I1").Merge', 'wsMain.Range("A1:T1").Merge')
text = text.replace('wsMain.Range("A1").Font.Size = 22', 'wsMain.Range("A1").Font.Size = 36\n    wsMain.Range("A1:AZ1").Interior.Color = RGB(220, 240, 220)')
text = text.replace('V9.41', 'V9.42')
text = text.replace('9.41', '9.42')
text = text.replace('CHANGES IN 9.42:', 'CHANGES IN 9.42:\n\'   - Visually improved main title: larger font, centered A-T, green background')

with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.42_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)
