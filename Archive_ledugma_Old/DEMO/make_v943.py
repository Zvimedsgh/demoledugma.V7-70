with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.42_CopyPaste.txt', 'r', encoding='windows-1255') as f:
    text = f.read()

text = text.replace('wsMain.Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove', 'If wsMain.Range("A1").Value <> wsMgmt.Range("K46").Value Then\n        wsMain.Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove\n    End If')

text = text.replace('V9.42', 'V9.43')
text = text.replace('9.42', '9.43')

with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.43_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)
