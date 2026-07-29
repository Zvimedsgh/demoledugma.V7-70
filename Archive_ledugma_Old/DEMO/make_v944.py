with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.43_CopyPaste.txt', 'r', encoding='windows-1255') as f:
    text = f.read()

# At the end of A00_SetupMainSheet, right before End Sub, add the font size 36 again
text = text.replace('3341 \n3342 End Sub', '3341     wsMain.Range("A1").Font.Size = 36\n3342 End Sub')

text = text.replace('V9.43', 'V9.44')
text = text.replace('9.43', '9.44')

with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.44_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)
