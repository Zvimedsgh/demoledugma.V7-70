with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.45_CopyPaste.txt', 'r', encoding='windows-1255') as f:
    text = f.read()

# Fix the bug where font size was reset
text = text.replace('wsLoop.Cells.Font.Size = 14\nNext wsLoop', 'wsLoop.Cells.Font.Size = 14\nNext wsLoop\n\n\' Restore A1 huge title font size after loop\nwsMain.Range("A1").Font.Size = 36')

text = text.replace('V9.45', 'V9.46')
text = text.replace('9.45', '9.46')

with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.46_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)
