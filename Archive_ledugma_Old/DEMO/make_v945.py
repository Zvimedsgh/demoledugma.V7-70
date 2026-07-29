with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.44_CopyPaste.txt', 'r', encoding='windows-1255') as f:
    text = f.read()

# Fix the bug where wsMain is used before being Set
text = text.replace('With wsMain.Range("F17")', 'With wsProgress.Range("F17")')

text = text.replace('V9.44', 'V9.45')
text = text.replace('9.44', '9.45')

with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.45_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)
