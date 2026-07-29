with open(r'C:\ledugma\DEMO\V9.48_CopyPaste.txt', 'r', encoding='windows-1255') as f:
    text = f.read()

# Clear A16
target = '        wsMain.Range("B17:C17").Clear\n        wsMain.Range("A17").Value ='
replacement = '        wsMain.Range("B17:C17").Clear\n        wsMain.Range("A16").ClearContents\n        wsMain.Range("A17").Value ='
text = text.replace(target, replacement)

text = text.replace('V9.48', 'V9.49')
text = text.replace('9.48', '9.49')

# Add changelog
text = text.replace('CHANGES IN 9.49:', 'CHANGES IN 9.49:\n\'   - Cleared A16 to prevent duplicate credit text.')

with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.49_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)

with open(r'C:\ledugma\DEMO\V9.49_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)
