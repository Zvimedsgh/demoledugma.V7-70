with open(r'C:\ledugma\DEMO\V9.47_CopyPaste.txt', 'r', encoding='windows-1255') as f:
    lines = f.readlines()

for i in range(len(lines)):
    if 'With wsProgress.Range("F17")' in lines[i] and i > 2000:
        lines[i] = lines[i].replace('wsProgress', 'wsMain')

text = "".join(lines)
text = text.replace('V9.47', 'V9.48')
text = text.replace('9.47', '9.48')

# Add changelog for V9.48
text = text.replace('CHANGES IN 9.48:', 'CHANGES IN 9.48:\n\'   - Fixed wsProgress Object Required bug in A30_ClientReport error handler.')

with open(r'C:\Users\Zvi\OneDrive - Zvi Goren צבי גורן\Desktop\V9.48_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)

with open(r'C:\ledugma\DEMO\V9.48_CopyPaste.txt', 'w', encoding='windows-1255') as f:
    f.write(text)
