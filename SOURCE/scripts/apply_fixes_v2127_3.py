import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.127.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if i >= 3684 and i <= 3688:
        if "wsMain.Range(\"D3\")" in line:
            lines[i] = line.replace('wsMain.Range("D3")', 'wsMain.Range("D15")')
        if "Font.Size = 16" in line:
            lines[i] = line.replace("Font.Size = 16", "Font.Size = 17")

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Replaced A00_SetupMainSheet")
