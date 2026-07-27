import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "ThisWorkbook.Names.Add \"rngFILES_FOLDER\", wsMgmt.Range(\"B176\")" in line:
        new_lines.append(line.replace("wsMgmt", "ThisWorkbook.Worksheets(H_SET_PARAMS())"))
    elif "ThisWorkbook.Names.Add \"rngREPORTS_FOLDER\", wsMgmt.Range(\"B177\")" in line:
        new_lines.append(line.replace("wsMgmt", "ThisWorkbook.Worksheets(H_SET_PARAMS())"))
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Fixed rngFILES_FOLDER and rngREPORTS_FOLDER to point to H_SET_PARAMS")
