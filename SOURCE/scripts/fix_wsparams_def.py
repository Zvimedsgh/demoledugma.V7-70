import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_20260702_1430.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "' ---- Set Error_Email parameter if not exists ----" in line and i > 3840:
        new_lines.append("Dim wsParams As Worksheet\n")
        new_lines.append("On Error Resume Next\n")
        new_lines.append("Set wsParams = ThisWorkbook.Worksheets(H_SET_PARAMS())\n")
        new_lines.append("On Error GoTo ERR_HANDLER\n")
        new_lines.append("If Not wsParams Is Nothing Then\n")
        new_lines.append(line)
    elif "2130  End If" in line and i > 3880:
        new_lines.append(line)
        new_lines.append("End If ' End If Not wsParams Is Nothing\n")
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.038 with wsParams definition fixed")
