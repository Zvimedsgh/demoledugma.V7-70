import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.066.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.067.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_067"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.067\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.067"\n')
    elif "' Sort alphabetically" in line and "UpdateClientList" in "".join(lines[max(0, i-50):i]):
        new_lines.append("    ' Sort mathematically (numbers first, then alphabetically)\n")
        new_lines.append("    270 arrAll = dict.keys\n")
        new_lines.append("    Dim isNumA As Boolean, isNumB As Boolean\n")
        new_lines.append("    280 For i = 0 To UBound(arrAll) - 1\n")
        new_lines.append("    290     For j = i + 1 To UBound(arrAll)\n")
        new_lines.append("                isNumA = IsNumeric(arrAll(i))\n")
        new_lines.append("                isNumB = IsNumeric(arrAll(j))\n")
        new_lines.append("                If isNumA And isNumB Then\n")
        new_lines.append("                    If CDbl(arrAll(i)) > CDbl(arrAll(j)) Then\n")
        new_lines.append("                        tmp = arrAll(i): arrAll(i) = arrAll(j): arrAll(j) = tmp\n")
        new_lines.append("                    ElseIf CDbl(arrAll(i)) = CDbl(arrAll(j)) Then\n")
        new_lines.append("                        If UCase(CStr(arrAll(i))) > UCase(CStr(arrAll(j))) Then\n")
        new_lines.append("                            tmp = arrAll(i): arrAll(i) = arrAll(j): arrAll(j) = tmp\n")
        new_lines.append("                        End If\n")
        new_lines.append("                    End If\n")
        new_lines.append("                ElseIf isNumA And Not isNumB Then\n")
        new_lines.append("                    ' Keep A before B (numbers before strings)\n")
        new_lines.append("                ElseIf Not isNumA And isNumB Then\n")
        new_lines.append("                    ' Swap A and B (numbers before strings)\n")
        new_lines.append("                    tmp = arrAll(i): arrAll(i) = arrAll(j): arrAll(j) = tmp\n")
        new_lines.append("                Else\n")
        new_lines.append("                    ' Both strings\n")
        new_lines.append("                    If UCase(CStr(arrAll(i))) > UCase(CStr(arrAll(j))) Then\n")
        new_lines.append("                        tmp = arrAll(i): arrAll(i) = arrAll(j): arrAll(j) = tmp\n")
        new_lines.append("                    End If\n")
        new_lines.append("                End If\n")
        new_lines.append("    330     Next j\n")
        new_lines.append("    340 Next i\n")
        skip = True
    elif skip and "340 Next i" in line:
        skip = False
    elif not skip:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.067")

