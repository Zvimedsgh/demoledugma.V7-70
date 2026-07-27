import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.084.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.085.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_build_review = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_085"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.085\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.085"\n')
    elif "Public Sub BuildReview()" in line:
        in_build_review = True
        new_lines.append(line)
        new_lines.append("    Dim outArr() As Variant\n")
        new_lines.append("    Dim outIdx As Long\n")
    elif in_build_review and "End Sub" in line:
        in_build_review = False
        new_lines.append(line)
    elif in_build_review and "700     outRow = 3" in line:
        new_lines.append(line)
        new_lines.append("    outIdx = 0\n")
        new_lines.append("    If lastRow > 1 And cnt > 0 Then ReDim outArr(1 To lastRow * cnt, 1 To cnt + 2)\n")
    elif in_build_review and "wsRev.Cells(outRow, 1).Value = r" in line:
        new_lines.append("                outIdx = outIdx + 1\n")
        new_lines.append("                outArr(outIdx, 1) = r\n")
    elif in_build_review and "wsRev.Cells(outRow, j + 1).Value = srcData(r, cols(j))" in line:
        new_lines.append("                    outArr(outIdx, j + 1) = srcData(r, cols(j))\n")
    elif in_build_review and "wsRev.Cells(outRow, cnt + 2).Value = singleText" in line:
        new_lines.append("                outArr(outIdx, cnt + 2) = singleText\n")
    elif in_build_review and "NextRow:" in line:
        new_lines.append(line)
    elif in_build_review and "990     Next r" in line:
        new_lines.append(line)
        new_lines.append("    ' Dump all errors to wsRev\n")
        new_lines.append("    If outIdx > 0 Then\n")
        new_lines.append("        wsRev.Range(wsRev.Cells(3, 1), wsRev.Cells(3 + outIdx - 1, cnt + 2)).Value2 = outArr\n")
        new_lines.append("    End If\n")
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.085 created with Array processing for BuildReview output.")

