import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.083.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.084.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_build_review = False
in_is_ignorable = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_084"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.084\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.084"\n')
    elif "Public Sub BuildReview()" in line:
        in_build_review = True
        new_lines.append(line)
        new_lines.append("    Dim srcData As Variant\n")
    elif in_build_review and "End Sub" in line:
        in_build_review = False
        new_lines.append(line)
    elif in_build_review and "Set wsSrc = " in line and not "wsSrc Is Nothing" in line:
        new_lines.append(line)
        # We need to make sure we load srcData after wsSrc is set, but wait, wsSrc is set inside If isDemoMode... Else... End If
        # Better to put the srcData load right before "lastRow = " or after it
    elif in_build_review and "580     lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row" in line:
        new_lines.append(line)
        new_lines.append("    On Error Resume Next\n")
        new_lines.append("    srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, 50)).Value2\n")
        new_lines.append("    On Error GoTo ERR_HANDLER\n")
    elif in_build_review and "IsIgnorableRow(wsSrc, r, keys, cols, cnt, dictFieldCol)" in line:
        new_lines.append(line.replace("IsIgnorableRow(wsSrc, r, keys, cols, cnt, dictFieldCol)", "IsIgnorableRowArr(srcData, r, keys, cols, cnt, dictFieldCol)"))
    elif in_build_review and "wsSrc.Cells(" in line and ".Value2" in line and not "wsSrc.Cells(wsSrc.Rows.Count" in line:
        # replace wsSrc.Cells(r, XYZ).Value2 with srcData(r, XYZ)
        l = line.replace("wsSrc.Cells(r, cols(i)).Value2", "srcData(r, cols(i))")
        l = l.replace("wsSrc.Cells(r, cols(j)).Value2", "srcData(r, cols(j))")
        l = l.replace("wsSrc.Cells(r, dictFieldCol(KEY_PREMIUM)).Value2", "srcData(r, dictFieldCol(KEY_PREMIUM))")
        l = l.replace("wsSrc.Cells(r, RAW_CURRENCY).Value2", "srcData(r, RAW_CURRENCY)")
        l = l.replace("wsSrc.Cells(r, dateColBR).Value2", "srcData(r, dateColBR)")
        l = l.replace("wsSrc.Cells(r, RAW_BRANCHNAME).Value2", "srcData(r, RAW_BRANCHNAME)")
        l = l.replace("wsSrc.Cells(r, RAW_COMPANY).Value2", "srcData(r, RAW_COMPANY)")
        l = l.replace("wsSrc.Cells(r, RAW_TELLERNAME).Value2", "srcData(r, RAW_TELLERNAME)")
        l = l.replace("wsSrc.Cells(r, RAW_AGENTNAME).Value2", "srcData(r, RAW_AGENTNAME)")
        new_lines.append(l)
    elif "Private Function IsIgnorableRow(" in line:
        # copy the old function just in case
        new_lines.append(line)
    elif "End Function" in line and len(new_lines) > 0 and "IsIgnorableRow =" in new_lines[-1]:
        new_lines.append(line)
        # append the new array-based function
        arr_func = """
Private Function IsIgnorableRowArr(srcData As Variant, ByVal r As Long, ByRef keys() As String, ByRef cols() As Long, ByVal cnt As Long, ByVal dictFieldCol As Object) As Boolean
    Dim allBlank As Boolean
    Dim i As Long
    allBlank = True
    For i = 1 To cnt
        If Not IsBlankValue(srcData(r, cols(i))) Then
            allBlank = False
            Exit For
        End If
    Next i
    IsIgnorableRowArr = allBlank
End Function
"""
        new_lines.append(arr_func)
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.084 created with Array processing for BuildReview.")

