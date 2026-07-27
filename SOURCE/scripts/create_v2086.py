import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.085.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.086.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_085"', 'Attribute VB_Name = "Goren_Claude_V2_086"')
content = content.replace("' VERSION: V2.085", "' VERSION: V2.086")
content = content.replace('Private Const APP_VERSION As String = "2.085"', 'Private Const APP_VERSION As String = "2.086"')

old_block = """If Not wsMatach Is Nothing And bordMonth <> "" Then
lastRow = wsMatach.Cells(wsMatach.Rows.Count, 1).End(xlUp).Row
For r = 2 To lastRow
cellVal = Trim$(CStr(wsMatach.Cells(r, 1).Value2))
If cellVal = bordMonth Then
If IsNumeric(wsMatach.Cells(r, rateCol).Value2) Then
rate = CDbl(wsMatach.Cells(r, rateCol).Value2)
If rate > 0 Then
GetCurrencyRate = rate
Exit Function
End If
End If
End If
Next r
End If"""

new_block = """If Not wsMatach Is Nothing And bordMonth <> "" Then
        Dim matchRow As Variant
        matchRow = Application.Match(bordMonth, wsMatach.Columns(1), 0)
        If Not IsError(matchRow) Then
            If IsNumeric(wsMatach.Cells(CLng(matchRow), rateCol).Value2) Then
                rate = CDbl(wsMatach.Cells(CLng(matchRow), rateCol).Value2)
                If rate > 0 Then
                    GetCurrencyRate = rate
                    Exit Function
                End If
            End If
        End If
End If"""

# the formatting in the file has indentations. I should use a regex.
import re
pattern = re.compile(r'If Not wsMatach Is Nothing And bordMonth <> "" Then.*?Next r\s*End If', re.DOTALL)
content = pattern.sub(new_block, content)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.086 created with Match optimized GetCurrencyRate.")

