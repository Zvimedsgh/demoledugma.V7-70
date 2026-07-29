with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

new_code = """127     For Each wsTmp3 In ThisWorkbook.Worksheets
128         If wsTmp3.Visible <> xlSheetVisible Then
129             hiddenCount3 = hiddenCount3 + 1
130             hiddenSheets3(hiddenCount3) = wsTmp3.Name
                ThisWorkbook.Unprotect "Z961814r"
131             wsTmp3.Visible = xlSheetVisible
132         End If
133     Next wsTmp3"""

text = re.sub(r'127\s+For Each wsTmp3.*?133\s+Next wsTmp3', new_code, text, flags=re.DOTALL | re.IGNORECASE)

# Do the same for SaveReportsToFolder
new_code2 = """210     For Each wsTemp In ThisWorkbook.Worksheets
211         If wsTemp.Visible <> xlSheetVisible Then
212             hiddenCount = hiddenCount + 1
213             hiddenSheets(hiddenCount) = wsTemp.Name
                ThisWorkbook.Unprotect "Z961814r"
214             wsTemp.Visible = xlSheetVisible
215         End If
216     Next wsTemp"""

text = re.sub(r'210\s+For Each wsTemp.*?216\s+Next wsTemp', new_code2, text, flags=re.DOTALL | re.IGNORECASE)


with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Added explicit Unprotect inside the loops!")
