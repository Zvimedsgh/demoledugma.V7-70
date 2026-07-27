import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.088.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.089.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_088"', 'Attribute VB_Name = "Goren_Claude_V2_089"')
content = content.replace("' VERSION: V2.088", "' VERSION: V2.089")
content = content.replace('Private Const APP_VERSION As String = "2.088"', 'Private Const APP_VERSION As String = "2.089"')

# Fix FetchBOIMonthlyAvg
pattern_fetch = re.compile(r'Private Function FetchBOIMonthlyAvg.*?Dim xmlHttp As Object', re.DOTALL)
new_fetch = """Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
Static apiFailed As Boolean
Static apiCallsCount As Long
If apiFailed Or apiCallsCount > 40 Then
    FetchBOIMonthlyAvg = 0
    Exit Function
End If
apiCallsCount = apiCallsCount + 1
On Error GoTo FAIL
Dim xmlHttp As Object"""
content = pattern_fetch.sub(new_fetch, content)

# Fix FAIL: block
pattern_fail = re.compile(r'FAIL:\s*\r?\n\s*FetchBOIMonthlyAvg = 0\s*\r?\n\s*End Function', re.DOTALL)
new_fail = """FAIL:
apiFailed = True
FetchBOIMonthlyAvg = 0
End Function"""
content = pattern_fail.sub(new_fail, content)

# Fix GetBordMonth
pattern_bord = re.compile(r'GetBordMonth = Format\$\(dt, "yyyy-mm"\)\s*\r?\n\s*Exit Function', re.DOTALL)
new_bord = """Dim yr As Long
yr = Year(dt)
If yr >= 2000 And yr <= 2050 Then
    GetBordMonth = Format$(dt, "yyyy-mm")
Else
    GetBordMonth = ""
End If
Exit Function"""
content = pattern_bord.sub(new_bord, content)

# Fix BuildMatachCache array logic (which also failed earlier)
pattern_cache = re.compile(r"' Create dictionary of unique bordero months from source data\s*\r?\n\s*Set dictMonths = CreateObject\(\"Scripting.Dictionary\"\)\s*\r?\n\s*For r = 2 To lastDataRow\s*\r?\n\s*bordDate = wsSource.Cells\(r, dateCol\).Value2\s*\r?\n\s*monthKey = GetBordMonth\(bordDate\)\s*\r?\n\s*If monthKey <> \"\" Then\s*\r?\n\s*If Not dictMonths.Exists\(monthKey\) Then dictMonths.Add monthKey, True\s*\r?\n\s*End If\s*\r?\n\s*Next r", re.DOTALL)
new_cache = """' Create dictionary of unique bordero months from source data
Set dictMonths = CreateObject("Scripting.Dictionary")
Dim dateArr As Variant
On Error Resume Next
dateArr = wsSource.Range(wsSource.Cells(1, dateCol), wsSource.Cells(lastDataRow, dateCol)).Value2
On Error GoTo 0
If Not IsEmpty(dateArr) Then
    For r = 2 To lastDataRow
    bordDate = dateArr(r, 1)
    monthKey = GetBordMonth(bordDate)
    If monthKey <> "" Then
    If Not dictMonths.Exists(monthKey) Then dictMonths.Add monthKey, True
    End If
    Next r
End If"""
content = pattern_cache.sub(new_cache, content)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.089 created with proper regex replacements.")
