import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.086.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.087.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_086"', 'Attribute VB_Name = "Goren_Claude_V2_087"')
content = content.replace("' VERSION: V2.086", "' VERSION: V2.087")
content = content.replace('Private Const APP_VERSION As String = "2.086"', 'Private Const APP_VERSION As String = "2.087"')

fetch_old = """Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
On Error GoTo FAIL
Dim xmlHttp As Object"""

fetch_new = """Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
Static apiFailed As Boolean
If apiFailed Then
    FetchBOIMonthlyAvg = 0
    Exit Function
End If
On Error GoTo FAIL
Dim xmlHttp As Object"""

content = content.replace(fetch_old, fetch_new)

fail_old = """FAIL:
FetchBOIMonthlyAvg = 0
End Function"""

fail_new = """FAIL:
apiFailed = True
FetchBOIMonthlyAvg = 0
End Function"""

content = content.replace(fail_old, fail_new)

cache_old = """' Create dictionary of unique bordero months from source data
Set dictMonths = CreateObject("Scripting.Dictionary")
For r = 2 To lastDataRow
bordDate = wsSource.Cells(r, dateCol).Value2
monthKey = GetBordMonth(bordDate)
If monthKey <> "" Then
If Not dictMonths.Exists(monthKey) Then dictMonths.Add monthKey, True
End If
Next r"""

cache_new = """' Create dictionary of unique bordero months from source data
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

content = content.replace(cache_old, cache_new)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.087 created with API timeout fix and cache array fix.")

