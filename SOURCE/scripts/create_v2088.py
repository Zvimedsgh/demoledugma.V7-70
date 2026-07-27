import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.087.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.088.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_087"', 'Attribute VB_Name = "Goren_Claude_V2_088"')
content = content.replace("' VERSION: V2.087", "' VERSION: V2.088")
content = content.replace('Private Const APP_VERSION As String = "2.087"', 'Private Const APP_VERSION As String = "2.088"')

# Fix FetchBOIMonthlyAvg
fetch_old = """Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
Static apiFailed As Boolean
If apiFailed Then
    FetchBOIMonthlyAvg = 0
    Exit Function
End If
On Error GoTo FAIL
Dim xmlHttp As Object"""

fetch_new = """Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
Static apiFailed As Boolean
Static apiCallsCount As Long
If apiFailed Or apiCallsCount > 40 Then
    FetchBOIMonthlyAvg = 0
    Exit Function
End If
apiCallsCount = apiCallsCount + 1
On Error GoTo FAIL
Dim xmlHttp As Object"""

content = content.replace(fetch_old, fetch_new)

# Fix GetBordMonth to filter absurd years
bord_old = """GetBordMonth = Format$(dt, "yyyy-mm")
Exit Function"""

bord_new = """Dim yr As Long
yr = Year(dt)
If yr >= 2000 And yr <= 2050 Then
    GetBordMonth = Format$(dt, "yyyy-mm")
Else
    GetBordMonth = ""
End If
Exit Function"""

content = content.replace(bord_old, bord_new)

# Fix branch logic in BuildReview just in case brLastRow is huge
br_old = """While Trim$(CStr(wsBranches.Cells(brLastRow, 1).Value2)) = "" And brLastRow > 3
brLastRow = brLastRow - 1
Wend
Dim brExists As Boolean
brExists = False
Dim brScan As Long
For brScan = 3 To brLastRow
If StrComp(Trim$(CStr(wsBranches.Cells(brScan, 1).Value2)), brOrigName, vbTextCompare) = 0 Then
brExists = True
Exit For
End If
Next brScan"""

br_new = """Dim brExists As Boolean
brExists = False
Dim matchBr As Variant
matchBr = Application.Match(brOrigName, wsBranches.Columns(1), 0)
If Not IsError(matchBr) Then
    brExists = True
End If"""

content = content.replace(br_old, br_new)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.088 created with API Cap, Date Filter, and Branch Match fix.")
