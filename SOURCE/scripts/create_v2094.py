import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.093.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.094.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_093"', 'Attribute VB_Name = "Goren_Claude_V2_094"')
content = content.replace("' VERSION: V2.093", "' VERSION: V2.094")
content = content.replace('Private Const APP_VERSION As String = "2.093"', 'Private Const APP_VERSION As String = "2.094"')

# Replace fallback rates in A00_SetupMainSheet
content = content.replace('wsMain.Range("K3").Value = 3.6', 'wsMain.Range("K3").Value = 3.002')
content = content.replace('wsMain.Range("K4").Value = 4#', 'wsMain.Range("K4").Value = 3.4239')

# Now let's restore FetchBOIMonthlyAvg with WinHttpRequest
fetch_new = """Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
    Static apiFailed As Boolean
    Static apiCallsCount As Long
    If apiFailed Or apiCallsCount > 40 Then
        FetchBOIMonthlyAvg = 0
        Exit Function
    End If
    apiCallsCount = apiCallsCount + 1
    
    On Error GoTo FAIL
    Dim xmlHttp As Object
    Dim url As String
    Dim responseText As String
    Dim lines() As String
    Dim fields() As String
    Dim i As Long
    Dim lineRate As Double
    Dim sumRate As Double
    Dim countRate As Long
    Dim startDate As String
    Dim endDate As String
    Dim lastValidRate As Double

    Set xmlHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' Strict 2 second timeout for Resolve, Connect, Send, Receive
    xmlHttp.SetTimeouts 2000, 2000, 2000, 2000

    If monthStr = "" Then
        ' Fetch most recent rate (last 7 days)
        startDate = Format$(Date - 7, "yyyy-mm-dd")
        endDate = Format$(Date, "yyyy-mm-dd")
    Else
        ' Fetch all daily rates for the entire month
        startDate = monthStr & "-01"
        endDate = monthStr & "-31"
    End If

    url = "https://edge.boi.gov.il/FusionEdgeServer/sdmx/v2/data/dataflow/BOI.STATISTICS/EXR/1.0/" & _
          "?c%5BBASE_CURRENCY%5D=" & UCase$(currencyKey) & _
          "&c%5BCOUNTER_CURRENCY%5D=ILS" & _
          "&c%5BDATA_TYPE%5D=OF00" & _
          "&startPeriod=" & startDate & _
          "&endPeriod=" & endDate & _
          "&format=csv"

    xmlHttp.Open "GET", url, False
    xmlHttp.send

    If xmlHttp.Status <> 200 Then GoTo FAIL

    responseText = xmlHttp.responseText
    If Len(responseText) < 20 Then GoTo FAIL

    lastValidRate = 0
    lines = Split(responseText, vbLf)
    sumRate = 0
    countRate = 0

    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 10 Then
            fields = Split(lines(i), ",")
            If UBound(fields) >= 13 Then
                If IsNumeric(Trim$(fields(13))) Then
                    lineRate = CDbl(Trim$(fields(13)))
                    If lineRate > 0 Then
                        sumRate = sumRate + lineRate
                        countRate = countRate + 1
                        lastValidRate = lineRate
                    End If
                End If
            End If
        End If
    Next i

    If countRate > 0 Then
        If monthStr = "" Then
            FetchBOIMonthlyAvg = lastValidRate
        Else
            FetchBOIMonthlyAvg = Round(sumRate / countRate, 4)
        End If
    End If
    Exit Function

FAIL:
    apiFailed = True
    FetchBOIMonthlyAvg = 0
End Function"""

# We need to replace the old neutered FetchBOIMonthlyAvg which currently looks like:
# Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
#     FetchBOIMonthlyAvg = 0
#     Exit Function
# Static apiFailed As Boolean
# ...
# End Function
pattern = re.compile(r'Private Function FetchBOIMonthlyAvg\(.*?End Function', re.DOTALL)
content = pattern.sub(fetch_new, content)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.094 created.")
