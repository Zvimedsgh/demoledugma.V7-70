import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.089.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.090.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_089"', 'Attribute VB_Name = "Goren_Claude_V2_090"')
content = content.replace("' VERSION: V2.089", "' VERSION: V2.090")
content = content.replace('Private Const APP_VERSION As String = "2.089"', 'Private Const APP_VERSION As String = "2.090"')

# We will inject a Logger function at the top of the file
logger_func = """
' -------------------------------------------------------------------------
' LOGGER FOR DEBUGGING HANGS
' -------------------------------------------------------------------------
Private Sub LogDebug(ByVal msg As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open "c:\LEVAV PROJECT\SOURCE\debug_log.txt" For Append As #ff
    Print #ff, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & msg
    Close #ff
    On Error GoTo 0
End Sub
"""

# Insert logger func after Option Explicit
content = content.replace("Option Explicit", "Option Explicit\n" + logger_func)

# Inject into BuildReview
content = content.replace("Public Sub BuildReview()", "Public Sub BuildReview()\n    LogDebug \"BuildReview STARTED\"")
content = content.replace("A00_SetupMainSheet", "LogDebug \"Calling A00_SetupMainSheet\"\n    A00_SetupMainSheet\n    LogDebug \"Finished A00_SetupMainSheet\"")
content = content.replace("530     Dim isDemoMode As Boolean", "LogDebug \"Starting Demo Mode checks\"\n530     Dim isDemoMode As Boolean")
content = content.replace("560         Set wsSrc = OpenDataSheet(wbSrc)", "LogDebug \"Calling OpenDataSheet\"\n560         Set wsSrc = OpenDataSheet(wbSrc)\nLogDebug \"Finished OpenDataSheet\"")
content = content.replace("580         lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row", "580         lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row\nLogDebug \"lastRow is \" & lastRow")
content = content.replace("585         Call BuildMatachCache(wsSrc, RAW_BORDEREU, lastRow)", "LogDebug \"Calling BuildMatachCache\"\n585         Call BuildMatachCache(wsSrc, RAW_BORDEREU, lastRow)\nLogDebug \"Finished BuildMatachCache\"")
content = content.replace("600         srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, 50)).Value2", "LogDebug \"Loading srcData array (1 to 50)...\"\n600         srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, 50)).Value2\nLogDebug \"Finished loading srcData array.\"")
content = content.replace("610         BuildArrays dictFieldCol, cols, keys, cnt, threshold", "LogDebug \"Calling BuildArrays\"\n610         BuildArrays dictFieldCol, cols, keys, cnt, threshold\nLogDebug \"Finished BuildArrays\"")
content = content.replace("710     For r = 2 To lastRow", "LogDebug \"Starting main row loop (2 to \" & lastRow & \")\"\n710     For r = 2 To lastRow")
content = content.replace("990     Next r", "990     Next r\nLogDebug \"Finished main row loop. outIdx = \" & outIdx")
content = content.replace("991     listsName = H_RESHIMOT()", "LogDebug \"Writing lists to reshimot...\"\n991     listsName = H_RESHIMOT()")
content = content.replace("1000    actionCol = cnt + 3", "LogDebug \"Adding Validation...\"\n1000    actionCol = cnt + 3")
content = content.replace("1119    Dim shpBtn As Shape", "LogDebug \"Adding btnSendForReview...\"\n1119    Dim shpBtn As Shape")
content = content.replace("1235    UpdatePeriodDropdown", "LogDebug \"Calling UpdatePeriodDropdown\"\n1235    UpdatePeriodDropdown\nLogDebug \"Finished UpdatePeriodDropdown\"")
content = content.replace("1245    UpdateClientList", "LogDebug \"Calling UpdateClientList\"\n1245    UpdateClientList\nLogDebug \"Finished UpdateClientList\"")
content = content.replace("1250    Exit Sub", "LogDebug \"BuildReview ENDED SUCCESSFULLY\"\n1250    Exit Sub")

# Inject into FetchBOIMonthlyAvg
content = content.replace("xmlHttp.send", "LogDebug \"Sending API Request for \" & currencyKey & \" \" & monthStr\nxmlHttp.send\nLogDebug \"API Request returned Status \" & xmlHttp.Status")

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.090 created with full logging.")
