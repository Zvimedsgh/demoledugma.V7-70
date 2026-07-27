import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.089.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.091.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_089"', 'Attribute VB_Name = "Goren_Claude_V2_091"')
content = content.replace("' VERSION: V2.089", "' VERSION: V2.091")
content = content.replace('Private Const APP_VERSION As String = "2.089"', 'Private Const APP_VERSION As String = "2.091"')

# Insert LogDebug at the top (after version block)
logger_func = """
' -------------------------------------------------------------------------
' LOGGER FOR DEBUGGING HANGS
' -------------------------------------------------------------------------
Private Sub LogDebug(ByVal msg As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open "c:\\LEVAV PROJECT\\SOURCE\\debug_log.txt" For Append As #ff
    Print #ff, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & msg
    Close #ff
    On Error GoTo 0
End Sub
"""

content = content.replace('Private Const APP_VERSION As String = "2.091"', 'Private Const APP_VERSION As String = "2.091"\n' + logger_func)

# We use simple string replaces instead of regex where possible to avoid escape issues
content = content.replace("Public Sub BuildReview()\n", "Public Sub BuildReview()\n    LogDebug \"BuildReview STARTED\"\n")
content = content.replace("Public Sub BuildReview()\r\n", "Public Sub BuildReview()\r\n    LogDebug \"BuildReview STARTED\"\r\n")

content = content.replace("A00_SetupMainSheet", "LogDebug \"Calling A00_SetupMainSheet\"\n    A00_SetupMainSheet\n    LogDebug \"Finished A00_SetupMainSheet\"")
# Fix the double replace in the Sub definition:
content = content.replace("Public Sub LogDebug \"Calling A00_SetupMainSheet\"\n    A00_SetupMainSheet\n    LogDebug \"Finished A00_SetupMainSheet\"()", "Public Sub A00_SetupMainSheet()")
content = content.replace("Public Sub LogDebug \"Calling A00_SetupMainSheet\"\r\n    A00_SetupMainSheet\r\n    LogDebug \"Finished A00_SetupMainSheet\"()", "Public Sub A00_SetupMainSheet()")

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

content = content.replace("xmlHttp.send\n", "LogDebug \"Sending API Request for \" & currencyKey & \" \" & monthStr\nxmlHttp.send\nLogDebug \"API Request returned Status \" & xmlHttp.Status\n")
content = content.replace("xmlHttp.send\r\n", "LogDebug \"Sending API Request for \" & currencyKey & \" \" & monthStr\r\nxmlHttp.send\r\nLogDebug \"API Request returned Status \" & xmlHttp.Status\r\n")

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.091 created.")
