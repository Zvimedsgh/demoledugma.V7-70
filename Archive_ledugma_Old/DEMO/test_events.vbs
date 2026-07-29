Set xl = CreateObject("Excel.Application")
xl.Visible = False
On Error Resume Next
Set wb = xl.Workbooks.Open("C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70_patched.xlsm")
If Err.Number <> 0 Then
    WScript.Echo "Error opening workbook: " & Err.Description
    WScript.Quit
End If
On Error GoTo 0

Set ws = wb.Sheets("ראשי")
WScript.Echo "Initial G6: " & ws.Range("G6").Value
WScript.Echo "Initial G7: " & ws.Range("G7").Value

' Enable events just in case
xl.EnableEvents = True

WScript.Echo "Changing G6 to רבעוני"
ws.Range("G6").Value = "רבעוני"
WScript.Echo "New G7: " & ws.Range("G7").Value
WScript.Echo "G7 Validation Formula: " & ws.Range("G7").Validation.Formula1

wb.Close False
xl.Quit
