On Error Resume Next
Set xlApp = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Excel is not running."
    WScript.Quit
End If
On Error GoTo 0

WScript.Echo "Connected to Excel. Active Workbook: " & xlApp.ActiveWorkbook.Name
