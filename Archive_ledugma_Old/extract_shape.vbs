On Error Resume Next
Set xlApp = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Could not connect to Excel."
    WScript.Quit
End If
On Error GoTo 0

Dim shp
Dim found
found = False

For Each ws In xlApp.ActiveWorkbook.Worksheets
    On Error Resume Next
    Set shp = ws.Shapes("shpInstallMsg")
    If Err.Number = 0 Then
        found = True
        WScript.Echo "WIDTH: " & shp.Width
        WScript.Echo "HEIGHT: " & shp.Height
        WScript.Echo "TEXT: " & shp.TextFrame2.TextRange.Text
        Exit For
    End If
    On Error GoTo 0
Next

If Not found Then
    WScript.Echo "Shape not found."
End If
