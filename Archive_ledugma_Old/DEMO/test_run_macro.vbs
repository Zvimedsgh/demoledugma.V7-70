On Error Resume Next
Set xlApp = GetObject(, "Excel.Application")
If xlApp Is Nothing Then
    WScript.Echo "Excel is not open."
    WScript.Quit
End If
On Error GoTo 0

On Error Resume Next
xlApp.Run "A00_SetupMainSheet_V9_40"
If Err.Number <> 0 Then
    WScript.Echo "Error running macro: " & Err.Description
Else
    WScript.Echo "Macro ran successfully!"
End If
