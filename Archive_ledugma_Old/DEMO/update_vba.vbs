Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim wb
Set wb = xlApp.Workbooks.Open("C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70.xlsm")

Dim comp, cm, lines, newLines
For Each comp In wb.VBProject.VBComponents
    If comp.Type = 1 Then ' vbext_ct_StdModule
        Set cm = comp.CodeModule
        If cm.CountOfLines > 0 Then
            lines = cm.Lines(1, cm.CountOfLines)
            If InStr(lines, """F20""") > 0 Then
                newLines = Replace(lines, """F20""", """F16""")
                cm.DeleteLines 1, cm.CountOfLines
                cm.AddFromString newLines
            End If
        End If
    End If
Next

wb.Save
wb.Close
xlApp.Quit
