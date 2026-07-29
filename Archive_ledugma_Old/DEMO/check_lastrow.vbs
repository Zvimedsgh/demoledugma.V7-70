Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Dim wb
Set wb = xlApp.Workbooks.Open("C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70.xlsm")
Dim ws
Set ws = wb.Worksheets("ניהול")

Dim lastRow
lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row ' -4162 is xlUp

WScript.StdOut.WriteLine "Last row in Management sheet Column 1: " & lastRow

Dim r, valA, valB
For r = 100 To 105
    valA = ws.Cells(r, 1).Value
    valB = ws.Cells(r, 2).Value
    WScript.StdOut.WriteLine "Row " & r & ": A=" & valA & ", B=" & valB
Next

wb.Close False
xlApp.Quit
