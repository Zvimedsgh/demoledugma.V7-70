Attribute VB_Name = "modFormatting"
Option Explicit

Public Sub SetAllSheetsRightToLeft()

    Dim ws As Worksheet

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        ws.DisplayRightToLeft = True
    Next ws

    Application.ScreenUpdating = True

    MsgBox "All worksheets were set to RTL", vbInformation

End Sub
