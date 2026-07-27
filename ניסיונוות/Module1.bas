Attribute VB_Name = "Module1"
Option Explicit

Private Const SOURCE_FOLDER As String = "C:\SOURCE\"
Private Const MAIN_SHEET As String = "Main"
Private Const MGMT_SHEET As String = "NIHUL"

Private Const MGMT_DATA_ROW As Long = 3

Private Const MAP_COL_KEY As Long = 1
Private Const FIELDS_COL_NAME As Long = 5
Private Const FIELDS_COL_LETTER As Long = 6
Private Const FIELDS_COL_FLAG As Long = 7

Private Const PARAM_COL_NAME As Long = 10
Private Const PARAM_COL_VALUE As Long = 11

Public Sub UploadAndProcessFile()

    Dim wbSource As Workbook
    Dim wsData As Worksheet
    Dim wsReview As Worksheet
    Dim wsMgmt As Worksheet
    
    Dim dictMap As Object, dictCols As Object, dictNames As Object
    
    Dim r As Long, lastRow As Long, lastCol As Long
    Dim reviewRow As Long
    Dim reason As String
    Dim threshold As Double
    
    Dim currentYear As String
    Dim sourceFile As String, processedFile As String

    On Error GoTo ERR_HANDLER

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    currentYear = Trim(ThisWorkbook.Sheets(MAIN_SHEET).Range("B3").Value)
    If currentYear = "" Then MsgBox "B3 ריק": Exit Sub

    Set wsMgmt = ThisWorkbook.Sheets(MGMT_SHEET)
    
    threshold = GetThreshold(wsMgmt)
    If threshold <= 0 Then MsgBox "סף פרמיה לא תקין": Exit Sub

    Set dictMap = BuildMap(wsMgmt)
    Set dictCols = CreateObject("Scripting.Dictionary")
    Set dictNames = CreateObject("Scripting.Dictionary")

    LoadFields wsMgmt, dictCols, dictNames

    sourceFile = SOURCE_FOLDER & currentYear & ".xlsx"
    processedFile = SOURCE_FOLDER & currentYear & "_metukan.xlsx"

    Set wbSource = Workbooks.Open(sourceFile)
    Set wsData = wbSource.Sheets(1)

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    On Error Resume Next
    wbSource.Sheets("לטיפול").Delete
    On Error GoTo 0

    Set wsReview = wbSource.Sheets.Add
    wsReview.Name = "לטיפול"

    wsReview.Rows(1).Value = wsData.Rows(1).Value
    wsReview.Cells(1, lastCol + 1).Value = "סיבת חריגה"

    reviewRow = 2

    For r = 2 To lastRow

        reason = ""

        If WorksheetFunction.CountA(wsData.Rows(r)) = 0 Then GoTo NextRow

        Dim k As Variant
        For Each k In dictCols.Keys
            If Len(Trim(wsData.Cells(r, dictCols(k)).Text)) = 0 Then
                reason = AddReason(reason, "חסר " & dictNames(k))
            End If
        Next k

        ' מיפוי ענף
        If dictCols.Exists("שם_ענף") Then
            If Not dictMap.Exists(Trim(wsData.Cells(r, dictCols("שם_ענף")).Text)) Then
                reason = AddReason(reason, "חוסר שיוך לענף מרכז")
            End If
        End If

        ' פרמיה
        If dictCols.Exists("פרמיה") Then
            If Abs(wsData.Cells(r, dictCols("פרמיה")).Value) > threshold Then
                reason = AddReason(reason, "פרמיה חריגה")
            End If
        End If

        If reason <> "" Then
            wsReview.Rows(reviewRow).Value = wsData.Rows(r).Value
            wsReview.Cells(reviewRow, lastCol + 1).Value = reason
            reviewRow = reviewRow + 1
        End If

NextRow:
    Next r

    ' ======= ?? תיקון הבעיה שלך כאן =======

    On Error Resume Next
    Workbooks(currentYear & "_metukan.xlsx").Close False
    On Error GoTo 0

    If Dir(processedFile) <> "" Then Kill processedFile

    wbSource.SaveAs processedFile, xlOpenXMLWorkbook
    wbSource.Close False

    Workbooks.Open processedFile

    MsgBox "הושלם"

    Exit Sub

ERR_HANDLER:
    MsgBox "שגיאה: " & Err.Description

End Sub

' -------- פונקציות --------

Private Function LoadFields(ws As Worksheet, dictCols, dictNames)

    Dim r As Long: r = MGMT_DATA_ROW

    Do While ws.Cells(r, FIELDS_COL_NAME).Value <> ""

        If ws.Cells(r, FIELDS_COL_FLAG).Value = "בודקים" Then
            dictCols(ws.Cells(r, FIELDS_COL_NAME).Value) = ColumnLetterToNumber(ws.Cells(r, FIELDS_COL_LETTER).Value)
            dictNames(ws.Cells(r, FIELDS_COL_NAME).Value) = ws.Cells(r, FIELDS_COL_NAME).Value
        End If

        r = r + 1
    Loop

End Function

Private Function BuildMap(ws As Worksheet)

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Dim r As Long: r = MGMT_DATA_ROW

    Do While ws.Cells(r, MAP_COL_KEY).Value <> ""
        dict(Trim(ws.Cells(r, MAP_COL_KEY).Value)) = True
        r = r + 1
    Loop

    Set BuildMap = dict

End Function

Private Function GetThreshold(ws As Worksheet) As Double

    Dim r As Long: r = MGMT_DATA_ROW

    Do While ws.Cells(r, PARAM_COL_NAME).Value <> ""

        If ws.Cells(r, PARAM_COL_NAME).Value = "סף_פרמיה" Then
            GetThreshold = ws.Cells(r, PARAM_COL_VALUE).Value
            Exit Function
        End If

        r = r + 1
    Loop

End Function

Private Function ColumnLetterToNumber(colLetter As String) As Long

    Dim i As Long
    For i = 1 To Len(colLetter)
        ColumnLetterToNumber = ColumnLetterToNumber * 26 + _
            (Asc(UCase(Mid(colLetter, i, 1))) - 64)
    Next i

End Function

Private Function AddReason(existing As String, newReason As String) As String
    If existing = "" Then
        AddReason = newReason
    Else
        AddReason = existing & ", " & newReason
    End If
End Function

