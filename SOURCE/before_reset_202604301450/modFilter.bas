Attribute VB_Name = "modFilter"
Option Explicit

Private Const FILTER_LIST_SHEET As String = "filter_lists"

Public Sub InitFilterDropdowns()

    With Range("rngFilterType").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:= _
            H("1492 1499 1500") & "," & _
            H("1495 1489 1512 1492") & "," & _
            H("1505 1493 1499 1503") & "," & _
            H("1506 1504 1507") & "," & _
            H("1506 1504 1507 32 1502 1512 1499 1494") & "," & _
            H("1496 1500 1512")
        .InCellDropdown = True
    End With

    RefreshFilterLists
    ApplyFilterValueDropdown

End Sub

Public Sub HandleFilterTypeChange(ByVal Target As Range)

    If Intersect(Target, Range("rngFilterType")) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    ApplyFilterValueDropdown
    Application.EnableEvents = True

End Sub

Public Sub RefreshFilterLists()

    Dim wsSrc As Worksheet
    Dim wsList As Worksheet

    Set wsSrc = GetCurrentFixedSheet()
    If wsSrc Is Nothing Then
        MsgBox "Current fixed sheet was not found", vbExclamation
        
        Exit Sub
    End If

    Set wsList = GetOrCreateFilterListSheet()
    wsList.Cells.Clear

   BuildUniqueList wsSrc, wsList, 9, 1
BuildUniqueList wsSrc, wsList, 14, 2
BuildUniqueList wsSrc, wsList, 11, 3
BuildUniqueList wsSrc, wsList, 13, 4
BuildUniqueList wsSrc, wsList, 16, 5

MsgBox "Filter dropdowns were refreshed", vbInformation

    CreateOrUpdateName "lst_filter_companies", SafeListRange(wsList, 1)
    CreateOrUpdateName "lst_filter_agents", SafeListRange(wsList, 2)
    CreateOrUpdateName "lst_filter_branches", SafeListRange(wsList, 3)
    CreateOrUpdateName "lst_filter_main_branches", SafeListRange(wsList, 4)
    CreateOrUpdateName "lst_filter_tellers", SafeListRange(wsList, 5)

    wsList.Visible = xlSheetVeryHidden

End Sub

Public Sub ApplyFilterValueDropdown()

    Dim fType As String
    Dim tgt As Range

    fType = Trim$(CStr(Range("rngFilterType").Value))
    Set tgt = Range("rngFilterValue")

    tgt.Validation.Delete
    tgt.Value = vbNullString

    Select Case fType
        Case H("1492 1499 1500")
            Exit Sub

        Case H("1495 1489 1512 1492")
            AddListValidation tgt, "lst_filter_companies"

        Case H("1505 1493 1499 1503")
            AddListValidation tgt, "lst_filter_agents"

        Case H("1506 1504 1507")
            AddListValidation tgt, "lst_filter_branches"

        Case H("1506 1504 1507 32 1502 1512 1499 1494")
            AddListValidation tgt, "lst_filter_main_branches"

        Case H("1496 1500 1512")
            AddListValidation tgt, "lst_filter_tellers"
    End Select

End Sub

Private Sub BuildUniqueList(ByVal wsSrc As Worksheet, ByVal wsList As Worksheet, ByVal srcCol As Long, ByVal outCol As Long)

    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long
    Dim v As String
    Dim k As Variant
    Dim outRow As Long

    If srcCol <= 0 Then
        wsList.Cells(1, outCol).Value = ""
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        v = Trim$(CStr(wsSrc.Cells(r, srcCol).Value))
        If v <> "" Then
            If Not IsTotalRow(v) Then
                If Not dict.Exists(v) Then dict.Add v, v
            End If
        End If
    Next r

    outRow = 1
    For Each k In dict.keys
        wsList.Cells(outRow, outCol).Value = CStr(k)
        outRow = outRow + 1
    Next k

End Sub

Private Function FindCol(ByVal ws As Worksheet, ByVal aliases As Variant) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim a As Variant
    Dim headerText As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        headerText = Trim$(CStr(ws.Cells(1, c).Value))

        For Each a In aliases
            If StrComp(headerText, CStr(a), vbTextCompare) = 0 Then
                FindCol = c
                Exit Function
            End If
        Next a
    Next c

    FindCol = 0

End Function

Private Function GetCurrentFixedSheet() As Worksheet

    Dim y As String
    Dim nm As String

    y = Trim$(CStr(Range("rngCurrentYear").Value))

    ' ? YOUR REAL SHEET NAME
    nm = H("1489 1505 1497 1505") & "_" & y   ' αριρ_2025

    If SheetExists(nm) Then
        Set GetCurrentFixedSheet = ThisWorkbook.Worksheets(nm)
        Exit Function
    End If

    ' fallback (just in case)
    nm = y & "_fixed"
    If SheetExists(nm) Then
        Set GetCurrentFixedSheet = ThisWorkbook.Worksheets(nm)
        Exit Function
    End If

End Function

Private Function GetOrCreateFilterListSheet() As Worksheet

    If SheetExists(FILTER_LIST_SHEET) Then
        Set GetOrCreateFilterListSheet = ThisWorkbook.Worksheets(FILTER_LIST_SHEET)
    Else
        Set GetOrCreateFilterListSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateFilterListSheet.Name = FILTER_LIST_SHEET
    End If

End Function

Private Function SafeListRange(ByVal ws As Worksheet, ByVal colNum As Long) As Range

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    Set SafeListRange = ws.Range(ws.Cells(1, colNum), ws.Cells(lastRow, colNum))

End Function

Private Sub AddListValidation(ByVal tgt As Range, ByVal listName As String)

    With tgt.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & listName
        .InCellDropdown = True
    End With

End Sub

Private Sub CreateOrUpdateName(ByVal nm As String, ByVal rng As Range)

    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=nm, RefersTo:="=" & rng.Address(True, True, xlA1, True)

End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function

Private Function IsTotalRow(ByVal txt As String) As Boolean
    IsTotalRow = (InStr(1, txt, H("1505 1492"), vbTextCompare) > 0)
End Function

Private Function H(ByVal codes As String) As String

    Dim arr() As String
    Dim i As Long
    Dim s As String

    arr = Split(codes, " ")

    For i = LBound(arr) To UBound(arr)
        s = s & ChrW(CLng(arr(i)))
    Next i

    H = s

End Function

