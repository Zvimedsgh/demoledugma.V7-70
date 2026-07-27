Attribute VB_Name = "ModLevav1"
' ============================================================================
' MODULE: modLevav
' PURPOSE: Complete system - BuildReview + ApplyCorrectionsAndBuildReports
' VERSION: 6.6 CLEAN FULL
' ============================================================================

' --- Windows API for Unicode MsgBox ---
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

' --- General constants ---
' SOURCE_FOLDER is now a function (see bottom of module) to support Hebrew path
' Sheet names are now functions to support Hebrew via ChrW
' CONTROL_SHEET_NAME = daf habait
' MANAGEMENT_SHEET_NAME = hagdarot
' REVIEW_SHEET_NAME = letipul
Private Const MANAGEMENT_START_ROW As Long = 2
Private Const DATA_SHEET_NAME As String = "TmpClientPolicyListEx"

' --- NIHUL field definition table ---
Private Const COL_FIELD_NAME_HE As Long = 5
Private Const COL_FIELD_COLUMN As Long = 6
Private Const COL_FIELD_CHECKING As Long = 7
Private Const COL_FIELD_KEY As Long = 8

' --- NIHUL parameter table ---
Private Const COL_PARAM_NAME As Long = 10
Private Const COL_PARAM_VALUE As Long = 11

' --- NIHUL helper translation table ---
Private Const COL_HELPER_KEY As Long = 14
Private Const COL_HELPER_VALUE As Long = 15

Private Const PARAM_PREMIUM_THRESHOLD As String = "PREMIUM_THRESHOLD"
Private Const PARAM_ERROR_EMAIL As String = "ERROR_EMAIL"
Private Const KEY_BRANCH_NAME As String = "BRANCH_NAME"
Private Const KEY_PREMIUM As String = "PREMIUM"
Private Const HELPER_REVIEW_SOURCE_ROW_HEADER As String = "REVIEW_SOURCE_ROW_HEADER"
Private Const HELPER_REVIEW_REASON_HEADER As String = "REVIEW_REASON_HEADER"
Private Const HELPER_REVIEW_REASON_CODE_HEADER As String = "REVIEW_REASON_CODE_HEADER"

' --- Output sheet names (Hebrew via functions below) ---

' --- Raw source column mapping ---
Private Const RAW_CUSTOMER As Long = 1
Private Const RAW_CUSTNAME As Long = 2
Private Const RAW_POLICY As Long = 11
Private Const RAW_ADDENDUM As Long = 12
Private Const RAW_COMPNUM As Long = 13
Private Const RAW_COMPANY As Long = 14
Private Const RAW_BRANCHNUM As Long = 15
Private Const RAW_BRANCHNAME As Long = 16
Private Const RAW_INSURANCE_START As Long = 17
Private Const RAW_BORDEREU As Long = 19
Private Const RAW_AGENTNUM As Long = 20
Private Const RAW_AGENTNAME As Long = 21
Private Const RAW_TELLERNUM As Long = 24
Private Const RAW_TELLERNAME As Long = 25
Private Const RAW_PREMIUM As Long = 28
Private Const RAW_COMMISSION As Long = 32
Private Const RAW_ACTIONCOL As Long = 39
Private Const RAW_IDNUMBER As Long = 45

' --- Base sheet columns ---
Private Const BASE_COL_ID As Long = 1
Private Const BASE_COL_YEAR As Long = 2
Private Const BASE_COL_MONTH As Long = 3
Private Const BASE_COL_IDENTITY As Long = 4
Private Const BASE_COL_CUSTOMER As Long = 5
Private Const BASE_COL_CUSTNAME As Long = 6
Private Const BASE_COL_POLICY As Long = 7
Private Const BASE_COL_ADDENDUM As Long = 8
Private Const BASE_COL_COMPANY As Long = 9
Private Const BASE_COL_COMPNUM As Long = 10
Private Const BASE_COL_BRANCHNAME As Long = 11
Private Const BASE_COL_BRANCHNUM As Long = 12
Private Const BASE_COL_MAINBRANCH As Long = 13
Private Const BASE_COL_AGENTNAME As Long = 14
Private Const BASE_COL_AGENTNUM As Long = 15
Private Const BASE_COL_TELLER As Long = 16
Private Const BASE_COL_TELLERNUM As Long = 17
Private Const BASE_COL_ACTION As Long = 18
Private Const BASE_COL_PREMIUM As Long = 19
Private Const BASE_COL_COMMISSION As Long = 20
Private Const BASE_COL_ISSUE As Long = 21
Private Const BASE_COL_TOFIX As Long = 22


' ============================================================================
' HELPER: Unicode MsgBox wrapper (uses Windows API MessageBoxW)
' ============================================================================
Private Const MB_RTLREADING As Long = &H100000
Private Const MB_RIGHT As Long = &H80000

Private Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
    MsgBoxU = MessageBoxW(0, StrPtr(sText), StrPtr(sCaption), uType Or MB_RTLREADING Or MB_RIGHT)
End Function


' ============================================================================
' HELPER: Find source file - supports both .xlsx and .xls
' ============================================================================
Private Function FindSourceFile(ByVal yearVal As String) As String
10      Dim p As String
20      Dim fso As Object
30      Set fso = CreateObject("Scripting.FileSystemObject")
40      p = SOURCE_FOLDER() & yearVal & ".xlsx"
50      If fso.FileExists(p) Then
60          FindSourceFile = p
70          Exit Function
80      End If
90      p = SOURCE_FOLDER() & yearVal & ".xls"
100     If fso.FileExists(p) Then
110         FindSourceFile = p
120         Exit Function
130     End If
140     FindSourceFile = ""
End Function


' ============================================================================
' HELPER: Open data worksheet from source workbook
' ============================================================================
Private Function OpenDataSheet(ByVal wb As Workbook) As Worksheet
10      On Error Resume Next
20      Dim ws As Worksheet
30      Set ws = wb.Worksheets(DATA_SHEET_NAME)
40      On Error GoTo 0
50      If ws Is Nothing Then
60          Set ws = wb.Worksheets(1)
70      End If
80      Set OpenDataSheet = ws
End Function


' ============================================================================
' HELPER: Get month range for comparison period from B4+C4
' Returns minMonth and maxMonth via ByRef
' ============================================================================
Private Sub GetMonthRange(ByVal wsMain As Worksheet, ByRef minMonth As Long, ByRef maxMonth As Long)
10      Dim periodType As String
20      Dim periodDetail As String
        Dim wsMgmt As Worksheet
        Dim monthIdx As Long
        Dim monthName As String
30      periodType = Trim$(CStr(wsMain.Range("B4").Value2))
40      periodDetail = Trim$(CStr(wsMain.Range("E4").Value2))

        ' Default: full year
50      minMonth = 1
60      maxMonth = 12

        ' "chodshi" = monthly
70      If InStr(1, periodType, ChrW$(1495) & ChrW$(1493) & ChrW$(1491) & ChrW$(1513) & ChrW$(1497), vbTextCompare) > 0 Then
            ' E4 contains Hebrew month name from NIHUL!R10:R21
            ' Match it against the month list to find month number
80          If periodDetail <> "" Then
90              Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
100             For monthIdx = 1 To 12
110                 monthName = Trim$(CStr(wsMgmt.Cells(9 + monthIdx, 18).Value2))
120                 If StrComp(periodDetail, monthName, vbTextCompare) = 0 Then
130                     minMonth = monthIdx
140                     maxMonth = monthIdx
150                     Exit For
160                 End If
170             Next monthIdx
180         End If

        ' "riv'oni" = quarterly
190     ElseIf InStr(1, periodType, ChrW$(1512) & ChrW$(1489) & ChrW$(1506) & ChrW$(1493) & ChrW$(1504) & ChrW$(1497), vbTextCompare) > 0 Then
            ' E4 = riv'on rishon/sheni/shlishi/revi'i from NIHUL!R5:R8
            ' Match by checking which quarter keyword is in the detail
200         If InStr(1, periodDetail, ChrW$(1512) & ChrW$(1488) & ChrW$(1513) & ChrW$(1493) & ChrW$(1503), vbTextCompare) > 0 Then
210             minMonth = 1: maxMonth = 3
220         ElseIf InStr(1, periodDetail, ChrW$(1513) & ChrW$(1504) & ChrW$(1497), vbTextCompare) > 0 Then
230             minMonth = 4: maxMonth = 6
240         ElseIf InStr(1, periodDetail, ChrW$(1513) & ChrW$(1500) & ChrW$(1497) & ChrW$(1513) & ChrW$(1497), vbTextCompare) > 0 Then
250             minMonth = 7: maxMonth = 9
260         ElseIf InStr(1, periodDetail, ChrW$(1512) & ChrW$(1489) & ChrW$(1497) & ChrW$(1506) & ChrW$(1497), vbTextCompare) > 0 Then
270             minMonth = 10: maxMonth = 12
280         End If

        ' "chatzi shnati" = half yearly
290     ElseIf InStr(1, periodType, ChrW$(1495) & ChrW$(1510) & ChrW$(1497), vbTextCompare) > 0 Then
            ' E4 = machatzit rishona/shniya from NIHUL!R2:R3
300         If InStr(1, periodDetail, ChrW$(1512) & ChrW$(1488) & ChrW$(1513) & ChrW$(1493) & ChrW$(1504), vbTextCompare) > 0 Then
310             minMonth = 1: maxMonth = 6
320         ElseIf InStr(1, periodDetail, ChrW$(1513) & ChrW$(1504) & ChrW$(1497), vbTextCompare) > 0 Then
330             minMonth = 7: maxMonth = 12
340         End If

        ' "shnatit" or anything else = full year (already set as default)
350     End If
End Sub

' ============================================================================
' HELPER: Get date column based on B5 selection
' ============================================================================
Private Function GetDateColumn(ByVal wsMain As Worksheet) As Long
10      Dim v As String
20      v = Trim$(CStr(wsMain.Range("B5").Value2))
        ' Hebrew: insurance start
30      If InStr(1, v, ChrW$(1514) & ChrW$(1495) & ChrW$(1497) & ChrW$(1500) & ChrW$(1514), vbTextCompare) > 0 Then
40          GetDateColumn = RAW_INSURANCE_START
50      Else
            ' Default: bordereu
60          GetDateColumn = RAW_BORDEREU
70      End If
End Function


' ============================================================================
' MACRO 1: BuildReview
' ============================================================================


Public Sub BuildReview()

    Dim wsReview As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim actionText As String
    Dim correctionText As String
    Dim fixedCount As Long

    On Error GoTo FailHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsReview = ThisWorkbook.Worksheets(ReviewSheetName())
    lastRow = wsReview.Cells(wsReview.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow

        actionText = UCase$(Trim$(CStr(wsReview.Cells(r, "N").Value)))
        correctionText = Trim$(CStr(wsReview.Cells(r, "O").Value))

        If actionText = "FIX" And correctionText <> "" Then
ApplyCorrectionFromReviewRow_v8 wsReview, r
            fixedCount = fixedCount + 1
        End If

    Next r

SafeExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "BuildReview finished. Fixed rows: " & fixedCount, vbInformation
    Exit Sub

FailHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "BuildReview failed at row " & r & ": " & Err.Description, vbExclamation

End Sub

Private Function ReviewSheetName() As String
    ReviewSheetName = ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
End Function
Private Sub ApplyCorrectionFromReviewRow_v8(ByVal wsReview As Worksheet, ByVal reviewRow As Long)

    Dim targetSheetName As String
    Dim targetRow As Long
    Dim targetCol As Long
    Dim correctionValue As Variant

    targetSheetName = CStr(wsReview.Cells(reviewRow, "B").Value)
    targetRow = CLng(wsReview.Cells(reviewRow, "C").Value)
    targetCol = CLng(wsReview.Cells(reviewRow, "D").Value)
    correctionValue = wsReview.Cells(reviewRow, "O").Value

    If targetSheetName = "" Then Exit Sub
    If targetRow <= 0 Then Exit Sub
    If targetCol <= 0 Then Exit Sub

    ThisWorkbook.Worksheets(targetSheetName).Cells(targetRow, targetCol).Value = correctionValue

End Sub

Private Sub BuildComparisonSheet(ByVal wsCurrent As Worksheet, ByVal wsRef As Worksheet, ByVal sheetName As String, ByVal groupCol As Long, ByVal minMonth As Long, ByVal maxMonth As Long, ByVal yearVal As String, ByVal refYear As String)

10      On Error GoTo ERR_HANDLER

20      DeleteSheetIfExists sheetName
30      Dim wsOut As Worksheet
40      Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
50      wsOut.Name = sheetName

        ' Collect unique keys from both sheets
60      Dim dictKeys As Object
70      Set dictKeys = CreateObject("Scripting.Dictionary")
80      dictKeys.CompareMode = vbTextCompare

90      Dim lastRowCur As Long
100     Dim lastRowRef As Long
110     Dim r As Long
120     Dim k As String
130     Dim m As Long
        Dim sortedMonths() As String
        Dim mIdx As Long
        Dim mCnt As Long
        Dim mi As Long
        Dim dictGlobalCustCur As Object
        Dim dictGlobalCustRef As Object
        Dim dictGlobalPolCur As Object
        Dim dictGlobalPolRef As Object


140     lastRowCur = wsCurrent.Cells(wsCurrent.Rows.Count, 1).End(xlUp).Row
150     lastRowRef = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row

160     For r = 2 To lastRowCur
170         m = 0
180         If Not IsBlankValue(wsCurrent.Cells(r, BASE_COL_MONTH).Value2) Then
190             m = CLng(wsCurrent.Cells(r, BASE_COL_MONTH).Value2)
200         End If
210         If m >= minMonth And m <= maxMonth Then
220             k = Trim$(CStr(wsCurrent.Cells(r, groupCol).Value2))
230             If k <> "" And LCase$(k) <> "(empty)" Then
240                 If Not dictKeys.Exists(k) Then dictKeys(k) = True
250             End If
260         End If
270     Next r

280     For r = 2 To lastRowRef
290         m = 0
300         If Not IsBlankValue(wsRef.Cells(r, BASE_COL_MONTH).Value2) Then
310             m = CLng(wsRef.Cells(r, BASE_COL_MONTH).Value2)
320         End If
330         If m >= minMonth And m <= maxMonth Then
340             k = Trim$(CStr(wsRef.Cells(r, groupCol).Value2))
350             If k <> "" And LCase$(k) <> "(empty)" Then
360                 If Not dictKeys.Exists(k) Then dictKeys(k) = True
370             End If
380         End If
390     Next r

        ' Write headers
400     WriteComparisonHeaders wsOut, yearVal, refYear, sheetName

        ' For each key, aggregate values
410     Dim outRow As Long
420     outRow = 3
430     Dim allKeys As Variant

        ' Sort keys for months sheet (numeric 1-12)
        If groupCol = BASE_COL_MONTH Then
            mCnt = 0
            ReDim sortedMonths(1 To 12)
            For mi = 1 To 12
                If dictKeys.Exists(CStr(mi)) Then
                    mCnt = mCnt + 1
                    sortedMonths(mCnt) = CStr(mi)
                End If
            Next mi
            If mCnt > 0 Then
                ReDim Preserve sortedMonths(1 To mCnt)
                ReDim allKeys(0 To mCnt - 1)
                For mi = 1 To mCnt
                    allKeys(mi - 1) = sortedMonths(mi)
                Next mi
            Else
                allKeys = dictKeys.keys
            End If
        Else
440         allKeys = dictKeys.keys
        End If

450     Dim idx As Long
460     Dim premCur As Double
470     Dim premRef As Double
480     Dim commCur As Double
490     Dim commRef As Double
500     Dim docsCur As Long
510     Dim docsRef As Long
520     Dim custCur As Long
530     Dim custRef As Long
540     Dim polCur As Long
550     Dim polRef As Long
560     Dim dictCustCur As Object
570     Dim dictCustRef As Object
580     Dim dictPolCur As Object
590     Dim dictPolRef As Object
600     Dim custKey As String
610     Dim polKey As String
615     Dim actValCur As String
616     Dim actValRef As String
617     Dim gKey As Variant

620     Dim totPremCur As Double
630     Dim totPremRef As Double
640     Dim totCommCur As Double
650     Dim totCommRef As Double
660     Dim totDocsCur As Long
670     Dim totDocsRef As Long
680     Dim totCustCur As Long
690     Dim totCustRef As Long
700     Dim totPolCur As Long
710     Dim totPolRef As Long

        ' Global unique dicts for correct totals
        Set dictGlobalCustCur = CreateObject("Scripting.Dictionary")
        Set dictGlobalCustRef = CreateObject("Scripting.Dictionary")
        Set dictGlobalPolCur = CreateObject("Scripting.Dictionary")
        Set dictGlobalPolRef = CreateObject("Scripting.Dictionary")

720     For idx = 0 To UBound(allKeys)
730         k = allKeys(idx)

740         premCur = 0: premRef = 0
750         commCur = 0: commRef = 0
760         docsCur = 0: docsRef = 0

770         Set dictCustCur = CreateObject("Scripting.Dictionary")
780         Set dictCustRef = CreateObject("Scripting.Dictionary")
790         Set dictPolCur = CreateObject("Scripting.Dictionary")
800         Set dictPolRef = CreateObject("Scripting.Dictionary")

            ' Aggregate current year
810         For r = 2 To lastRowCur
820             m = 0
830             If Not IsBlankValue(wsCurrent.Cells(r, BASE_COL_MONTH).Value2) Then
840                 m = CLng(wsCurrent.Cells(r, BASE_COL_MONTH).Value2)
850             End If
860             If m < minMonth Or m > maxMonth Then GoTo NextCurRow

870             If StrComp(Trim$(CStr(wsCurrent.Cells(r, groupCol).Value2)), k, vbTextCompare) = 0 Then
880                 premCur = premCur + CDbl(wsCurrent.Cells(r, BASE_COL_PREMIUM).Value2)
890                 commCur = commCur + CDbl(wsCurrent.Cells(r, BASE_COL_COMMISSION).Value2)
900                 docsCur = docsCur + 1

                    ' Customer counting: unique customer numbers with at least one non-cancelled row
910                 custKey = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_CUSTOMER).Value2))
920                 If custKey <> "" Then
                        actValCur = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_ACTION).Value2))
                        ' Check if this row is NOT a cancellation
                        If InStr(1, actValCur, ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500), vbTextCompare) = 0 Then
930                         If Not dictCustCur.Exists(custKey) Then dictCustCur(custKey) = True
                        End If
940                 End If

950                 polKey = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_POLICY).Value2))
960                 If polKey <> "" Then
970                     If Not dictPolCur.Exists(polKey) Then dictPolCur(polKey) = True
980                 End If
990             End If
NextCurRow:
1000        Next r

            ' Aggregate reference year
1010        For r = 2 To lastRowRef
1020            m = 0
1030            If Not IsBlankValue(wsRef.Cells(r, BASE_COL_MONTH).Value2) Then
1040                m = CLng(wsRef.Cells(r, BASE_COL_MONTH).Value2)
1050            End If
1060            If m < minMonth Or m > maxMonth Then GoTo NextRefRow2

1070            If StrComp(Trim$(CStr(wsRef.Cells(r, groupCol).Value2)), k, vbTextCompare) = 0 Then
1080                premRef = premRef + CDbl(wsRef.Cells(r, BASE_COL_PREMIUM).Value2)
1090                commRef = commRef + CDbl(wsRef.Cells(r, BASE_COL_COMMISSION).Value2)
1100                docsRef = docsRef + 1

                    ' Customer counting: unique customer numbers with at least one non-cancelled row
1110                custKey = Trim$(CStr(wsRef.Cells(r, BASE_COL_CUSTOMER).Value2))
1120                If custKey <> "" Then
                        actValRef = Trim$(CStr(wsRef.Cells(r, BASE_COL_ACTION).Value2))
                        ' Check if this row is NOT a cancellation
                        If InStr(1, actValRef, ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500), vbTextCompare) = 0 Then
1130                        If Not dictCustRef.Exists(custKey) Then dictCustRef(custKey) = True
                        End If
1140                End If

1150                polKey = Trim$(CStr(wsRef.Cells(r, BASE_COL_POLICY).Value2))
1160                If polKey <> "" Then
1170                    If Not dictPolRef.Exists(polKey) Then dictPolRef(polKey) = True
1180                End If
1190            End If
NextRefRow2:
1200        Next r

1210        custCur = dictCustCur.Count
1220        custRef = dictCustRef.Count
1230        polCur = dictPolCur.Count
1240        polRef = dictPolRef.Count

            ' Write row - for months, show Hebrew name
            If groupCol = BASE_COL_MONTH Then
1250            wsOut.Cells(outRow, 1).Value = HebrewMonthName(CLng(k))
            Else
1251            wsOut.Cells(outRow, 1).Value = k
            End If
1260        wsOut.Cells(outRow, 2).Value = premRef
1270        wsOut.Cells(outRow, 3).Value = premCur
1280        wsOut.Cells(outRow, 4).Value = SafePct(premCur, premRef)
1290        wsOut.Cells(outRow, 5).Value = docsRef
1300        wsOut.Cells(outRow, 6).Value = docsCur
1310        wsOut.Cells(outRow, 7).Value = SafePct(docsCur, docsRef)
1320        wsOut.Cells(outRow, 8).Value = custRef
1330        wsOut.Cells(outRow, 9).Value = custCur
1340        wsOut.Cells(outRow, 10).Value = SafePct(custCur, custRef)
1350        wsOut.Cells(outRow, 11).Value = polRef
1360        wsOut.Cells(outRow, 12).Value = polCur
1370        wsOut.Cells(outRow, 13).Value = SafePct(polCur, polRef)
1380        wsOut.Cells(outRow, 14).Value = commRef
1390        wsOut.Cells(outRow, 15).Value = commCur
1400        wsOut.Cells(outRow, 16).Value = SafePct(commCur, commRef)

            ' Accumulate totals (premium, commission, docs are additive)
1410        totPremCur = totPremCur + premCur
1420        totPremRef = totPremRef + premRef
1430        totCommCur = totCommCur + commCur
1440        totCommRef = totCommRef + commRef
1450        totDocsCur = totDocsCur + docsCur
1460        totDocsRef = totDocsRef + docsRef

            ' Merge per-group unique dicts into global dicts
            For Each gKey In dictCustCur.keys
                If Not dictGlobalCustCur.Exists(gKey) Then dictGlobalCustCur(gKey) = True
            Next gKey
            For Each gKey In dictCustRef.keys
                If Not dictGlobalCustRef.Exists(gKey) Then dictGlobalCustRef(gKey) = True
            Next gKey
            For Each gKey In dictPolCur.keys
                If Not dictGlobalPolCur.Exists(gKey) Then dictGlobalPolCur(gKey) = True
            Next gKey
            For Each gKey In dictPolRef.keys
                If Not dictGlobalPolRef.Exists(gKey) Then dictGlobalPolRef(gKey) = True
            Next gKey

1510        outRow = outRow + 1
1520    Next idx

        ' Compute global unique totals for customers and policies
        totCustCur = dictGlobalCustCur.Count
        totCustRef = dictGlobalCustRef.Count
        totPolCur = dictGlobalPolCur.Count
        totPolRef = dictGlobalPolRef.Count

        ' Write totals row
1530    wsOut.Cells(outRow, 1).Value = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499)
1540    wsOut.Cells(outRow, 2).Value = totPremRef
1550    wsOut.Cells(outRow, 3).Value = totPremCur
1560    wsOut.Cells(outRow, 4).Value = SafePct(totPremCur, totPremRef)
1570    wsOut.Cells(outRow, 5).Value = totDocsRef
1580    wsOut.Cells(outRow, 6).Value = totDocsCur
1590    wsOut.Cells(outRow, 7).Value = SafePct(totDocsCur, totDocsRef)
1600    wsOut.Cells(outRow, 8).Value = totCustRef
1610    wsOut.Cells(outRow, 9).Value = totCustCur
1620    wsOut.Cells(outRow, 10).Value = SafePct(totCustCur, totCustRef)
1630    wsOut.Cells(outRow, 11).Value = totPolRef
1640    wsOut.Cells(outRow, 12).Value = totPolCur
1650    wsOut.Cells(outRow, 13).Value = SafePct(totPolCur, totPolRef)
1660    wsOut.Cells(outRow, 14).Value = totCommRef
1670    wsOut.Cells(outRow, 15).Value = totCommCur
1680    wsOut.Cells(outRow, 16).Value = SafePct(totCommCur, totCommRef)
1690    wsOut.Rows(outRow).Font.Bold = True

        ' Format
1700    wsOut.Columns.AutoFit
1710    Dim col As Long
1720    For col = 2 To 16
1730        If col = 4 Or col = 7 Or col = 10 Or col = 13 Or col = 16 Then
1740            wsOut.Columns(col).NumberFormat = "0.0%"
1750        Else
1760            wsOut.Columns(col).NumberFormat = "#,##0"
1770        End If
1780    Next col

1790    Exit Sub

ERR_HANDLER:
1800    Err.Raise Err.Number, "BuildComparisonSheet(" & sheetName & "):" & Erl, Err.Description
End Sub



' ============================================================================
' HELPER: Build comparison sheet with text include/exclude filter
' Used for Agents_No_Levav and Company_Drachim slices
' ============================================================================
Private Sub BuildComparisonSheetFiltered(ByVal wsCurrent As Worksheet, ByVal wsRef As Worksheet, ByVal sheetName As String, ByVal groupCol As Long, ByVal minMonth As Long, ByVal maxMonth As Long, ByVal yearVal As String, ByVal refYear As String, ByVal filterCol As Long, ByVal filterText As String, ByVal excludeMatch As Boolean)

10      On Error GoTo ERR_HANDLER

20      Dim tmpCur As Worksheet
30      Dim tmpRef As Worksheet
40      Dim lastRow As Long
50      Dim r As Long
60      Dim outRow As Long
70      Dim c As Long
80      Dim tmpCurName As String
90      Dim tmpRefName As String

100     tmpCurName = "_tmp_cur_" & Replace(sheetName, " ", "_")
110     tmpRefName = "_tmp_ref_" & Replace(sheetName, " ", "_")

120     DeleteSheetIfExists tmpCurName
130     DeleteSheetIfExists tmpRefName

140     Set tmpCur = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
150     tmpCur.Name = tmpCurName
160     Set tmpRef = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
170     tmpRef.Name = tmpRefName

180     For c = 1 To BASE_COL_TOFIX
190         tmpCur.Cells(1, c).Value = wsCurrent.Cells(1, c).Value
200         tmpRef.Cells(1, c).Value = wsRef.Cells(1, c).Value
210     Next c

220     outRow = 2
230     lastRow = wsCurrent.Cells(wsCurrent.Rows.Count, 1).End(xlUp).Row
240     For r = 2 To lastRow
250         If RowPassesTextFilter(wsCurrent, r, filterCol, filterText, excludeMatch) Then
260             For c = 1 To BASE_COL_TOFIX
270                 tmpCur.Cells(outRow, c).Value = wsCurrent.Cells(r, c).Value
280             Next c
290             outRow = outRow + 1
300         End If
310     Next r

320     outRow = 2
330     lastRow = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row
340     For r = 2 To lastRow
350         If RowPassesTextFilter(wsRef, r, filterCol, filterText, excludeMatch) Then
360             For c = 1 To BASE_COL_TOFIX
370                 tmpRef.Cells(outRow, c).Value = wsRef.Cells(r, c).Value
380             Next c
390             outRow = outRow + 1
400         End If
410     Next r

420     BuildComparisonSheet tmpCur, tmpRef, sheetName, groupCol, minMonth, maxMonth, yearVal, refYear

430     Application.DisplayAlerts = False
440     tmpCur.Delete
450     tmpRef.Delete
460     Application.DisplayAlerts = True

470     Exit Sub

ERR_HANDLER:
480     On Error Resume Next
490     Application.DisplayAlerts = False
500     If Not tmpCur Is Nothing Then tmpCur.Delete
510     If Not tmpRef Is Nothing Then tmpRef.Delete
520     Application.DisplayAlerts = True
530     Err.Raise Err.Number, "BuildComparisonSheetFiltered(" & sheetName & "):" & Erl, Err.Description
End Sub

Private Function RowPassesTextFilter(ByVal ws As Worksheet, ByVal r As Long, ByVal filterCol As Long, ByVal filterText As String, ByVal excludeMatch As Boolean) As Boolean

10      Dim v As String
20      Dim hasText As Boolean

30      v = Trim$(CStr(ws.Cells(r, filterCol).Value2))
40      hasText = (InStr(1, v, filterText, vbTextCompare) > 0)

50      If excludeMatch Then
60          RowPassesTextFilter = Not hasText
70      Else
80          RowPassesTextFilter = hasText
90      End If

End Function

' ============================================================================
' HELPER: Write comparison sheet headers (two-row merged layout)
' Row 1: category headers (merged, blue background, white bold text)
' Row 2: year sub-headers (light blue background, bold text)
' Data starts at row 3
' ============================================================================
Private Sub WriteComparisonHeaders(ByVal ws As Worksheet, ByVal yearVal As String, ByVal refYear As String, ByVal sheetName As String)
10      On Error GoTo ERR_HANDLER

        Dim blueColor As Long
        Dim lightBlueColor As Long
        Dim pctLabel As String
20      blueColor = RGB(0, 70, 140)
30      lightBlueColor = RGB(155, 200, 235)

        ' --- Row 1: Category headers (merged cells) ---
        ' Col A: name header (merged rows 1-2)
40      ws.Range("A1:A2").Merge
50      ws.Cells(1, 1).Value = ChrW(1513) & ChrW(1501)
60      ws.Cells(1, 1).HorizontalAlignment = xlCenter
70      ws.Cells(1, 1).VerticalAlignment = xlCenter

        ' Cols B-D: production (merged)
80      ws.Range(ws.Cells(1, 2), ws.Cells(1, 4)).Merge
90      ws.Cells(1, 2).Value = ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(1511) & ChrW(1510) & ChrW(1497) & ChrW(1492)
100     ws.Cells(1, 2).HorizontalAlignment = xlCenter

        ' Cols E-G: documents (merged)
110     ws.Range(ws.Cells(1, 5), ws.Cells(1, 7)).Merge
120     ws.Cells(1, 5).Value = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501)
130     ws.Cells(1, 5).HorizontalAlignment = xlCenter

        ' Cols H-J: insured (merged)
140     ws.Range(ws.Cells(1, 8), ws.Cells(1, 10)).Merge
150     ws.Cells(1, 8).Value = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501)
160     ws.Cells(1, 8).HorizontalAlignment = xlCenter

        ' Cols K-M: policies (merged)
170     ws.Range(ws.Cells(1, 11), ws.Cells(1, 13)).Merge
180     ws.Cells(1, 11).Value = ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1493) & ChrW(1514)
190     ws.Cells(1, 11).HorizontalAlignment = xlCenter

        ' Cols N-P: commission (merged)
200     ws.Range(ws.Cells(1, 14), ws.Cells(1, 16)).Merge
210     ws.Cells(1, 14).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1492)
220     ws.Cells(1, 14).HorizontalAlignment = xlCenter

        ' --- Row 1 formatting: blue background, white bold ---
230     ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Interior.Color = blueColor
240     ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Color = RGB(255, 255, 255)
250     ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Bold = True
260     ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Size = 12

        ' --- Row 2: Year sub-headers ---
        ' Hebrew: shinuy% = change%
270     pctLabel = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497) & "%"

280     ws.Cells(2, 1).Value = ""
290     ws.Cells(2, 2).Value = refYear
300     ws.Cells(2, 3).Value = yearVal
310     ws.Cells(2, 4).Value = pctLabel
320     ws.Cells(2, 5).Value = refYear
330     ws.Cells(2, 6).Value = yearVal
340     ws.Cells(2, 7).Value = pctLabel
350     ws.Cells(2, 8).Value = refYear
360     ws.Cells(2, 9).Value = yearVal
370     ws.Cells(2, 10).Value = pctLabel
380     ws.Cells(2, 11).Value = refYear
390     ws.Cells(2, 12).Value = yearVal
400     ws.Cells(2, 13).Value = pctLabel
410     ws.Cells(2, 14).Value = refYear
420     ws.Cells(2, 15).Value = yearVal
430     ws.Cells(2, 16).Value = pctLabel

        ' --- Row 2 formatting: light blue background, bold ---
440     ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).Interior.Color = lightBlueColor
450     ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).Font.Bold = True
460     ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).HorizontalAlignment = xlCenter

        ' --- Borders ---
470     ws.Range(ws.Cells(1, 1), ws.Cells(2, 16)).Borders.LineStyle = xlContinuous
480     ws.Range(ws.Cells(1, 1), ws.Cells(2, 16)).Borders.Weight = xlThin

490     Exit Sub

ERR_HANDLER:
500     Err.Raise Err.Number, "WriteComparisonHeaders:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build summary sheet
' ============================================================================
Private Sub BuildSummarySheet(ByVal countRef As Long, ByVal countCurrent As Long, ByVal reviewCount As Long, ByVal corrCount As Long, ByVal ignoreCount As Long, ByVal unhandledCount As Long, ByVal yearVal As String, ByVal refYear As String, ByVal periodDesc As String)

10      On Error GoTo ERR_HANDLER
20      DeleteSheetIfExists SHEET_SUMMARY()
30      Dim wsOut As Worksheet
40      Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
50      wsOut.Name = SHEET_SUMMARY()

60      wsOut.Cells(1, 1).Value = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1514) & ChrW(1492) & ChrW(1500) & ChrW(1497) & ChrW(1498)
70      wsOut.Cells(1, 1).Font.Bold = True
80      wsOut.Cells(1, 1).Font.Size = 14

90      wsOut.Cells(3, 1).Value = ChrW(1508) & ChrW(1512) & ChrW(1496)
100     wsOut.Cells(3, 2).Value = ChrW(1506) & ChrW(1512) & ChrW(1498)
110     wsOut.Rows(3).Font.Bold = True

200     wsOut.Cells(4, 1).Value = ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
210     wsOut.Cells(4, 2).Value = periodDesc
220     wsOut.Cells(5, 1).Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1497) & ChrW(1497) & ChrW(1495) & ChrW(1493) & ChrW(1505)
230     wsOut.Cells(5, 2).Value = refYear
240     wsOut.Cells(6, 1).Value = ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1504) & ChrW(1489) & ChrW(1491) & ChrW(1511) & ChrW(1514)
250     wsOut.Cells(6, 2).Value = yearVal
260     wsOut.Cells(7, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & refYear
270     wsOut.Cells(7, 2).Value = countRef
280     wsOut.Cells(8, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & yearVal & " (" & ChrW(1502) & ChrW(1514) & ChrW(1493) & ChrW(1511) & ChrW(1503) & ")"
290     wsOut.Cells(8, 2).Value = countCurrent
300     wsOut.Cells(9, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1489) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1512) & ChrW(1514)
310     wsOut.Cells(9, 2).Value = reviewCount
320     wsOut.Cells(10, 1).Value = ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1513) & ChrW(1497) & ChrW(1493) & ChrW(1513) & ChrW(1502) & ChrW(1493)
330     wsOut.Cells(10, 2).Value = corrCount
340     wsOut.Cells(11, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1513) & ChrW(1492) & ChrW(1493) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1502) & ChrW(1493)
350     wsOut.Cells(11, 2).Value = ignoreCount
360     wsOut.Cells(12, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
370     wsOut.Cells(12, 2).Value = unhandledCount

380     wsOut.Columns(2).NumberFormat = "#,##0"
390     wsOut.Columns.AutoFit
410     Exit Sub

ERR_HANDLER:
420     Err.Raise Err.Number, "BuildSummarySheet:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Safe percentage calculation
' ============================================================================
Private Function SafePct(ByVal newVal As Double, ByVal oldVal As Double) As Double
10      If oldVal = 0 Then
20          If newVal = 0 Then
30              SafePct = 0
40          Else
50              SafePct = 1
60          End If
70      Else
80          SafePct = (newVal - oldVal) / Abs(oldVal)
90      End If
End Function


' ============================================================================
' HELPER: Load helper dictionary from NIHUL columns N-O
' ============================================================================
Private Function LoadHelperDictionary(ByVal ws As Worksheet) As Object
10      On Error GoTo ERR_HANDLER
20      Dim dict As Object
30      Set dict = CreateObject("Scripting.Dictionary")
40      dict.CompareMode = vbTextCompare
50      Dim r As Long
60      Dim lastRow As Long
70      Dim k As String
80      lastRow = ws.Cells(ws.Rows.Count, COL_HELPER_KEY).End(xlUp).Row
90      For r = 1 To lastRow
100         k = Trim$(CStr(ws.Cells(r, COL_HELPER_KEY).Value2))
110         If k <> "" Then
120             dict(k) = Trim$(CStr(ws.Cells(r, COL_HELPER_VALUE).Value2))
130         End If
140     Next r
        ' Add fallback translations for missing codes
150     If Not dict.Exists("MISSING_CUSTOMER_NUMBER") Then dict("MISSING_CUSTOMER_NUMBER") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
151     If Not dict.Exists("MISSING_CUSTOMER_NAME") Then dict("MISSING_CUSTOMER_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
152     If Not dict.Exists("MISSING_POLICY") Then dict("MISSING_POLICY") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1492)
153     If Not dict.Exists("MISSING_ADDENDUM") Then dict("MISSING_ADDENDUM") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1514) & ChrW(1493) & ChrW(1505) & ChrW(1508) & ChrW(1514)
154     If Not dict.Exists("MISSING_COMPANY_NAME") Then dict("MISSING_COMPANY_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492)
155     If Not dict.Exists("MISSING_BRANCH_NAME") Then dict("MISSING_BRANCH_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1506) & ChrW(1504) & ChrW(1507)
156     If Not dict.Exists("MISSING_AGENT_NAME") Then dict("MISSING_AGENT_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503)
157     If Not dict.Exists("MISSING_UNDERWRITER_TELLER_NAME") Then dict("MISSING_UNDERWRITER_TELLER_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1496) & ChrW(1500) & ChrW(1512)
158     If Not dict.Exists("MISSING_CURRENCY") Then dict("MISSING_CURRENCY") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)
159     If Not dict.Exists("MISSING_PREMIUM") Then dict("MISSING_PREMIUM") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492)
160     If Not dict.Exists("MISSING_COMPANY_COMMISSION") Then dict("MISSING_COMPANY_COMMISSION") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1514) & " " & ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492)
161     If Not dict.Exists("PREMIUM_OVER_THRESHOLD") Then dict("PREMIUM_OVER_THRESHOLD") = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492) & " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1492)
162     If Not dict.Exists("PREMIUM_NOT_NUMERIC") Then dict("PREMIUM_NOT_NUMERIC") = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & ChrW(1497)

170     Set LoadHelperDictionary = dict
180     Exit Function
ERR_HANDLER:
190     Err.Raise Err.Number, "LoadHelperDictionary:" & Erl, Err.Description
End Function


' ============================================================================
' HELPER: Load branch mapping from NIHUL
' ============================================================================
Private Function LoadBranchMapping(ByVal ws As Worksheet) As Object
10      On Error GoTo ERR_HANDLER
20      Dim dict As Object
30      Set dict = CreateObject("Scripting.Dictionary")
40      dict.CompareMode = vbTextCompare
50      Dim r As Long
60      Dim lastRow As Long
70      Dim brName As String
80      Dim mainBr As String
90      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
100     For r = 3 To lastRow
110         brName = UCase$(Trim$(CStr(ws.Cells(r, 1).Value2)))
120         mainBr = Trim$(CStr(ws.Cells(r, 2).Value2))
130         If brName <> "" And mainBr <> "" Then
140             dict(brName) = mainBr
150         End If
160     Next r
170     Set LoadBranchMapping = dict
180     Exit Function
ERR_HANDLER:
190     Err.Raise Err.Number, "LoadBranchMapping:" & Erl, Err.Description
End Function


' ============================================================================
' HELPER: Check whether a branch already exists in Settings column A.
' ============================================================================
Private Function SettingsBranchExists(ByVal ws As Worksheet, ByVal branchName As String) As Boolean
10      Dim lastRow As Long
20      Dim r As Long
30      Dim Target As String
40      Target = UCase$(Trim$(branchName))
50      If Target = "" Then
60          SettingsBranchExists = True
70          Exit Function
80      End If
90      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
100     For r = 3 To lastRow
110         If UCase$(Trim$(CStr(ws.Cells(r, 1).Value2))) = Target Then
120             SettingsBranchExists = True
130             Exit Function
140         End If
150     Next r
160     SettingsBranchExists = False
End Function


' ============================================================================
' MACRO: Update Settings English columns.
' A = Branch Hebrew, B = Main Branch Hebrew, C = Branch English, D = Main Branch English.
' This macro never changes columns A or B.
' It repairs C/D when they are blank or UNKNOWN.
' ============================================================================
Public Sub UpdateSettingsEnglishCodes()
10      On Error GoTo ERR_HANDLER
20      Dim ws As Worksheet
30      Dim lastRow As Long
40      Dim r As Long
50      Dim branchHe As String
60      Dim mainHe As String
70      Dim currentC As String
80      Dim currentD As String
90      Dim updatedC As Long
100     Dim updatedD As Long
110     Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
120     lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
130     For r = 3 To lastRow
140         branchHe = Trim$(CStr(ws.Cells(r, 1).Value2))
150         mainHe = Trim$(CStr(ws.Cells(r, 2).Value2))
160         currentC = UCase$(Trim$(CStr(ws.Cells(r, 3).Value2)))
170         currentD = UCase$(Trim$(CStr(ws.Cells(r, 4).Value2)))
180         If branchHe <> "" Then
190             If currentC = "" Or currentC = "UNKNOWN" Then
200                 ws.Cells(r, 3).Value = MakeEnglishBranchName(branchHe)
210                 updatedC = updatedC + 1
220             End If
230         End If
240         If mainHe <> "" Then
250             If currentD = "" Or currentD = "UNKNOWN" Then
260                 ws.Cells(r, 4).Value = MakeEnglishBranchName(mainHe)
270                 updatedD = updatedD + 1
280             End If
290         End If
300     Next r
310     ws.Columns("A:D").AutoFit
320     MsgBox "Done" & vbCrLf & "Updated C: " & updatedC & vbCrLf & "Updated D: " & updatedD
330     Exit Sub
ERR_HANDLER:
340     MsgBox "UpdateSettingsEnglishCodes failed: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub


' ============================================================================
' HELPER: Build a stable English branch name from Hebrew or mixed text.
' No Hebrew literals are used here; Hebrew letters are handled by Unicode codes.
' ============================================================================
Private Function MakeEnglishBranchName(ByVal txt As String) As String
10      Dim i As Long
20      Dim code As Long
30      Dim part As String
40      Dim res As String
50      txt = Trim$(txt)
60      For i = 1 To Len(txt)
70          code = AscW(Mid$(txt, i, 1))
80          Select Case code
                Case 48 To 57, 65 To 90, 97 To 122
90                  part = ChrW$(code)
                Case 1488: part = "A"
                Case 1489: part = "B"
                Case 1490: part = "G"
                Case 1491: part = "D"
                Case 1492: part = "H"
                Case 1493: part = "V"
                Case 1494: part = "Z"
                Case 1495: part = "CH"
                Case 1496: part = "T"
                Case 1497: part = "Y"
                Case 1498, 1499: part = "K"
                Case 1500: part = "L"
                Case 1501, 1502: part = "M"
                Case 1503, 1504: part = "N"
                Case 1505: part = "S"
                Case 1506: part = "A"
                Case 1507, 1508: part = "P"
                Case 1509, 1510: part = "TZ"
                Case 1511: part = "K"
                Case 1512: part = "R"
                Case 1513: part = "SH"
                Case 1514: part = "T"
                Case Else
100                 part = "_"
110         End Select
120         res = res & part
130     Next i
140     res = UCase$(res)
150     Do While InStr(res, "__") > 0
160         res = Replace(res, "__", "_")
170     Loop
180     If Left$(res, 1) = "_" Then res = Mid$(res, 2)
190     If Right$(res, 1) = "_" Then res = Left$(res, Len(res) - 1)
200     If res = "" Then res = "BRANCH_" & CStr(Abs(SimpleTextHash(txt)))
210     MakeEnglishBranchName = res
End Function

Private Function SimpleTextHash(ByVal txt As String) As Long
10      Dim i As Long
20      Dim H As Long
30      H = 17
40      For i = 1 To Len(txt)
50          H = ((H * 31) + AscW(Mid$(txt, i, 1))) And &H7FFFFFFF
60      Next i
70      SimpleTextHash = H
End Function


' ============================================================================
' HELPER: Validate helper key exists
' ============================================================================
Private Sub ValidateHelperKey(ByVal dict As Object, ByVal key As String)
10      If Not dict.Exists(key) Then
20          Err.Raise vbObjectError + 2000, "ValidateHelperKey", "HELPER KEY NOT FOUND: " & key
30      End If
End Sub


' ============================================================================
' HELPER: Load checked fields from NIHUL
' ============================================================================
Private Sub LoadCheckedFields(ByVal ws As Worksheet, ByVal dictCol As Object, ByVal dictDisp As Object)
10      On Error GoTo ERR_HANDLER
20      Dim r As Long
30      Dim lastRow As Long
40      Dim fName As String
50      Dim fCol As String
60      Dim fCheck As String
70      Dim fKey As String
80      lastRow = ws.Cells(ws.Rows.Count, COL_FIELD_NAME_HE).End(xlUp).Row
90      For r = MANAGEMENT_START_ROW To lastRow
100         fName = Trim$(CStr(ws.Cells(r, COL_FIELD_NAME_HE).Value2))
110         fCol = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_COLUMN).Value2)))
120         fCheck = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_CHECKING).Value2)))
130         fKey = Trim$(CStr(ws.Cells(r, COL_FIELD_KEY).Value2))
140         If fKey <> "" And fCol <> "" And fCheck = "CHECK" Then
150             dictCol(fKey) = ColumnLetterToNumber(fCol)
160             dictDisp(fKey) = fName
170         End If
180     Next r
190     Exit Sub
ERR_HANDLER:
200     Err.Raise Err.Number, "LoadCheckedFields:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build arrays from dictionaries
' ============================================================================
Private Sub BuildArrays(ByVal dictCol As Object, ByVal dictDisp As Object, ByRef keys() As String, ByRef cols() As Long, ByRef disp() As String, ByRef cnt As Long)
10      On Error GoTo ERR_HANDLER
20      cnt = dictCol.Count
30      If cnt = 0 Then Exit Sub
40      ReDim keys(1 To cnt)
50      ReDim cols(1 To cnt)
60      ReDim disp(1 To cnt)
70      Dim i As Long
80      Dim k As Variant
90      i = 0
100     For Each k In dictCol.keys
110         i = i + 1
120         keys(i) = CStr(k)
130         cols(i) = CLng(dictCol(k))
140         disp(i) = CStr(dictDisp(k))
150     Next k
160     Exit Sub
ERR_HANDLER:
170     Err.Raise Err.Number, "BuildArrays:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Get string parameter from NIHUL
' ============================================================================
Private Function GetStringParameter(ByVal ws As Worksheet, ByVal paramName As String) As String
10      On Error GoTo ERR_HANDLER
20      Dim r As Long
30      Dim lastRow As Long
40      Dim nm As String
50      lastRow = ws.Cells(ws.Rows.Count, COL_PARAM_NAME).End(xlUp).Row
60      For r = 1 To lastRow
70          nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
80          If nm = UCase$(paramName) Then
90              GetStringParameter = Trim$(CStr(ws.Cells(r, COL_PARAM_VALUE).Value2))
100             Exit Function
110         End If
120     Next r
130     GetStringParameter = ""
140     Exit Function
ERR_HANDLER:
150     GetStringParameter = ""
End Function


' ============================================================================
' HELPER: Get numeric parameter from NIHUL
' ============================================================================
Private Function GetNumericParameter(ByVal ws As Worksheet, ByVal paramName As String) As Double
10      On Error GoTo ERR_HANDLER
20      Dim r As Long
30      Dim lastRow As Long
40      Dim nm As String
50      Dim v As Variant
60      Dim n As Double
70      lastRow = ws.Cells(ws.Rows.Count, COL_PARAM_NAME).End(xlUp).Row
80      For r = 1 To lastRow
90          nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
100         If nm = UCase$(paramName) Then
110             v = ws.Cells(r, COL_PARAM_VALUE).Value2
120             If TryParseVariantNumber(v, n) Then
130                 GetNumericParameter = n
140             Else
150                 Err.Raise vbObjectError + 3000, "GetNumericParameter", "PARAMETER NOT NUMERIC: " & paramName
160             End If
170             Exit Function
180         End If
190     Next r
200     Err.Raise vbObjectError + 3001, "GetNumericParameter", "PARAMETER NOT FOUND: " & paramName
ERR_HANDLER:
210     Err.Raise Err.Number, "GetNumericParameter:" & Erl, Err.Description
End Function


' ============================================================================
' HELPER: Column letter to number
' ============================================================================
Private Function ColumnLetterToNumber(ByVal col As String) As Long
10      On Error GoTo ERR_HANDLER
20      Dim i As Long
30      Dim ch As String
40      For i = 1 To Len(col)
50          ch = Mid$(col, i, 1)
60          If ch < "A" Or ch > "Z" Then
70              Err.Raise vbObjectError + 4000, "ColumnLetterToNumber", "INVALID COLUMN LETTER: " & col
80          End If
90          ColumnLetterToNumber = ColumnLetterToNumber * 26 + (Asc(ch) - 64)
100     Next i
110     Exit Function
ERR_HANDLER:
120     Err.Raise Err.Number, "ColumnLetterToNumber:" & Erl, Err.Description
End Function


' ============================================================================
' HELPER: Delete sheets
' ============================================================================
Private Sub DeleteReviewSheetIfExists()
10      On Error Resume Next
20      Application.DisplayAlerts = False
30      If SheetExists(REVIEW_SHEET_NAME()) Then
40          ThisWorkbook.Worksheets(REVIEW_SHEET_NAME()).Delete
50      End If
60      Application.DisplayAlerts = True
End Sub

Private Sub DeleteSheetIfExists(ByVal sName As String)
10      On Error Resume Next
20      Application.DisplayAlerts = False
30      If SheetExists(sName) Then
40          ThisWorkbook.Worksheets(sName).Delete
50      End If
60      Application.DisplayAlerts = True
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
10      On Error GoTo NOT_FOUND
20      Dim ws As Worksheet
30      Set ws = ThisWorkbook.Worksheets(sheetName)
40      SheetExists = True
50      Exit Function
NOT_FOUND:
60      SheetExists = False
End Function


' ============================================================================
' HELPER: Check if row should be ignored
' ============================================================================
Private Function IsIgnorableRow(ByVal ws As Worksheet, ByVal r As Long, ByRef keys() As String, ByRef cols() As Long, ByVal cnt As Long, ByVal dictFieldCol As Object) As Boolean
10      Dim allBlank As Boolean
20      Dim i As Long
30      allBlank = True
40      For i = 1 To cnt
50          If Not IsBlankValue(ws.Cells(r, cols(i)).Value2) Then
60              allBlank = False
70              Exit For
80          End If
90      Next i
100     IsIgnorableRow = allBlank
End Function


' ============================================================================
' HELPER: Check blank value
' ============================================================================
Private Function IsBlankValue(ByVal v As Variant) As Boolean
10      If IsEmpty(v) Then
20          IsBlankValue = True
30      ElseIf IsNull(v) Then
40          IsBlankValue = True
50      ElseIf VarType(v) = vbString Then
60          IsBlankValue = (Trim$(CStr(v)) = "")
70      Else
80          IsBlankValue = False
90      End If
End Function


' ============================================================================
' HELPER: Try parse variant to number
' ============================================================================
Private Function TryParseVariantNumber(ByVal v As Variant, ByRef result As Double) As Boolean
10      On Error GoTo FAIL
20      If IsNumeric(v) Then
30          result = CDbl(v)
40          TryParseVariantNumber = True
50      Else
60          TryParseVariantNumber = False
70      End If
80      Exit Function
FAIL:
90      TryParseVariantNumber = False
End Function


' ============================================================================
' HELPER: Add reason code
' ============================================================================
Private Function AddReason(ByVal existing As String, ByVal newReason As String) As String
10      If existing = "" Then
20          AddReason = newReason
30      Else
40          AddReason = existing & ", " & newReason
50      End If
End Function


' ============================================================================
' HELPER: Translate reason codes to Hebrew
' ============================================================================
Private Function TranslateReason(ByVal reasonCode As String, ByVal dictHelper As Object) As String
10      On Error Resume Next
20      Dim parts() As String
30      Dim i As Long
40      Dim translated As String
50      Dim part As String
60      parts = Split(reasonCode, ", ")
70      For i = 0 To UBound(parts)
80          part = Trim$(parts(i))
90          If dictHelper.Exists(part) Then
100             part = dictHelper(part)
110         End If
120         If translated = "" Then
130             translated = part
140         Else
150             translated = translated & ", " & part
160         End If
170     Next i
180     TranslateReason = translated
End Function




' ============================================================================
' HELPER: Return Hebrew month name for month number 1-12
' ============================================================================
Private Function HebrewMonthName(ByVal m As Long) As String
10      If m = 1 Then
20          HebrewMonthName = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
30      ElseIf m = 2 Then
40          HebrewMonthName = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
50      ElseIf m = 3 Then
60          HebrewMonthName = ChrW(1502) & ChrW(1512) & ChrW(1509)
70      ElseIf m = 4 Then
80          HebrewMonthName = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
90      ElseIf m = 5 Then
100         HebrewMonthName = ChrW(1502) & ChrW(1488) & ChrW(1497)
110     ElseIf m = 6 Then
120         HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
130     ElseIf m = 7 Then
140         HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
150     ElseIf m = 8 Then
160         HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
170     ElseIf m = 9 Then
180         HebrewMonthName = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
190     ElseIf m = 10 Then
200         HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
210     ElseIf m = 11 Then
220         HebrewMonthName = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
230     ElseIf m = 12 Then
240         HebrewMonthName = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
250     Else
260         HebrewMonthName = CStr(m)
270     End If
End Function


' ============================================================================
' SETUP: Create labels, dropdowns, period lists, and buttons on Main sheet
' Run this once to set up the control panel
' ============================================================================
Public Sub SetupMainSheet()

        Dim dateTypeList As String

10      Dim wsMain As Worksheet
20      Dim wsMgmt As Worksheet
30      Dim shp As Shape
        Dim s As Shape
        Dim blueClr As Long
        Dim lblShnBasis As String
        Dim lblShnShotef As String
        Dim lblTkufa As String
        Dim lblPirutTkufa As String
        Dim lblSugTaarih As String
        Dim periodTypeList As String
        Dim halfList As String
        Dim quarterList As String
        Dim monthList As String
        Dim periodTypeFormula As String
        Dim dvC4 As String

40      On Error GoTo ERR_HANDLER

        ' Try to find Main sheet by current Hebrew name or old English name
50      On Error Resume Next
51      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
52      If wsMain Is Nothing Then Set wsMain = ThisWorkbook.Worksheets("Main")
53      On Error GoTo ERR_HANDLER
54      If wsMain Is Nothing Then Err.Raise vbObjectError + 9001, "SetupMainSheet", "Cannot find Main sheet"
        ' Rename to Hebrew if still English
55      If wsMain.Name <> CONTROL_SHEET_NAME() Then wsMain.Name = CONTROL_SHEET_NAME()
        ' Try to find NIHUL/hagdarot sheet by current Hebrew name or old English name
56      On Error Resume Next
57      Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
58      If wsMgmt Is Nothing Then Set wsMgmt = ThisWorkbook.Worksheets("NIHUL")
59      On Error GoTo ERR_HANDLER
60      If wsMgmt Is Nothing Then Err.Raise vbObjectError + 9002, "SetupMainSheet", "Cannot find NIHUL/hagdarot sheet"
        ' Rename to Hebrew if still English
62      If wsMgmt.Name <> MANAGEMENT_SHEET_NAME() Then wsMgmt.Name = MANAGEMENT_SHEET_NAME()
70      blueClr = RGB(0, 70, 140)

        ' ---- Hebrew label strings ----
        ' "shnat basis" = year of reference
80      lblShnBasis = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505)
        ' "shna shoteft" = current year
90      lblShnShotef = ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1513) & ChrW(1493) & ChrW(1496) & ChrW(1508) & ChrW(1514)
        ' "tkufa" = period
100     lblTkufa = ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
        ' "pirot tkufa" = period detail
110     lblPirutTkufa = ChrW(1508) & ChrW(1497) & ChrW(1512) & ChrW(1493) & ChrW(1496) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
        ' "sug taarih" = date type
120     lblSugTaarih = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1514) & ChrW(1488) & ChrW(1512) & ChrW(1497) & ChrW(1498)

        ' ---- Write labels in column A ----
        ' Row 2: shnat basis (B2=ref year)
        ' Row 3: shna shoteft (B3=current year)
        ' Row 4: tkufa (B4=period type), pirot tkufa label in D4 (C4=detail)
        ' Row 5: sug taarih (B5=date type)
130     wsMain.Range("A2").Value = lblShnBasis
140     wsMain.Range("A3").Value = lblShnShotef
150     wsMain.Range("A4").Value = lblTkufa
155     wsMain.Range("D4").Value = lblPirutTkufa
160     wsMain.Range("A5").Value = lblSugTaarih

        ' Format labels
180     wsMain.Range("A2:A5").Font.Bold = True
185     wsMain.Range("D4").Font.Bold = True
190     wsMain.Range("A2:A5").Font.Size = 11
195     wsMain.Range("D4").Font.Size = 11
200     wsMain.Range("A2:A5").Font.Color = blueClr
205     wsMain.Range("D4").Font.Color = blueClr
210     wsMain.Range("A2:A5").HorizontalAlignment = xlRight
215     wsMain.Range("D4").HorizontalAlignment = xlRight
220     wsMain.Columns("A").ColumnWidth = 16
230     wsMain.Columns("B").ColumnWidth = 18
235     wsMain.Columns("C").ColumnWidth = 4
237     wsMain.Columns("D").ColumnWidth = 16
240     wsMain.Columns("E").ColumnWidth = 18

        ' ---- Write period lookup lists on NIHUL column Q ----
        ' Q1 header: "period_type"
250     wsMgmt.Cells(1, 17).Value = "PERIOD_TYPE"
        ' Q2-Q5: the four period types
        ' shnatit = yearly
260     wsMgmt.Cells(2, 17).Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
        ' chatzi shnati = half yearly
270     wsMgmt.Cells(3, 17).Value = ChrW(1495) & ChrW(1510) & ChrW(1497) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
        ' riv'oni = quarterly
280     wsMgmt.Cells(4, 17).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497)
        ' chodshi = monthly
290     wsMgmt.Cells(5, 17).Value = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497)

        ' R1 header: "HALF_YEAR"
300     wsMgmt.Cells(1, 18).Value = "HALF_YEAR"
        ' R2: machatzit rishona = first half
310     wsMgmt.Cells(2, 18).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504) & ChrW(1492)
        ' R3: machatzit shniya = second half
320     wsMgmt.Cells(3, 18).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1492)

        ' Define named ranges on NIHUL for dropdown sources
        ' Period types: NIHUL!Q2:Q5
330     On Error Resume Next
340     ThisWorkbook.Names("lst_period_type").Delete
350     ThisWorkbook.Names("lst_half_year").Delete
360     ThisWorkbook.Names("lst_quarter").Delete
370     ThisWorkbook.Names("lst_month").Delete
380     On Error GoTo ERR_HANDLER

390     ThisWorkbook.Names.Add Name:="lst_period_type", RefersTo:="=" & MANAGEMENT_SHEET_NAME() & "!$Q$2:$Q$5"
400     ThisWorkbook.Names.Add Name:="lst_half_year", RefersTo:="=" & MANAGEMENT_SHEET_NAME() & "!$R$2:$R$3"

        ' Quarter list in NIHUL column R rows 5-8
        ' riv'on rishon = Q1
410     wsMgmt.Cells(5, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503)
        ' riv'on sheni = Q2
420     wsMgmt.Cells(6, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497)
        ' riv'on shlishi = Q3
430     wsMgmt.Cells(7, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497)
        ' riv'on revi'i = Q4
440     wsMgmt.Cells(8, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497)
450     ThisWorkbook.Names.Add Name:="lst_quarter", RefersTo:="=" & MANAGEMENT_SHEET_NAME() & "!$R$5:$R$8"

        ' Month list in NIHUL column R rows 10-21 (Hebrew month names)
460     wsMgmt.Cells(10, 18).Value = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
470     wsMgmt.Cells(11, 18).Value = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
480     wsMgmt.Cells(12, 18).Value = ChrW(1502) & ChrW(1512) & ChrW(1509)
490     wsMgmt.Cells(13, 18).Value = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
500     wsMgmt.Cells(14, 18).Value = ChrW(1502) & ChrW(1488) & ChrW(1497)
510     wsMgmt.Cells(15, 18).Value = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
520     wsMgmt.Cells(16, 18).Value = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
530     wsMgmt.Cells(17, 18).Value = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
540     wsMgmt.Cells(18, 18).Value = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
550     wsMgmt.Cells(19, 18).Value = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
560     wsMgmt.Cells(20, 18).Value = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
570     wsMgmt.Cells(21, 18).Value = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
580     ThisWorkbook.Names.Add Name:="lst_month", RefersTo:="=" & MANAGEMENT_SHEET_NAME() & "!$R$10:$R$21"

        ' ---- B4 dropdown: period type ----
590     wsMain.Range("B4").Validation.Delete
600     wsMain.Range("B4").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_period_type"

        ' ---- B5 dropdown: date type ----
        ' bordereu / thilat bituah
610     wsMain.Range("B5").Validation.Delete

620     dateTypeList = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493) & "," & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1514) & " " & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495)
630     wsMain.Range("B5").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=dateTypeList

        ' ---- C4 dependent dropdown: uses INDIRECT based on B4 value ----
        ' We use a formula approach: map B4 value to named range
        ' INDIRECT formula: =IF(B4=shnatit,"",IF(B4=chatzi,lst_half_year,IF(B4=riv'oni,lst_quarter,IF(B4=chodshi,lst_month,""))))
        ' Since INDIRECT with dynamic named ranges is complex, we use Worksheet_Change event instead
        ' For now, set C4 validation to allow any list - it will be updated by the event macro
640     wsMain.Range("E4").Validation.Delete

        ' ---- Remove old buttons ----
650     On Error Resume Next
        For Each s In wsMain.Shapes
660         If s.Name = "btnBuildReview" Or s.Name = "btnApplyCorrections" Or s.Name = "btnBuildPresentation" Then s.Delete
670     Next s
680     On Error GoTo ERR_HANDLER

        ' ---- Button 1: BuildReview ----
        ' "1 - bdikat netunim" = Data Review
690     Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 50, 180, 220, 45)
700     shp.Name = "btnBuildReview"
710     shp.Fill.ForeColor.RGB = blueClr
720     shp.TextFrame2.TextRange.Text = "1 - " & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
730     shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
740     shp.TextFrame2.TextRange.Font.Size = 14
750     shp.TextFrame2.TextRange.Font.Bold = msoTrue
760     shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
770     shp.OnAction = "BuildReview"

        ' ---- Button 2: ApplyCorrectionsAndBuildReports ----
        ' "2 - yisshum vedochot" = Apply & Reports
780     Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 50, 240, 220, 45)
790     shp.Name = "btnApplyCorrections"
800     shp.Fill.ForeColor.RGB = RGB(0, 120, 60)
810     shp.TextFrame2.TextRange.Text = "2 - " & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(34) & ChrW(1495) & ChrW(1493) & ChrW(1514)
820     shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
830     shp.TextFrame2.TextRange.Font.Size = 14
840     shp.TextFrame2.TextRange.Font.Bold = msoTrue
850     shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
860     shp.OnAction = "ApplyCorrectionsAndBuildReports"

        ' ---- Button 3: BuildPresentation ----
        ' "3 - yitzur matzget" = Create Presentation
870     On Error Resume Next
        For Each s In wsMain.Shapes
872         If s.Name = "btnBuildPresentation" Then s.Delete
874     Next s
876     On Error GoTo ERR_HANDLER
878     Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, 50, 300, 220, 45)
880     shp.Name = "btnBuildPresentation"
882     shp.Fill.ForeColor.RGB = RGB(160, 80, 0)
884     shp.TextFrame2.TextRange.Text = "3 - " & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
886     shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
888     shp.TextFrame2.TextRange.Font.Size = 14
890     shp.TextFrame2.TextRange.Font.Bold = msoTrue
892     shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
894     shp.OnAction = "BuildPresentation"

        ' ---- Store Hebrew message texts in column S of hagdarot ----
        ' S1=header, S2=done+found, S3=issues, S4=done ok, S5=line, S6=error
        ' S7=step, S8=confirm title, S9=btn2 confirm, S10=setup err, S11=dropdown err, S12=setup ok
        ' S13-S16=confirmation lines 1-4
865     wsMgmt.Cells(1, 19).Value = ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
866     wsMgmt.Cells(2, 19).Value = ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & "-" & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " "
867     wsMgmt.Cells(3, 19).Value = " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1493) & ChrW(1514)
868     wsMgmt.Cells(4, 19).Value = ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492)
869     wsMgmt.Cells(5, 19).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1492) & " "
        wsMgmt.Cells(6, 19).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " "
        wsMgmt.Cells(7, 19).Value = ChrW(1513) & ChrW(1500) & ChrW(1489) & ":" & " "
        wsMgmt.Cells(8, 19).Value = ChrW(1488) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1508) & ChrW(1504) & ChrW(1497) & " " & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491)
        wsMgmt.Cells(9, 19).Value = ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1501) & " " & ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1512) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & "?"
        wsMgmt.Cells(10, 19).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & ":" & " "
        wsMgmt.Cells(11, 19).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1493) & ChrW(1503) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492) & ":" & " "
        wsMgmt.Cells(12, 19).Value = ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & " " & ChrW(1492) & ChrW(1493) & ChrW(1511) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!"
        wsMgmt.Cells(13, 19).Value = "1" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1492) & ChrW(1506) & ChrW(1500) & ChrW(1497) & ChrW(1514) & " " & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1508) & ChrW(1492) & " " & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & ChrW(1497) & "?"
        wsMgmt.Cells(14, 19).Value = "2" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1514) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1488) & ChrW(1495) & ChrW(1512) & ChrW(1493) & ChrW(1503) & " " & "(" & ChrW(1492) & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1512) & ChrW(1514) & " " & ChrW(1494) & ChrW(1493) & " " & ChrW(1497) & ChrW(1502) & ChrW(1495) & ChrW(1511) & ChrW(1493) & " " & ChrW(1489) & ChrW(1514) & ChrW(1492) & ChrW(1500) & ChrW(1497) & ChrW(1498) & " " & ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & ")"
        wsMgmt.Cells(15, 19).Value = "3" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(1514) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1508) & ChrW(1512) & ChrW(1496) & ChrW(1497) & " " & ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1504) & ChrW(1491) & ChrW(1512) & ChrW(1513) & " " & ChrW(1489) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & "?"
        wsMgmt.Cells(16, 19).Value = "4" & "." & " " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1506) & ChrW(1500) & " " & "O" & "K" & " " & ChrW(1500) & ChrW(1492) & ChrW(1502) & ChrW(1513) & ChrW(1498)
        ' S17 = "don't show this message again?" text
        wsMgmt.Cells(17, 19).Value = ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1492) & " " & ChrW(1494) & ChrW(1493) & " " & ChrW(1513) & ChrW(1493) & ChrW(1489) & "?"
        ' S20 = flag: "1" means don't show confirmation again (empty = show)
        ' Don't overwrite S20 if already set

        ' ---- Set Error_Email parameter if not exists ----
895     Dim paramLastRow As Long
896     paramLastRow = wsMgmt.Cells(wsMgmt.Rows.Count, COL_PARAM_NAME).End(xlUp).Row
897     Dim foundEmail As Boolean
898     foundEmail = False
899     Dim pr As Long
        For pr = 1 To paramLastRow
            If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "ERROR_EMAIL" Then foundEmail = True: Exit For
        Next pr
        If Not foundEmail Then
            wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "ERROR_EMAIL"
            wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "zvi@gorentech.co.il"
        End If

        ' ---- Set default values if empty ----
900     If IsEmpty(wsMain.Range("B4").Value) Or wsMain.Range("B4").Value = "" Then
902         wsMain.Range("B4").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
904     End If

        ' Success message - read from S12
910     MsgBoxU wsMgmt.Cells(12, 19).Value, vbInformation

920     Exit Sub

ERR_HANDLER:
930     MsgBoxU wsMgmt.Cells(10, 19).Value & Err.Description, vbCritical

End Sub


' ============================================================================
' EVENT HANDLER: Goes in the Main sheet module (Sheet code)
' Call UpdatePeriodDropdown from Worksheet_Change when B4 changes
' This sub updates E4 validation based on B4 period type selection
' ============================================================================
Public Sub UpdatePeriodDropdown()

10      Dim wsMain As Worksheet
        Dim periodType As String
        Dim listName As String

20      On Error GoTo ERR_HANDLER

30      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
40      periodType = Trim$(CStr(wsMain.Range("B4").Value2))

        ' Clear E4
50      wsMain.Range("E4").Value = ""
60      On Error Resume Next
70      wsMain.Range("E4").Validation.Delete
80      On Error GoTo ERR_HANDLER

        ' Check chatzi shnati BEFORE shnatit (shnatit is substring of chatzi shnati)
        ' "chatzi shnati" = half yearly
90      If InStr(1, periodType, ChrW(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) > 0 Then
100         listName = "lst_half_year"
        ' "riv'oni" = quarterly
110     ElseIf InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
120         listName = "lst_quarter"
        ' "chodshi" = monthly
130     ElseIf InStr(1, periodType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
140         listName = "lst_month"
        ' "shnatit" = yearly -> no second dropdown needed
150     ElseIf InStr(1, periodType, ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497), vbTextCompare) > 0 Then
160         Exit Sub
170     Else
180         Exit Sub
190     End If

200     wsMain.Range("E4").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & listName

210     Exit Sub

ERR_HANDLER:
220     MsgBoxU ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME()).Cells(11, 19).Value & Err.Description, vbCritical

End Sub

' ============================================================================
' MACRO 3: BuildPresentation
' Creates a PPTX management presentation from comparison sheets
' v5: Landscape, split charts/tables, page numbers, insured column
' Phase 1: Export charts as images (Excel only)
' Phase 2: Build PowerPoint slides from images + data
' ============================================================================
Public Sub BuildPresentation()

10      On Error GoTo ERR_HANDLER

        Dim wsMain As Worksheet
20      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())

      
    
        Dim periodType As String
        Dim periodDetail As String
        Dim periodDesc As String
        Dim yearVal As String
        Dim refYear As String



40      yearVal = Trim$(CStr(wsMain.Range("B3").Value2))
50      refYear = Trim$(CStr(wsMain.Range("B2").Value2))
60      periodType = Trim$(CStr(wsMain.Range("B4").Value2))
70      periodDetail = Trim$(CStr(wsMain.Range("E4").Value2))
80      periodDesc = periodType
90      If periodDetail <> "" Then periodDesc = periodDesc & " " & periodDetail

        ' Validate that comparison sheets exist
100     If Not SheetExists(SHEET_COMPANIES()) Then
110         MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1511) & ChrW(1493) & ChrW(1491) & ChrW(1501) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 2", vbCritical
120         Exit Sub
130     End If

        ' Show processing message
140     With wsMain.Range("B8")
150         .Value = ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & "..."
160         .Font.Size = 16
170         .Font.Bold = True
180         .Font.Color = RGB(200, 0, 0)
190         .Interior.Color = RGB(255, 255, 200)
200     End With
210     Application.ScreenUpdating = True
220     DoEvents

        ' ================================================================
        ' PHASE 1: Create all chart images in Excel (NO PowerPoint yet)
        ' ================================================================
        Dim tmpPath As String
230     tmpPath = Environ$("TEMP") & "\"

        Dim imgTotal As String
240     imgTotal = tmpPath & "levav_total.gif"
250     ExportTotalChart imgTotal, yearVal, refYear

        ' Build list of sheets to process
        Dim sheetList(1 To 8) As String
        Dim titleList(1 To 8) As String
        Dim sheetCount As Long
260     sheetCount = 0

270     If SheetExists(SHEET_MONTHS()) Then
280         sheetCount = sheetCount + 1
290         sheetList(sheetCount) = SHEET_MONTHS()
292         titleList(sheetCount) = TITLE_MONTHS()
300     End If
310     If SheetExists(SHEET_COMPANIES()) Then
320         sheetCount = sheetCount + 1
330         sheetList(sheetCount) = SHEET_COMPANIES()
332         titleList(sheetCount) = TITLE_COMPANIES()
340     End If
350     If SheetExists(SHEET_MAINBRANCH()) Then
360         sheetCount = sheetCount + 1
370         sheetList(sheetCount) = SHEET_MAINBRANCH()
372         titleList(sheetCount) = TITLE_MAINBRANCH()
380     End If
390     If SheetExists(SHEET_TELLERS()) Then
400         sheetCount = sheetCount + 1
410         sheetList(sheetCount) = SHEET_TELLERS()
412         titleList(sheetCount) = TITLE_TELLERS()
420     End If
430     If SheetExists(SHEET_AGENTS()) Then
440         sheetCount = sheetCount + 1
450         sheetList(sheetCount) = SHEET_AGENTS()
452         titleList(sheetCount) = TITLE_AGENTS()
460     End If
461     If SheetExists(SHEET_AGENTS_NO_LEVAV()) Then
462         sheetCount = sheetCount + 1
463         sheetList(sheetCount) = SHEET_AGENTS_NO_LEVAV()
464         titleList(sheetCount) = TITLE_AGENTS_NO_LEVAV()
465     End If
466     ' Drachim is not added to the regular sheet list.
467     ' It is shown later as one total-row slide only; helper sheet stays hidden.

        ' Export 2 charts per sheet
        Dim si As Long
        Dim imgFiles() As String
470     ReDim imgFiles(1 To sheetCount * 2)
        Dim exportOK() As Boolean
480     ReDim exportOK(1 To sheetCount)
490     For si = 1 To sheetCount
500         imgFiles(si * 2 - 1) = tmpPath & "levav_prem_" & si & ".gif"
510         imgFiles(si * 2) = tmpPath & "levav_comm_" & si & ".gif"
520         On Error Resume Next
530         ExportCompCharts sheetList(si), imgFiles(si * 2 - 1), imgFiles(si * 2), yearVal, refYear
540         If Err.Number = 0 Then
550             exportOK(si) = True
560         Else
570             exportOK(si) = False
572             Err.Clear
580         End If
590         On Error GoTo ERR_HANDLER
600     Next si

605     DoEvents

        ' ================================================================
        ' PHASE 2: Open PowerPoint and build slides
        ' ================================================================
        Dim ppApp As Object
        Dim ppPres As Object
        Dim ppSlide As Object
610     Set ppApp = CreateObject("PowerPoint.Application")
615     ppApp.Visible = True
620     Set ppPres = ppApp.Presentations.Add

        ' Set LANDSCAPE slide size (13.33" x 7.5")
625     ppPres.PageSetup.SlideWidth = 960
630     ppPres.PageSetup.SlideHeight = 540

        Dim slideIdx As Long
        Dim slideW As Single
        Dim slideH As Single
635     slideIdx = 0
640     slideW = 960
645     slideH = 540

        ' Slide title names (Hebrew)
        Dim titleNames(1 To 6) As String
650     titleNames(1) = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
655     titleNames(2) = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
660     titleNames(3) = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
665     titleNames(4) = ChrW(1496) & ChrW(1500) & ChrW(1512) & ChrW(1497) & ChrW(1493) & ChrW(1514)
670     titleNames(5) = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)

        ' SLIDE 1: Title
675     slideIdx = slideIdx + 1
680     Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
685     BuildTitleSlide ppSlide, yearVal, refYear, periodDesc, slideW, slideH

        ' SLIDE 2: Total Summary chart
690     slideIdx = slideIdx + 1
695     Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
700     BuildTotalSlideFromImage ppSlide, imgTotal, yearVal, refYear, slideW

        ' For each comparison sheet: 3 slides (prem chart, comm chart, table)
730     For si = 1 To sheetCount
740         If exportOK(si) Then
                ' Slide: Premium chart
750             slideIdx = slideIdx + 1
760             Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
770             BuildChartSlide ppSlide, imgFiles(si * 2 - 1), ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleList(si), yearVal, refYear, slideW
                ' Slide: Commission chart
780             slideIdx = slideIdx + 1
790             Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
800             BuildChartSlide ppSlide, imgFiles(si * 2), ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleList(si), yearVal, refYear, slideW
810         End If
            ' Slide: Data table (always, even if charts failed)
820         slideIdx = slideIdx + 1
830         Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
840         BuildTableSlide ppSlide, sheetList(si), titleList(si), yearVal, refYear, slideW, slideH
850     Next si

        ' Drachim: one slide only, based on total row of hidden helper sheet
852     If SheetExists(SHEET_DRACHIM()) Then
854         slideIdx = slideIdx + 1
856         Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
858         BuildTotalRowTableSlide ppSlide, SHEET_DRACHIM(), TITLE_DRACHIM(), yearVal, refYear, slideW, slideH
859     End If

        ' Add page numbers to all slides
        Dim pg As Long
860     For pg = 1 To ppPres.Slides.Count
870         AddPageNumber ppPres.Slides(pg), pg, ppPres.Slides.Count, slideW, slideH
880     Next pg

        ' Save presentation
        Dim savePath As String
890     savePath = ThisWorkbook.Path & "\" & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1492) & ChrW(1504) & ChrW(1492) & ChrW(1500) & ChrW(1492) & " " & yearVal & ".pptx"
900     ppPres.SaveAs savePath
        ' Leave presentation open for user to view
910     Set ppPres = Nothing
920     Set ppApp = Nothing

        ' Cleanup temp images
950     On Error Resume Next
960     Kill imgTotal
970     For si = 1 To sheetCount * 2
980         Kill imgFiles(si)
990     Next si
1000    On Error GoTo ERR_HANDLER

        ' Clear processing message
1010    With wsMain.Range("B8")
1020        .Value = ""
1030        .Interior.ColorIndex = xlNone
1040    End With

        ' Success message
1050    MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!" & vbCrLf & savePath, vbInformation

1060    Exit Sub

ERR_HANDLER:
        Dim errLine As Long
        Dim errDesc As String
        Dim errNum As Long
1070    errLine = Erl
1072    errDesc = Err.Description
1074    errNum = Err.Number
1076    On Error Resume Next
        With wsMain.Range("B8")
            .Value = ""
            .Interior.ColorIndex = xlNone
        End With
        If Not ppPres Is Nothing Then ppPres.Close
        If Not ppApp Is Nothing Then ppApp.Quit
        Kill imgTotal
        Dim ei As Long
        For ei = 1 To sheetCount * 2
            Kill imgFiles(ei)
        Next ei
1080    MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & ":" & vbCrLf & "Line: " & errLine & vbCrLf & "Err #" & errNum & ": " & errDesc, vbCritical

End Sub


' ============================================================================
' HELPER: Export Total Summary chart to image file
' ============================================================================
Private Sub ExportTotalChart(ByVal imgPath As String, ByVal yearVal As String, ByVal refYear As String)

10      On Error GoTo ERR_HANDLER

        Dim ws As Worksheet
        Dim lastRow As Long
20      Set ws = ThisWorkbook.Worksheets(SHEET_MONTHS())
30      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        Dim sumPR As Double
        Dim sumPC As Double
        Dim sumCR As Double
        Dim sumCC As Double
40      sumPR = CDbl(ws.Cells(lastRow, 2).Value2)
50      sumPC = CDbl(ws.Cells(lastRow, 3).Value2)
60      sumCR = CDbl(ws.Cells(lastRow, 14).Value2)
70      sumCC = CDbl(ws.Cells(lastRow, 15).Value2)

        Dim tmpWs As Worksheet
        Dim co As Object
        Dim xlCht As Object
80      Application.ScreenUpdating = False
90      Set tmpWs = ThisWorkbook.Worksheets.Add

100     tmpWs.Cells(1, 1).Value = ""
110     tmpWs.Cells(1, 2).Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & refYear
120     tmpWs.Cells(1, 3).Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & yearVal
130     tmpWs.Cells(1, 4).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & refYear
140     tmpWs.Cells(1, 5).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & yearVal
150     tmpWs.Cells(2, 1).Value = ChrW(1505) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1500)
160     tmpWs.Cells(2, 2).Value = sumPR
170     tmpWs.Cells(2, 3).Value = sumPC
180     tmpWs.Cells(2, 4).Value = sumCR
190     tmpWs.Cells(2, 5).Value = sumCC

200     Set co = tmpWs.ChartObjects.Add(10, 10, 600, 400)
210     Set xlCht = co.Chart
220     xlCht.ChartType = 51
230     xlCht.SetSourceData tmpWs.Range("A1:E2")
240     xlCht.HasTitle = False
250     xlCht.HasLegend = True

260     xlCht.Export imgPath

270     Application.DisplayAlerts = False
280     tmpWs.Delete
290     Application.DisplayAlerts = True
300     Application.ScreenUpdating = True

310     Exit Sub
ERR_HANDLER:
320     On Error Resume Next
        Application.DisplayAlerts = False
        If Not tmpWs Is Nothing Then tmpWs.Delete
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
330     Err.Raise Err.Number, "ExportTotalChart:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Export Comparison charts (premiums + commissions) to 2 image files
' ============================================================================
Private Sub ExportCompCharts(ByVal sheetName As String, ByVal imgPrem As String, ByVal imgComm As String, ByVal yearVal As String, ByVal refYear As String)

10      On Error GoTo ERR_HANDLER

        Dim ws As Worksheet
        Dim lastRow As Long
        Dim dataRows As Long
        Dim r As Long
20      Set ws = ThisWorkbook.Worksheets(sheetName)
30      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
40      dataRows = lastRow - 2

        Dim arrNames() As String
        Dim arrPremR() As Double
        Dim arrPremC() As Double
        Dim arrCommR() As Double
        Dim arrCommC() As Double
        Dim nItems As Long
50      nItems = dataRows - 1
60      If nItems < 1 Then Exit Sub

70      ReDim arrNames(1 To nItems)
80      ReDim arrPremR(1 To nItems)
82      ReDim arrPremC(1 To nItems)
90      ReDim arrCommR(1 To nItems)
92      ReDim arrCommC(1 To nItems)

        Dim idx As Long
100     idx = 0
110     For r = 3 To lastRow - 1
120         idx = idx + 1
130         If idx > nItems Then Exit For
140         arrNames(idx) = Trim$(CStr(ws.Cells(r, 1).Value2))
150         arrPremR(idx) = CDbl(ws.Cells(r, 2).Value2)
160         arrPremC(idx) = CDbl(ws.Cells(r, 3).Value2)
170         arrCommR(idx) = CDbl(ws.Cells(r, 14).Value2)
180         arrCommC(idx) = CDbl(ws.Cells(r, 15).Value2)
190     Next r
200     nItems = idx

        Dim chartItems As Long
210     chartItems = nItems
220     If chartItems > 15 Then chartItems = 15

        Dim tmpWs As Worksheet
        Dim co As Object
        Dim xlCht As Object
        Dim ci As Long
230     Application.ScreenUpdating = False
240     Set tmpWs = ThisWorkbook.Worksheets.Add

        ' ---- Chart 1: Premiums ----
250     tmpWs.Cells(1, 1).Value = ""
260     tmpWs.Cells(1, 2).Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & refYear
270     tmpWs.Cells(1, 3).Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & yearVal
280     For ci = 1 To chartItems
290         tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
300         tmpWs.Cells(ci + 1, 2).Value = arrPremR(ci)
310         tmpWs.Cells(ci + 1, 3).Value = arrPremC(ci)
320     Next ci

330     Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
340     Set xlCht = co.Chart
350     xlCht.ChartType = 51
360     xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3))
370     xlCht.HasTitle = False
380     xlCht.HasLegend = True
382     On Error Resume Next
384     xlCht.Axes(2).MinimumScaleIsAuto = True
386     xlCht.Axes(2).MaximumScaleIsAuto = True
388     On Error GoTo ERR_HANDLER

390     xlCht.Export imgPrem

400     tmpWs.ChartObjects.Delete
410     tmpWs.Cells.Clear

        ' ---- Chart 2: Commissions ----
420     tmpWs.Cells(1, 1).Value = ""
430     tmpWs.Cells(1, 2).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & refYear
440     tmpWs.Cells(1, 3).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & yearVal
450     For ci = 1 To chartItems
460         tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
470         tmpWs.Cells(ci + 1, 2).Value = arrCommR(ci)
480         tmpWs.Cells(ci + 1, 3).Value = arrCommC(ci)
490     Next ci

500     Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
510     Set xlCht = co.Chart
520     xlCht.ChartType = 51
530     xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3))
540     xlCht.HasTitle = False
550     xlCht.HasLegend = True
552     On Error Resume Next
554     xlCht.Axes(2).MinimumScaleIsAuto = True
556     xlCht.Axes(2).MaximumScaleIsAuto = True
558     On Error GoTo ERR_HANDLER

560     xlCht.Export imgComm

570     Application.DisplayAlerts = False
580     tmpWs.Delete
590     Application.DisplayAlerts = True
600     Application.ScreenUpdating = True

610     Exit Sub
ERR_HANDLER:
620     On Error Resume Next
        Application.DisplayAlerts = False
        If Not tmpWs Is Nothing Then tmpWs.Delete
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
630     Err.Raise Err.Number, "ExportCompCharts(" & sheetName & "):" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build Title Slide (landscape)
' ============================================================================
Private Sub BuildTitleSlide(ByVal ppSlide As Object, ByVal yearVal As String, ByVal refYear As String, ByVal periodDesc As String, ByVal slideW As Single, ByVal slideH As Single)

10      On Error GoTo ERR_HANDLER

        Dim shp As Object

        ' Yellow/gold background
20      Set shp = ppSlide.Shapes.AddShape(1, 0, 0, slideW, slideH)
30      shp.Fill.ForeColor.RGB = RGB(240, 190, 50)
40      shp.Line.Visible = False

        ' White center rectangle
        Dim wL As Single
        Dim wT As Single
        Dim wW As Single
        Dim wH As Single
50      wL = 50
52      wT = 40
54      wW = slideW - 100
56      wH = slideH - 80
60      Set shp = ppSlide.Shapes.AddShape(1, wL, wT, wW, wH)
70      shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
80      shp.Line.Visible = False

        ' Title: "matzget hanhala"
90      Set shp = ppSlide.Shapes.AddTextbox(1, wL + 30, wT + 50, wW - 60, 80)
100     shp.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1492) & ChrW(1504) & ChrW(1492) & ChrW(1500) & ChrW(1492)
110     shp.TextFrame.TextRange.Font.Size = 40
120     shp.TextFrame.TextRange.Font.Bold = True
130     shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
140     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
150     shp.TextFrame.WordWrap = True

        ' Period
160     Set shp = ppSlide.Shapes.AddTextbox(1, wL + 30, wT + 160, wW - 60, 100)
170     shp.TextFrame.TextRange.Text = refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal & vbCrLf & "(" & periodDesc & ")"
180     shp.TextFrame.TextRange.Font.Size = 32
190     shp.TextFrame.TextRange.Font.Color.RGB = RGB(80, 80, 80)
200     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
210     shp.TextFrame.WordWrap = True

        ' Company name
220     Set shp = ppSlide.Shapes.AddTextbox(1, wL + 30, wT + wH - 100, wW - 60, 70)
230     shp.TextFrame.TextRange.Text = ChrW(1500) & ChrW(1489) & ChrW(1489) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495)
240     shp.TextFrame.TextRange.Font.Size = 34
250     shp.TextFrame.TextRange.Font.Bold = True
260     shp.TextFrame.TextRange.Font.Color.RGB = RGB(0, 130, 60)
270     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
280     shp.TextFrame.WordWrap = True

        ' Bottom gold line
290     Set shp = ppSlide.Shapes.AddShape(1, 50, slideH - 35, slideW - 100, 5)
300     shp.Fill.ForeColor.RGB = RGB(200, 160, 30)
310     shp.Line.Visible = False

320     Exit Sub
ERR_HANDLER:
330     Err.Raise Err.Number, "BuildTitleSlide:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build Total Summary Slide from pre-exported image
' ============================================================================
Private Sub BuildTotalSlideFromImage(ByVal ppSlide As Object, ByVal imgPath As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single)

10      On Error GoTo ERR_HANDLER

        Dim shp As Object

        ' Title textbox
20      Set shp = ppSlide.Shapes.AddTextbox(1, 20, 10, slideW - 40, 50)
30      shp.TextFrame.TextRange.Text = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1493) & ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
40      shp.TextFrame.TextRange.Font.Size = 24
50      shp.TextFrame.TextRange.Font.Bold = True
60      shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
70      shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
80      shp.TextFrame.WordWrap = True

        ' Insert chart image (landscape: wider)
90      ppSlide.Shapes.AddPicture imgPath, 0, 1, 80, 70, slideW - 160, 440

100     Exit Sub
ERR_HANDLER:
110     Err.Raise Err.Number, "BuildTotalSlideFromImage:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build a single chart slide (one chart image + title)
' ============================================================================
Private Sub BuildChartSlide(ByVal ppSlide As Object, ByVal imgPath As String, ByVal chartTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single)

10      On Error GoTo ERR_HANDLER

        Dim shp As Object

        ' Title
20      Set shp = ppSlide.Shapes.AddTextbox(1, 20, 10, slideW - 40, 50)
30      shp.TextFrame.TextRange.Text = chartTitle & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
40      shp.TextFrame.TextRange.Font.Size = 22
50      shp.TextFrame.TextRange.Font.Bold = True
60      shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
70      shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
80      shp.TextFrame.WordWrap = True

        ' Insert chart image
90      ppSlide.Shapes.AddPicture imgPath, 0, 1, 60, 65, slideW - 120, 450

100     Exit Sub
ERR_HANDLER:
110     Err.Raise Err.Number, "BuildChartSlide:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build a data table slide
' 13 cols: name | premRef | premCur | prem% | docsRef | docsCur | docs% |
'          insuredRef | insuredCur | ins% | commRef | commCur | comm%
' ============================================================================
Private Sub BuildTableSlide(ByVal ppSlide As Object, ByVal sheetName As String, ByVal slideTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, ByVal slideH As Single)

10      On Error GoTo ERR_HANDLER

        Dim ws As Worksheet
        Dim lastRow As Long
        Dim r As Long
        Dim shp As Object
20      Set ws = ThisWorkbook.Worksheets(sheetName)
30      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        Dim nItems As Long
40      nItems = lastRow - 3
50      If nItems < 1 Then Exit Sub

        ' Read data from sheet
        Dim arrNames() As String
        Dim arrPremR() As Double
        Dim arrPremC() As Double
        Dim arrDocR() As Long
        Dim arrDocC() As Long
        Dim arrInsR() As Long
        Dim arrInsC() As Long
        Dim arrCommR() As Double
        Dim arrCommC() As Double

60      ReDim arrNames(1 To nItems)
70      ReDim arrPremR(1 To nItems)
72      ReDim arrPremC(1 To nItems)
80      ReDim arrDocR(1 To nItems)
82      ReDim arrDocC(1 To nItems)
90      ReDim arrInsR(1 To nItems)
92      ReDim arrInsC(1 To nItems)
100     ReDim arrCommR(1 To nItems)
102     ReDim arrCommC(1 To nItems)

        Dim idx As Long
110     idx = 0
120     For r = 3 To lastRow - 1
130         idx = idx + 1
140         If idx > nItems Then Exit For
150         arrNames(idx) = Trim$(CStr(ws.Cells(r, 1).Value2))
160         arrPremR(idx) = CDbl(ws.Cells(r, 2).Value2)
170         arrPremC(idx) = CDbl(ws.Cells(r, 3).Value2)
180         arrDocR(idx) = CLng(ws.Cells(r, 5).Value2)
190         arrDocC(idx) = CLng(ws.Cells(r, 6).Value2)
200         arrInsR(idx) = CLng(ws.Cells(r, 8).Value2)
210         arrInsC(idx) = CLng(ws.Cells(r, 9).Value2)
220         arrCommR(idx) = CDbl(ws.Cells(r, 14).Value2)
230         arrCommC(idx) = CDbl(ws.Cells(r, 15).Value2)
240     Next r
250     nItems = idx

        ' Title
260     Set shp = ppSlide.Shapes.AddTextbox(1, 20, 5, slideW - 40, 35)
270     shp.TextFrame.TextRange.Text = slideTitle & " - " & ChrW(1496) & ChrW(1489) & ChrW(1500) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
280     shp.TextFrame.TextRange.Font.Size = 18
290     shp.TextFrame.TextRange.Font.Bold = True
300     shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
310     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
320     shp.TextFrame.WordWrap = True

        ' Table: nItems + 3 rows (2 headers + data + total), 13 cols
        Dim tblRows As Long
        Dim tblCols As Long
        Dim tblTop As Single
        Dim tblLeft As Single
        Dim tblWidth As Single
        Dim tblHeight As Single
        Dim rowH As Single
        Dim ppTbl As Object
        Dim tbl As Object
        Dim c As Long
        Dim tblR As Long
        Dim blueClr As Long
        Dim lightBlue As Long
        Dim pctLabel As String

330     tblRows = nItems + 3
340     tblCols = 13
350     tblTop = 42
360     tblLeft = 5
370     tblWidth = slideW - 10
372     Dim availableH As Single
374     Dim dataFontSize As Single
376     Dim headerFontSize As Single
378     availableH = slideH - tblTop - 28
380     rowH = availableH / tblRows
390     If rowH > 22 Then rowH = 22
392     If rowH < 7 Then rowH = 7
394     dataFontSize = rowH - 4
396     If dataFontSize > 12 Then dataFontSize = 12
398     If dataFontSize < 5 Then dataFontSize = 5
399     headerFontSize = dataFontSize
400     tblHeight = rowH * tblRows

410     Set ppTbl = ppSlide.Shapes.AddTable(tblRows, tblCols, tblLeft, tblTop, tblWidth, tblHeight)
420     Set tbl = ppTbl.Table
422     For tblR = 1 To tblRows
424         tbl.Rows(tblR).Height = rowH
426     Next tblR
427     ppTbl.LockAspectRatio = 0
428     ppTbl.Left = tblLeft
429     ppTbl.Top = tblTop

        ' Column widths (13 cols)
430     tbl.Columns(1).Width = tblWidth * 0.08
440     tbl.Columns(2).Width = tblWidth * 0.105
450     tbl.Columns(3).Width = tblWidth * 0.105
460     tbl.Columns(4).Width = tblWidth * 0.06
470     tbl.Columns(5).Width = tblWidth * 0.07
480     tbl.Columns(6).Width = tblWidth * 0.07
490     tbl.Columns(7).Width = tblWidth * 0.06
500     tbl.Columns(8).Width = tblWidth * 0.07
510     tbl.Columns(9).Width = tblWidth * 0.07
520     tbl.Columns(10).Width = tblWidth * 0.06
530     tbl.Columns(11).Width = tblWidth * 0.09
540     tbl.Columns(12).Width = tblWidth * 0.09
550     tbl.Columns(13).Width = tblWidth * 0.07

560     blueClr = RGB(0, 100, 170)
570     lightBlue = RGB(180, 210, 240)
580     pctLabel = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497)

        ' Header row 1 - category names
590     tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = ""
        ' premiot
600     tbl.Cell(1, 2).Shape.TextFrame.TextRange.Text = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514)
        ' mismachim
610     tbl.Cell(1, 5).Shape.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501)
        ' mevutachim
620     tbl.Cell(1, 8).Shape.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501)
        ' amulot
630     tbl.Cell(1, 11).Shape.TextFrame.TextRange.Text = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514)

        ' Header row 2 - year sub-headers
640     tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = ""
650     tbl.Cell(2, 2).Shape.TextFrame.TextRange.Text = refYear
660     tbl.Cell(2, 3).Shape.TextFrame.TextRange.Text = yearVal
670     tbl.Cell(2, 4).Shape.TextFrame.TextRange.Text = pctLabel
680     tbl.Cell(2, 5).Shape.TextFrame.TextRange.Text = refYear
690     tbl.Cell(2, 6).Shape.TextFrame.TextRange.Text = yearVal
700     tbl.Cell(2, 7).Shape.TextFrame.TextRange.Text = pctLabel
710     tbl.Cell(2, 8).Shape.TextFrame.TextRange.Text = refYear
720     tbl.Cell(2, 9).Shape.TextFrame.TextRange.Text = yearVal
730     tbl.Cell(2, 10).Shape.TextFrame.TextRange.Text = pctLabel
740     tbl.Cell(2, 11).Shape.TextFrame.TextRange.Text = refYear
750     tbl.Cell(2, 12).Shape.TextFrame.TextRange.Text = yearVal
760     tbl.Cell(2, 13).Shape.TextFrame.TextRange.Text = pctLabel

        ' Format header rows
770     For c = 1 To tblCols
780         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
782         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Size = headerFontSize
790         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Bold = True
800         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
810         tbl.Cell(1, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
820         tbl.Cell(1, c).Shape.Fill.ForeColor.RGB = blueClr
830         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
832         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Size = headerFontSize
840         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Bold = True
850         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
860         tbl.Cell(2, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
870         tbl.Cell(2, c).Shape.Fill.ForeColor.RGB = blueClr
880     Next c

        ' Data rows
        Dim chgVal As Double
890     For idx = 1 To nItems
900         tblR = idx + 2
910         tbl.Cell(tblR, 1).Shape.TextFrame.TextRange.Text = arrNames(idx)
            ' Premiums
920         tbl.Cell(tblR, 2).Shape.TextFrame.TextRange.Text = Format$(arrPremR(idx), "#,##0")
930         tbl.Cell(tblR, 3).Shape.TextFrame.TextRange.Text = Format$(arrPremC(idx), "#,##0")
940         If arrPremR(idx) <> 0 Then
950             chgVal = (arrPremC(idx) - arrPremR(idx)) / Abs(arrPremR(idx)) * 100
960             tbl.Cell(tblR, 4).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
970         Else
980             tbl.Cell(tblR, 4).Shape.TextFrame.TextRange.Text = "-"
990         End If
            ' Documents
1000        tbl.Cell(tblR, 5).Shape.TextFrame.TextRange.Text = Format$(arrDocR(idx), "#,##0")
1010        tbl.Cell(tblR, 6).Shape.TextFrame.TextRange.Text = Format$(arrDocC(idx), "#,##0")
1020        If arrDocR(idx) <> 0 Then
1030            chgVal = (CDbl(arrDocC(idx)) - CDbl(arrDocR(idx))) / Abs(CDbl(arrDocR(idx))) * 100
1040            tbl.Cell(tblR, 7).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
1050        Else
1060            tbl.Cell(tblR, 7).Shape.TextFrame.TextRange.Text = "-"
1070        End If
            ' Insured
1080        tbl.Cell(tblR, 8).Shape.TextFrame.TextRange.Text = Format$(arrInsR(idx), "#,##0")
1090        tbl.Cell(tblR, 9).Shape.TextFrame.TextRange.Text = Format$(arrInsC(idx), "#,##0")
1100        If arrInsR(idx) <> 0 Then
1110            chgVal = (CDbl(arrInsC(idx)) - CDbl(arrInsR(idx))) / Abs(CDbl(arrInsR(idx))) * 100
1120            tbl.Cell(tblR, 10).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
1130        Else
1140            tbl.Cell(tblR, 10).Shape.TextFrame.TextRange.Text = "-"
1150        End If
            ' Commissions
1160        tbl.Cell(tblR, 11).Shape.TextFrame.TextRange.Text = Format$(arrCommR(idx), "#,##0")
1170        tbl.Cell(tblR, 12).Shape.TextFrame.TextRange.Text = Format$(arrCommC(idx), "#,##0")
1180        If arrCommR(idx) <> 0 Then
1190            chgVal = (arrCommC(idx) - arrCommR(idx)) / Abs(arrCommR(idx)) * 100
1200            tbl.Cell(tblR, 13).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
1210        Else
1220            tbl.Cell(tblR, 13).Shape.TextFrame.TextRange.Text = "-"
1230        End If

            ' Format data cells
1240        For c = 1 To tblCols
1250            tbl.Cell(tblR, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
1252            tbl.Cell(tblR, c).Shape.TextFrame.TextRange.Font.Size = dataFontSize
1260            tbl.Cell(tblR, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
1270            If idx Mod 2 = 0 Then
1280                tbl.Cell(tblR, c).Shape.Fill.ForeColor.RGB = lightBlue
1290            Else
1300                tbl.Cell(tblR, c).Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
1310            End If
1320        Next c
1330    Next idx

        ' Total row
        Dim totR As Long
        Dim totPR As Double
        Dim totPC As Double
        Dim totCR As Double
        Dim totCC As Double
        Dim totDR As Long
        Dim totDC As Long
        Dim totIR As Long
        Dim totIC As Long
1340    totR = nItems + 3
1350    totPR = CDbl(ws.Cells(lastRow, 2).Value2)
1360    totPC = CDbl(ws.Cells(lastRow, 3).Value2)
1370    totDR = CLng(ws.Cells(lastRow, 5).Value2)
1380    totDC = CLng(ws.Cells(lastRow, 6).Value2)
1390    totIR = CLng(ws.Cells(lastRow, 8).Value2)
1400    totIC = CLng(ws.Cells(lastRow, 9).Value2)
1410    totCR = CDbl(ws.Cells(lastRow, 14).Value2)
1420    totCC = CDbl(ws.Cells(lastRow, 15).Value2)

        ' Total label: SahaK
1430    tbl.Cell(totR, 1).Shape.TextFrame.TextRange.Text = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499)
1440    tbl.Cell(totR, 2).Shape.TextFrame.TextRange.Text = Format$(totPR, "#,##0")
1450    tbl.Cell(totR, 3).Shape.TextFrame.TextRange.Text = Format$(totPC, "#,##0")
1460    If totPR <> 0 Then
1470        tbl.Cell(totR, 4).Shape.TextFrame.TextRange.Text = Format$((totPC - totPR) / Abs(totPR) * 100, "0.0") & "%"
1480    Else
1490        tbl.Cell(totR, 4).Shape.TextFrame.TextRange.Text = "-"
1500    End If
1510    tbl.Cell(totR, 5).Shape.TextFrame.TextRange.Text = Format$(totDR, "#,##0")
1520    tbl.Cell(totR, 6).Shape.TextFrame.TextRange.Text = Format$(totDC, "#,##0")
1530    If totDR <> 0 Then
1540        tbl.Cell(totR, 7).Shape.TextFrame.TextRange.Text = Format$((CDbl(totDC) - CDbl(totDR)) / Abs(CDbl(totDR)) * 100, "0.0") & "%"
1550    Else
1560        tbl.Cell(totR, 7).Shape.TextFrame.TextRange.Text = "-"
1570    End If
1580    tbl.Cell(totR, 8).Shape.TextFrame.TextRange.Text = Format$(totIR, "#,##0")
1590    tbl.Cell(totR, 9).Shape.TextFrame.TextRange.Text = Format$(totIC, "#,##0")
1600    If totIR <> 0 Then
1610        tbl.Cell(totR, 10).Shape.TextFrame.TextRange.Text = Format$((CDbl(totIC) - CDbl(totIR)) / Abs(CDbl(totIR)) * 100, "0.0") & "%"
1620    Else
1630        tbl.Cell(totR, 10).Shape.TextFrame.TextRange.Text = "-"
1640    End If
1650    tbl.Cell(totR, 11).Shape.TextFrame.TextRange.Text = Format$(totCR, "#,##0")
1660    tbl.Cell(totR, 12).Shape.TextFrame.TextRange.Text = Format$(totCC, "#,##0")
1670    If totCR <> 0 Then
1680        tbl.Cell(totR, 13).Shape.TextFrame.TextRange.Text = Format$((totCC - totCR) / Abs(totCR) * 100, "0.0") & "%"
1690    Else
1700        tbl.Cell(totR, 13).Shape.TextFrame.TextRange.Text = "-"
1710    End If

        ' Format total row
1720    For c = 1 To tblCols
1730        tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
1732        tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Size = dataFontSize
1740        tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Bold = True
1750        tbl.Cell(totR, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
1760        tbl.Cell(totR, c).Shape.Fill.ForeColor.RGB = RGB(220, 230, 240)
1770    Next c

1780    Exit Sub
ERR_HANDLER:
1790    Err.Raise Err.Number, "BuildTableSlide(" & sheetName & "):" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Build one slide from the TOTAL row only
' Used for Drachim: no visible report sheet, one year-vs-year summary slide.
' ============================================================================
Private Sub BuildTotalRowTableSlide(ByVal ppSlide As Object, ByVal sheetName As String, ByVal slideTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, ByVal slideH As Single)

10      On Error GoTo ERR_HANDLER

        Dim ws As Worksheet
        Dim lastRow As Long
        Dim totalRow As Long
        Dim shp As Object
        Dim tbl As Object
        Dim rr As Long
        Dim cc As Long

20      Set ws = ThisWorkbook.Worksheets(sheetName)
30      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
40      If lastRow < 3 Then Exit Sub
50      totalRow = lastRow

60      Set shp = ppSlide.Shapes.AddTextbox(1, 20, 8, slideW - 40, 42)
70      shp.TextFrame.TextRange.Text = slideTitle & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
80      shp.TextFrame.TextRange.Font.Size = 22
90      shp.TextFrame.TextRange.Font.Bold = True
100     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
110     shp.TextFrame.WordWrap = True

120     Set tbl = ppSlide.Shapes.AddTable(6, 4, 70, 85, slideW - 140, 330).Table

130     tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Metric"
140     tbl.Cell(1, 2).Shape.TextFrame.TextRange.Text = refYear
150     tbl.Cell(1, 3).Shape.TextFrame.TextRange.Text = yearVal
160     tbl.Cell(1, 4).Shape.TextFrame.TextRange.Text = "%"

170     tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492)
180     tbl.Cell(2, 2).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 2).Value2, "#,##0")
190     tbl.Cell(2, 3).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 3).Value2, "#,##0")
200     tbl.Cell(2, 4).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 4).Value2, "0.0%")

210     tbl.Cell(3, 1).Shape.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501)
220     tbl.Cell(3, 2).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 5).Value2, "#,##0")
230     tbl.Cell(3, 3).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 6).Value2, "#,##0")
240     tbl.Cell(3, 4).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 7).Value2, "0.0%")

250     tbl.Cell(4, 1).Shape.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501)
260     tbl.Cell(4, 2).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 8).Value2, "#,##0")
270     tbl.Cell(4, 3).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 9).Value2, "#,##0")
280     tbl.Cell(4, 4).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 10).Value2, "0.0%")

290     tbl.Cell(5, 1).Shape.TextFrame.TextRange.Text = ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1493) & ChrW(1514) & " " & ChrW(1508) & ChrW(1506) & ChrW(1497) & ChrW(1500) & ChrW(1493) & ChrW(1514)
300     tbl.Cell(5, 2).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 11).Value2, "#,##0")
310     tbl.Cell(5, 3).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 12).Value2, "#,##0")
320     tbl.Cell(5, 4).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 13).Value2, "0.0%")

330     tbl.Cell(6, 1).Shape.TextFrame.TextRange.Text = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1492)
340     tbl.Cell(6, 2).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 14).Value2, "#,##0")
350     tbl.Cell(6, 3).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 15).Value2, "#,##0")
360     tbl.Cell(6, 4).Shape.TextFrame.TextRange.Text = Format(ws.Cells(totalRow, 16).Value2, "0.0%")

370     For rr = 1 To 6
380         tbl.Rows(rr).Height = 48
390         For cc = 1 To 4
400             tbl.Cell(rr, cc).Shape.TextFrame.TextRange.Font.Size = 16
410             tbl.Cell(rr, cc).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
420         Next cc
430     Next rr

440     For cc = 1 To 4
450         tbl.Cell(1, cc).Shape.TextFrame.TextRange.Font.Bold = True
460     Next cc

470     Exit Sub

ERR_HANDLER:
480     Err.Raise Err.Number, "BuildTotalRowTableSlide:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Add page number to slide
' ============================================================================
Private Sub AddPageNumber(ByVal ppSlide As Object, ByVal pageNum As Long, ByVal totalPages As Long, ByVal slideW As Single, ByVal slideH As Single)

10      On Error Resume Next

        Dim shp As Object
20      Set shp = ppSlide.Shapes.AddTextbox(1, slideW - 120, 2, 110, 22)
30      shp.TextFrame.TextRange.Text = pageNum & " / " & totalPages
40      shp.TextFrame.TextRange.Font.Name = "Arial"
50      shp.TextFrame.TextRange.Font.Size = 10
60      shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
70      shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
80      shp.TextFrame.WordWrap = False
90      shp.TextFrame.MarginTop = 0
100     shp.TextFrame.MarginBottom = 0
110     shp.ZOrder 0

End Sub



' ============================================================================
' SHEET NAME FUNCTIONS (Hebrew via ChrW)
' ============================================================================

Private Function SOURCE_FOLDER() As String
    ' English-only path to avoid VBA encoding problems.
    ' Put source files here, for example: C:\LEVAV PROJECT\SOURCE\2025.xlsx
    SOURCE_FOLDER = "C:\LEVAV PROJECT\SOURCE\"
End Function

Private Function CONTROL_SHEET_NAME() As String
    ' daf habait
    CONTROL_SHEET_NAME = ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
End Function

Private Function MANAGEMENT_SHEET_NAME() As String
    ' hagdarot
    MANAGEMENT_SHEET_NAME = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

Private Function REVIEW_SHEET_NAME() As String
    ' letipul
    REVIEW_SHEET_NAME = ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
End Function

Private Function SHEET_COMPANIES() As String
    ' hevrot
    SHEET_COMPANIES = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

Private Function SHEET_BRANCH() As String
    ' anafim
    SHEET_BRANCH = ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_MAINBRANCH() As String
    ' anaf merkaz
    SHEET_MAINBRANCH = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
End Function

Private Function SHEET_TELLERS() As String
    ' tlerim
    SHEET_TELLERS = ChrW(1496) & ChrW(1500) & ChrW(1512) & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_AGENTS() As String
    ' sochnim
    SHEET_AGENTS = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_AGENTS_NO_LEVAV() As String
    SHEET_AGENTS_NO_LEVAV = "Agents_No_Levav"
End Function

Private Function SHEET_DRACHIM() As String
    SHEET_DRACHIM = "Company_Drachim"
End Function

Private Function TITLE_MONTHS() As String
    TITLE_MONTHS = SHEET_MONTHS()
End Function

Private Function TITLE_COMPANIES() As String
    TITLE_COMPANIES = SHEET_COMPANIES()
End Function

Private Function TITLE_MAINBRANCH() As String
    TITLE_MAINBRANCH = SHEET_MAINBRANCH()
End Function

Private Function TITLE_TELLERS() As String
    TITLE_TELLERS = SHEET_TELLERS()
End Function

Private Function TITLE_AGENTS() As String
    TITLE_AGENTS = SHEET_AGENTS()
End Function

Private Function TITLE_AGENTS_NO_LEVAV() As String
    ' sochnim lelo Levav
    TITLE_AGENTS_NO_LEVAV = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & HEB_LEVAV()
End Function

Private Function TITLE_DRACHIM() As String
    ' hevrat Drachim
    TITLE_DRACHIM = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1514) & " " & HEB_DRACHIM()
End Function

Private Function HEB_LEVAV() As String
    HEB_LEVAV = ChrW(1500) & ChrW(1489) & ChrW(1489)
End Function

Private Function HEB_DRACHIM() As String
    HEB_DRACHIM = ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1497) & ChrW(1501)
End Function


Private Function SHEET_MONTHS() As String
    ' hodshim
    SHEET_MONTHS = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_SUMMARY() As String
    ' sikum
    SHEET_SUMMARY = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501)
End Function


' ============================================================================
' MACRO: SendForReview - collects rows marked "???? ??????" and sends via Outlook
' Called from the "?????? ?????" button on the review sheet
' ============================================================================
Public Sub SendForReview()

10      On Error GoTo ERR_HANDLER

20      Dim wsRev As Worksheet
30      Set wsRev = ActiveSheet

        ' Verify we are on a review sheet (name starts with REVIEW_SHEET_NAME)
40      If InStr(1, wsRev.Name, REVIEW_SHEET_NAME(), vbTextCompare) = 0 Then
50          MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1502) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500), vbExclamation
51      Dim wsMsg As Worksheet
52      Set wsMsg = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())

53      If Trim$(CStr(wsMsg.Cells(18, 19).Value)) <> "" Then
54      MsgBoxU CStr(wsMsg.Cells(18, 19).Value), vbInformation
End If
60          Exit Sub
70      End If

        ' Find action column (header contains "peula" = ?????)
80      Dim lastCol As Long
90      lastCol = wsRev.Cells(1, wsRev.Columns.Count).End(xlToLeft).Column
100     Dim actionCol As Long
110     actionCol = 0
120     Dim c As Long
130     For c = 1 To lastCol
140         If InStr(1, CStr(wsRev.Cells(1, c).Value2), ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492), vbTextCompare) > 0 Then
150             actionCol = c
160             Exit For
170         End If
180     Next c
190     If actionCol = 0 Then
200         MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1492) & " " & ChrW(1506) & ChrW(1502) & ChrW(1493) & ChrW(1491) & ChrW(1514) & " " & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492), vbExclamation
210         Exit Sub
220     End If

        ' Count rows with "ha'aver livdika"
230     Dim lastRow As Long
240     lastRow = wsRev.Cells(wsRev.Rows.Count, 1).End(xlUp).Row
250     Dim sendCount As Long
260     sendCount = 0
270     Dim r As Long
        ' "ha'aver livdika" = ???? ??????
        Dim actionMatch As String
280     actionMatch = ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)
290     For r = 2 To lastRow
300         If InStr(1, CStr(wsRev.Cells(r, actionCol).Value2), actionMatch, vbTextCompare) > 0 Then
310             sendCount = sendCount + 1
320         End If
330     Next r

340    If sendCount = 0 Then
    Exit Sub
End If

        ' Get email address from parameters
380     Dim wsMgmt As Worksheet
390     Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
400     Dim emailAddr As String
410     emailAddr = GetStringParameter(wsMgmt, PARAM_ERROR_EMAIL)
420     If emailAddr = "" Then
            ' "lo hugdra ktovet email" = email address not defined
430         MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1492) & ChrW(1493) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1492) & " " & ChrW(1499) & ChrW(1514) & ChrW(1493) & ChrW(1489) & ChrW(1514) & " " & ChrW(1488) & ChrW(1497) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514), vbExclamation
440         Exit Sub
450     End If

        ' Create temp workbook with matching rows
460     Dim wbTemp As Workbook
470     Set wbTemp = Workbooks.Add
480     Dim wsTemp As Worksheet
490     Set wsTemp = wbTemp.Worksheets(1)

        ' Copy header row
500     Dim hdrCol As Long
510     For hdrCol = 1 To lastCol
520         wsTemp.Cells(1, hdrCol).Value = wsRev.Cells(1, hdrCol).Value
530     Next hdrCol
540     wsTemp.Rows(1).Font.Bold = True

        ' Copy matching rows
550     Dim outRow As Long
560     outRow = 2
570     For r = 2 To lastRow
580         If InStr(1, CStr(wsRev.Cells(r, actionCol).Value2), actionMatch, vbTextCompare) > 0 Then
590             For hdrCol = 1 To lastCol
600                 wsTemp.Cells(outRow, hdrCol).Value = wsRev.Cells(r, hdrCol).Value
610             Next hdrCol
620             outRow = outRow + 1
630         End If
640     Next r
650     wsTemp.Columns.AutoFit

        ' Save temp file
660     Dim tempPath As String
670     tempPath = ThisWorkbook.Path & "\" & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & "_" & wsRev.Name & ".xlsx"
680     Application.DisplayAlerts = False
690     wbTemp.SaveAs tempPath, xlOpenXMLWorkbook
700     wbTemp.Close SaveChanges:=False
710     Application.DisplayAlerts = True

        ' Build email body
        ' "hi lahav" = ?? ???
720     Dim bodyLine1 As String
730     bodyLine1 = ChrW(1492) & ChrW(1497) & " " & ChrW(1500) & ChrW(1492) & ChrW(1489)
        ' "likrat hachanat doch avurchem nimtze'u hachrigim haram" = ????? ???? ??? ?????? ????? ??????? ??"?
740     Dim bodyLine2 As String
750     bodyLine2 = ChrW(1500) & ChrW(1511) & ChrW(1512) & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1499) & ChrW(1504) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1499) & ChrW(1501) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1492) & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & " " & ChrW(1492) & ChrW(1512) & ChrW(34) & ChrW(1502)
        ' "al mnat lehafik et hadoch ani mevakeshet tguvatcha al mnat she'etaken beheta'am"
760     Dim bodyLine3 As String
770     bodyLine3 = ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & ChrW(1500) & ChrW(1492) & ChrW(1508) & ChrW(1497) & ChrW(1511) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1491) & ChrW(1493) & ChrW(1495) & " " & ChrW(1488) & ChrW(1504) & ChrW(1497) & " " & ChrW(1502) & ChrW(1489) & ChrW(1511) & ChrW(1513) & ChrW(1514) & " " & ChrW(1514) & ChrW(1490) & ChrW(1493) & ChrW(1489) & ChrW(1514) & ChrW(1498) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & ChrW(1513) & ChrW(1488) & ChrW(1514) & ChrW(1511) & ChrW(1503) & " " & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1488) & ChrW(1501)
        ' toda = ????
780     Dim bodyLine4 As String
790     bodyLine4 = ChrW(1514) & ChrW(1493) & ChrW(1491) & ChrW(1492)
        ' orit = ?????
800     Dim bodyLine5 As String
810     bodyLine5 = ChrW(1488) & ChrW(1493) & ChrW(1512) & ChrW(1497) & ChrW(1514)

820     Dim emailBody As String
830     emailBody = bodyLine1 & vbCrLf & vbCrLf & bodyLine2 & vbCrLf & bodyLine3 & vbCrLf & vbCrLf & bodyLine4 & vbCrLf & bodyLine5

        ' Email subject: "charigim shenimtze'u letipulcha" = ?????? ?????? ???????
840     Dim emailSubject As String
850     emailSubject = ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & " " & ChrW(1513) & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1498)

        ' Create Outlook email (late binding)
860     Dim olApp As Object
870     Dim olMail As Object
880     Set olApp = CreateObject("Outlook.Application")
890     Set olMail = olApp.CreateItem(0)
900     olMail.To = emailAddr
910     olMail.Subject = emailSubject
920     olMail.Body = emailBody
930     olMail.Attachments.Add tempPath
940     olMail.Display
945     If Trim$(CStr(ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME()).Cells(18, 19).Value)) <> "" Then
946         MsgBoxU CStr(ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME()).Cells(18, 19).Value), vbInformation
947     End If

        ' Success message: "email huchan be'hatzlacha im X shurot" = ???? ???? ?????? ?? X ?????
950     MsgBoxU ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & ChrW(1492) & ChrW(1493) & ChrW(1499) & ChrW(1503) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & " " & ChrW(1506) & ChrW(1501) & " " & sendCount & " " & ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514), vbInformation

960     Exit Sub

ERR_HANDLER:
970     MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1495) & ChrW(1514) & " " & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & ":" & vbCrLf & Err.Description, vbCritical

End Sub

Public Sub CreateCompaniesChart()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim chObj As ChartObject
    Dim rngCats As Range

   Set ws = ThisWorkbook.Worksheets(SHEET_COMPANIES())

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in Companies sheet", vbExclamation
        Exit Sub
    End If

    ' Delete old chart
    Dim obj As ChartObject
    For Each obj In ws.ChartObjects
        obj.Delete
    Next obj

    Set rngCats = ws.Range("A2:A" & lastRow)

    ' Create chart
    Set chObj = ws.ChartObjects.Add( _
        Left:=ws.Range("Q2").Left, _
        Top:=ws.Range("Q2").Top, _
        Width:=700, _
        Height:=400)

    With chObj.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .chartTitle.Text = "פרמיה ועמלה לפי חברה"

        ' Premium series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = ws.Range("B1").Value
        .SeriesCollection(1).Values = ws.Range("B2:B" & lastRow)
        .SeriesCollection(1).XValues = rngCats

        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = ws.Range("C1").Value
        .SeriesCollection(2).Values = ws.Range("C2:C" & lastRow)
        .SeriesCollection(2).XValues = rngCats

        ' Commission series
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Name = ws.Range("N1").Value
        .SeriesCollection(3).Values = ws.Range("N2:N" & lastRow)
        .SeriesCollection(3).XValues = rngCats

        .SeriesCollection.NewSeries
        .SeriesCollection(4).Name = ws.Range("O1").Value
        .SeriesCollection(4).Values = ws.Range("O2:O" & lastRow)
        .SeriesCollection(4).XValues = rngCats

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        .Axes(xlCategory).TickLabels.Orientation = 45
    End With

    MsgBox "Companies chart created successfully", vbInformation

End Sub

Private Function GetReviewSheet() As Worksheet
    Set GetReviewSheet = ThisWorkbook.Worksheets(ReviewSheetName())
End Function




