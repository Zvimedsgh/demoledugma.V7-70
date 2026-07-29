VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================================
' Sheet1 Code (דף הבית) — v7.33
' Place in: VBA Editor > Sheet1 (דף הבית) code module
' ============================================================================

Private Const HOME_GREEN As Long = 14024924  ' RGB(220, 240, 220) = &HDCF0DC

Private Sub Worksheet_Activate()
    Me.Range("A1").Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Target.Cells.Count > 1 Then Exit Sub
    
    Dim goldColor As Long
    goldColor = RGB(255, 215, 0)
    
    ' --- G5 changed (Period Type) ---
    If Not Intersect(Target, Me.Range("G5")) Is Nothing Then
        ' If not "שנתי" (annual) - move to G6, show "בחר/י", highlight gold, add dropdown
        If Target.Value <> ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497) Then
            Application.EnableEvents = False
            
            ' Set G6 dropdown from period list (based on G5 selection)
            On Error Resume Next
            Me.Range("G6").Validation.Delete
            On Error GoTo 0
            
            ' Determine which named range to use for G6 dropdown
            Dim periodList As String
            Dim g5Val As String
            g5Val = Trim$(CStr(Target.Value))
            
            Select Case g5Val
                Case ChrW(1495) & ChrW(1510) & ChrW(1497) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)  ' חצי שנתי
                    periodList = "=lst_half_year"
                Case ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497)  ' רבעוני
                    periodList = "=lst_quarter"
                Case ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497)  ' חודשי
                    periodList = "=lst_month"
                Case Else
                    periodList = ""
            End Select
            
            If Len(periodList) > 0 Then
                On Error Resume Next
                Me.Range("G6").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=periodList
                Me.Range("G6").Validation.InCellDropdown = True
                On Error GoTo 0
            End If
            
            Me.Range("G6").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)  ' בחר/י
            Me.Range("G6").Interior.Color = goldColor
            Me.Range("G6").Select
            Application.EnableEvents = True
        Else
            ' If "שנתי" selected - clear G6, restore green, remove dropdown
            Application.EnableEvents = False
            Me.Range("G6").Value = ""
            Me.Range("G6").Interior.Color = HOME_GREEN
            On Error Resume Next
            Me.Range("G6").Validation.Delete
            On Error GoTo 0
            Application.EnableEvents = True
        End If
    End If
    
    ' --- G6 changed (Period Value) - restore green when value selected ---
    If Not Intersect(Target, Me.Range("G6")) Is Nothing Then
        If Len(Trim$(Target.Value)) > 0 And Target.Value <> ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497) Then
            Me.Range("G6").Interior.Color = HOME_GREEN
        End If
    End If
    
    ' --- G9 changed (Filter Type) ---
    If Not Intersect(Target, Me.Range("G9")) Is Nothing Then
        Dim filterType As String
        filterType = Trim$(CStr(Target.Value))
        
        ' Clear G10 and set dynamic dropdown
        Application.EnableEvents = False
        On Error Resume Next
        Me.Range("G10").Validation.Delete
        On Error GoTo 0
        
        ' Determine which column in "רשימות" sheet to use
        Dim listCol As Long
        listCol = 0
        
        Select Case filterType
            Case ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492)  ' חברה
                listCol = 1  ' Column A
            Case ChrW(1496) & ChrW(1500) & ChrW(1512)  ' טלר
                listCol = 2  ' Column B
            Case ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503)  ' סוכן
                listCol = 3  ' Column C
            Case ChrW(1506) & ChrW(1504) & ChrW(1507)  ' ענף
                listCol = 4  ' Column D
            Case ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)  ' ענף מרכז
                listCol = 5  ' Column E
        End Select
        
        If listCol > 0 Then
            ' Build dropdown from רשימות sheet
            Dim wsLists As Worksheet
            Dim lastRow As Long
            Dim listRange As String
            
            On Error Resume Next
            Set wsLists = ThisWorkbook.Worksheets(ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1493) & ChrW(1514))  ' רשימות
            On Error GoTo 0
            
            If Not wsLists Is Nothing Then
                lastRow = wsLists.Cells(wsLists.Rows.Count, listCol).End(xlUp).Row
                If lastRow >= 1 Then
                    listRange = "='" & wsLists.Name & "'!" & "$" & Chr(64 + listCol) & "$1:$" & Chr(64 + listCol) & "$" & lastRow
                    On Error Resume Next
                    Me.Range("G10").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listRange
                    Me.Range("G10").Validation.InCellDropdown = True
                    On Error GoTo 0
                End If
            End If
            
            ' Set "בחר/י" in G10, highlight gold, move cursor there
            Me.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)  ' בחר/י
            Me.Range("G10").Interior.Color = goldColor
            Me.Range("G10").Select
        Else
            ' "בחר/י" selected in G9 - reset G10 to default
            Me.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)  ' בחר/י
            Me.Range("G10").Interior.Color = HOME_GREEN
        End If
        
        Application.EnableEvents = True
    End If
    
    ' --- G10 changed (Filter Value) - restore green when value selected ---
    If Not Intersect(Target, Me.Range("G10")) Is Nothing Then
        If Len(Trim$(Target.Value)) > 0 And Target.Value <> ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497) Then
            Me.Range("G10").Interior.Color = HOME_GREEN
        End If
    End If

End Sub
