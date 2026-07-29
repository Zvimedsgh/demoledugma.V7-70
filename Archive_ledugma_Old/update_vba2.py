import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.13.bas', 'r', encoding='utf-8') as f:
    content = f.read()

new_check_perms = '''Public Sub CheckUserPermissions()
    ' --- Failsafe: Reset Application state ---
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ' -----------------------------------------
    
    Dim wsMain As Worksheet
    Dim homeSheetName As String
    Dim ncSheetName As String
    Dim shp As Shape
    Dim creditVersionCell As String
    Dim dateCell As String
    Dim creditText As String
    
    On Error GoTo PERM_ERR
    
    homeSheetName = CONTROL_SHEET_NAME()
    ncSheetName = NC_MASTER_SHEET_NAME()
    
    Set wsMain = ThisWorkbook.Worksheets(homeSheetName)
    If wsMain Is Nothing Then Exit Sub
    
    ' Fetch credit locations from management sheet for the user (optional, fallback if needed)
    ' Actually, since we don't scan the list here anymore, we can just use default locations or fetch if authorized.
    Dim wsMgmt As Worksheet
    Set wsMgmt = ThisWorkbook.Worksheets(H_SET_PERMISSIONS())
    If Not wsMgmt Is Nothing Then
        Dim winUser As String, r As Long, cellVal As String
        winUser = LCase$(Trim$(Environ("USERNAME")))
        If Len(winUser) > 0 Then
            r = 2
            Do While True
                cellVal = Trim$(CStr(wsMgmt.Cells(r, 1).Value2))
                If UCase$(cellVal) = "EOD" Or cellVal = "" Then Exit Do
                If LCase$(cellVal) = winUser Then
                    creditVersionCell = Trim$(CStr(wsMgmt.Cells(r, 3).Value2))
                    dateCell = Trim$(CStr(wsMgmt.Cells(r, 4).Value2))
                    Exit Do
                End If
                r = r + 1
            Loop
        End If
    End If
    
    Application.ScreenUpdating = False
    
    Dim isDemo As Boolean
    isDemo = IsSystemInDemoMode()
    
    If isDemo Then
        Call HideWorkSheets
        
        On Error Resume Next
        For Each shp In wsMain.Shapes
            Select Case shp.Name
                Case "btnBuildReview", "btnApplyCorrections", "btnShowResults", "btnBuildPresentation"
                    ' Keep active (buttons 1-4)
                    ' They keep their original macros assigned in A00
                Case "btnSaveReports", "btnViewReports", "btnNewClients"
                    ' Buttons 5, 6, 7 - keep color but restrict action
                    shp.OnAction = "DemoModeRestricted"
            End Select
        Next shp
        On Error GoTo PERM_ERR
    Else
        ' Full Access
        Call HideWorkSheets
        
        On Error Resume Next
        For Each shp In wsMain.Shapes
            Select Case shp.Name
                Case "btnBuildReview"
                    shp.OnAction = "BuildReview"
                Case "btnApplyCorrections"
                    shp.OnAction = "ApplyCorrectionsAndBuildReports"
                Case "btnShowResults"
                    shp.OnAction = "ShowResults" ' assuming standard macro name, A00 sets it
                Case "btnBuildPresentation"
                    shp.OnAction = "BuildPresentation"
                Case "btnSaveReports"
                    shp.OnAction = "SaveReportsToFolder"
                Case "btnViewReports"
                    shp.OnAction = "ViewReportsFolder"
                Case "btnNewClients"
                    shp.OnAction = "Nav_NewClients" ' standard macro name
            End Select
        Next shp
        On Error GoTo PERM_ERR
    End If

    ' Write credit+version to Column C cell, date to Column D cell
    creditText = ChrW(1504) & ChrW(1489) & ChrW(1504) & ChrW(1492) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1497) & _
        ChrW(1491) & ChrW(1497) & " " & ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & _
        " 054-6677396"
    On Error Resume Next
    If Len(creditVersionCell) > 0 And creditVersionCell <> "A23" Then
        wsMain.Range(creditVersionCell).ClearContents
        Err.Clear
    End If
    If Len(dateCell) > 0 Then
        wsMain.Range(dateCell).Value = Format$(Date, "dd/mm/yyyy")
        wsMain.Range(dateCell).Font.Size = 9
        wsMain.Range(dateCell).Font.Color = RGB(0, 0, 102)
        wsMain.Range(dateCell).Font.Bold = True
    End If
    On Error GoTo PERM_ERR

    ' --- Update Home Page UI on Open ---
    On Error Resume Next
    wsMain.Unprotect "Z961814r"
    wsMain.Range("A1").Formula = "=GetActiveAgencyName()"
    
    If isDemo Then
        wsMain.Range("A1").Font.Color = RGB(200, 0, 0) ' Red for Demo
    Else
        wsMain.Range("A1").Font.Color = RGB(0, 0, 0) ' Black for Real
    End If
    
    ApplyDemoLockOnOpen
    Application.Goto wsMain.Range("G10")
    wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True
    On Error GoTo 0
    ' -----------------------------------

    ' Navigate to home page
    wsMain.Range("A19").Value = ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " v" & APP_VERSION
    wsMain.Range("A19").Font.Size = 10
    wsMain.Range("A19").Font.Color = RGB(150, 150, 150)
    wsMain.Activate

    Exit Sub
PERM_ERR:
    Application.ScreenUpdating = True
End Sub'''

pattern_perms = re.compile(r'Public Sub CheckUserPermissions\(\).*?End Sub', re.DOTALL)
content = pattern_perms.sub(new_check_perms, content, count=1)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.13.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Done step 2')
