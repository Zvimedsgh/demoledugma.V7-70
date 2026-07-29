import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.13.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Replace IsUserAuthorized and IsSystemInDemoMode with GetUserAccessLevel
helpers = '''Private g_CachedUserAccess As Integer ' 0=uninit, 1=unauth, 2=limited, 3=full

Public Function GetUserAccessLevel() As Integer
    If g_CachedUserAccess > 0 Then
        GetUserAccessLevel = g_CachedUserAccess
        Exit Function
    End If
    
    Dim winUser As String
    winUser = LCase$(Trim$(Environ("USERNAME")))
    If Len(winUser) = 0 Then
        g_CachedUserAccess = 1 ' unauthorized
        GetUserAccessLevel = 1
        Exit Function
    End If
    
    Dim wsMgmt As Worksheet
    On Error Resume Next
    Set wsMgmt = ThisWorkbook.Worksheets(H_SET_PERMISSIONS())
    If wsMgmt Is Nothing Then
        g_CachedUserAccess = 1
        GetUserAccessLevel = 1
        Exit Function
    End If
    
    Dim r As Long
    Dim cellVal As String
    Dim accessLevel As String
    r = 2
    Do While True
        cellVal = Trim$(CStr(wsMgmt.Cells(r, 1).Value2))
        If UCase$(cellVal) = "EOD" Or cellVal = "" Then Exit Do
        If LCase$(cellVal) = winUser Then
            accessLevel = Trim$(CStr(wsMgmt.Cells(r, 2).Value2))
            If UCase$(accessLevel) = "FULLACCESS" Or accessLevel = ChrW(1502) & ChrW(1500) & ChrW(1488) Then
                g_CachedUserAccess = 3 ' full
                GetUserAccessLevel = 3
                Exit Function
            Else
                g_CachedUserAccess = 2 ' limited
                GetUserAccessLevel = 2
                Exit Function
            End If
        End If
        r = r + 1
    Loop
    g_CachedUserAccess = 1 ' unauthorized
    GetUserAccessLevel = 1
End Function

Public Function IsSystemInDemoMode() As Boolean
    If FORCE_DEMO_MODE Then
        IsSystemInDemoMode = True
        Exit Function
    End If
    
    ' If user is unauthorized, force demo mode!
    If GetUserAccessLevel() = 1 Then
        IsSystemInDemoMode = True
        Exit Function
    End If
    
    ' If user is Limited or Full, respect the DEMO_MODE parameter
    Dim demoParamStr As String
    On Error Resume Next
    demoParamStr = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If demoParamStr = ChrW(1499) & ChrW(1503) Or demoParamStr = "YES" Or demoParamStr = "" Then
        IsSystemInDemoMode = True
    Else
        IsSystemInDemoMode = False
    End If
End Function'''

pattern_helpers = re.compile(r'Public Function IsUserAuthorized\(\) As Boolean.*?End Function\n\nPublic Function IsSystemInDemoMode\(\) As Boolean.*?End Function', re.DOTALL)
content = pattern_helpers.sub(helpers, content, count=1)

# Remove the old g_CachedUserAuth
content = content.replace("Private g_CachedUserAuth As Integer ' 0=uninit, 1=true, -1=false\n", "")

# 2. Rewrite CheckUserPermissions completely
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
    
    Dim accessLevelNum As Integer
    accessLevelNum = GetUserAccessLevel()
    
    Dim isDemo As Boolean
    isDemo = IsSystemInDemoMode()
    
    If isDemo Then
        Call HideWorkSheets
        
        On Error Resume Next
        For Each shp In wsMain.Shapes
            Select Case shp.Name
                Case "btnBuildReview", "btnApplyCorrections", "btnShowResults", "btnBuildPresentation"
                    ' Keep active (buttons 1-4)
                Case "btnSaveReports", "btnViewReports", "btnNewClients"
                    ' Buttons 5, 6, 7 - keep color but restrict action
                    shp.OnAction = "DemoModeRestricted"
            End Select
        Next shp
        On Error GoTo PERM_ERR
    ElseIf accessLevelNum = 3 Then
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
                    shp.OnAction = "ShowResultSheets"
                Case "btnBuildPresentation"
                    shp.OnAction = "BuildPresentation"
                Case "btnSaveReports"
                    shp.OnAction = "SaveReportsToFolder"
                Case "btnViewReports"
                    shp.OnAction = "ViewReportsFolder"
                Case "btnNewClients"
                    shp.OnAction = "NewClients"
            End Select
        Next shp
        On Error GoTo PERM_ERR
    Else
        ' Limited access - show only home + new clients, hide everything else
        Dim ws As Worksheet
        ThisWorkbook.Worksheets(homeSheetName).Visible = xlSheetVisible
        If SheetExists(ncSheetName) Then
            ThisWorkbook.Worksheets(ncSheetName).Visible = xlSheetVisible
        End If

        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> homeSheetName And ws.Name <> ncSheetName Then
                Dim hideIt2 As Boolean
                hideIt2 = True
                
                If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then
                    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA1").Value) <> "YES" Then hideIt2 = False
                End If
                If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then
                    If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then hideIt2 = False
                End If
                
                If hideIt2 Then ws.Visible = xlSheetVeryHidden
            End If
        Next ws

        ' Disable buttons 1-5 (remove OnAction + grey them out) - don't hide!
        On Error Resume Next
        For Each shp In wsMain.Shapes
            Select Case shp.Name
                Case "btnBuildReview", "btnApplyCorrections", "btnShowResults", "btnBuildPresentation", "btnSaveReports", "btnViewReports"
                    shp.OnAction = ""  ' disable click
                    shp.Fill.ForeColor.RGB = RGB(180, 180, 180)  ' grey
                Case "btnNewClients"
                    shp.OnAction = "NewClients" ' Keep active so they can type data
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
print('Done step 4')
