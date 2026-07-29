VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ============================================================================
' LEVAV SYSTEM - ThisWorkbook Code  |  v6.26
' ============================================================================
' PASTE THIS CODE IN: ThisWorkbook module
' (In VBA Editor, double-click "ThisWorkbook" in Project Explorer, then paste)
' ============================================================================
' On close: backup file + hide all sheets except home
' ============================================================================

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo SKIP_CLOSE_ACTIONS
    
    ' ---- Step 1: Create backup ----
    Dim backupPath As String
    Dim wsMgmt As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    On Error Resume Next
    Set wsMgmt = ThisWorkbook.Worksheets(Chr(1492) & Chr(1490) & Chr(1491) & Chr(1512) & Chr(1493) & Chr(1514))
    On Error GoTo SKIP_CLOSE_ACTIONS
    
    If Not wsMgmt Is Nothing Then
        ' Read BACKUP_PATH parameter from settings (column J=10, K=11)
        backupPath = ""
        lastRow = wsMgmt.Cells(wsMgmt.Rows.Count, 10).End(xlUp).Row
        For r = 1 To lastRow
            If UCase$(Trim$(CStr(wsMgmt.Cells(r, 10).Value2))) = "BACKUP_PATH" Then
                backupPath = Trim$(CStr(wsMgmt.Cells(r, 11).Value2))
                Exit For
            End If
        Next r
        
        If backupPath <> "" Then
            ' Ensure trailing backslash
            If Right$(backupPath, 1) <> "\" Then backupPath = backupPath & "\"
            
            ' Create folder if it doesn't exist
            On Error Resume Next
            MkDir backupPath
            On Error GoTo SKIP_CLOSE_ACTIONS
            
            ' Build backup filename with timestamp
            Dim backupFile As String
            backupFile = backupPath & "Levav_backup_" & Format$(Now, "YYYYMMDD_HHMM") & ".xlsm"
            
            ' Save backup copy
            On Error Resume Next
            ThisWorkbook.SaveCopyAs backupFile
            On Error GoTo SKIP_CLOSE_ACTIONS
            
            ' Delete old backups (keep only 10 most recent)
            Dim fso As Object
            Dim folder As Object
            Dim file As Object
            Dim fileCount As Long
            Dim oldestFile As Object
            Dim oldestDate As Date
            
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FolderExists(backupPath) Then
                Set folder = fso.GetFolder(backupPath)
                
                ' Count backup files
                fileCount = 0
                For Each file In folder.Files
                    If Left$(file.Name, 13) = "Levav_backup_" And Right$(file.Name, 5) = ".xlsm" Then
                        fileCount = fileCount + 1
                    End If
                Next file
                
                ' Delete oldest if more than 10
                Do While fileCount > 10
                    Set oldestFile = Nothing
                    oldestDate = Now + 1 ' future date
                    For Each file In folder.Files
                        If Left$(file.Name, 13) = "Levav_backup_" And Right$(file.Name, 5) = ".xlsm" Then
                            If file.DateLastModified < oldestDate Then
                                oldestDate = file.DateLastModified
                                Set oldestFile = file
                            End If
                        End If
                    Next file
                    If Not oldestFile Is Nothing Then
                        oldestFile.Delete
                        fileCount = fileCount - 1
                    Else
                        Exit Do
                    End If
                Loop
            End If
            Set fso = Nothing
        End If
    End If
    
    ' ---- Step 2: Hide all sheets except home ----
    Dim ws As Worksheet
    Dim ctrlName As String
    ' "daf habait" = home sheet name
    ctrlName = ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)

    On Error Resume Next
    
    ' Make sure home sheet is visible
    ThisWorkbook.Worksheets(ctrlName).Visible = xlSheetVisible
    
    ' Hide everything else
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, ctrlName, vbTextCompare) <> 0 Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' Activate home sheet
    ThisWorkbook.Worksheets(ctrlName).Activate
    On Error GoTo 0
    
    ' Save the workbook with hidden state
    ThisWorkbook.Save
    
    Exit Sub

SKIP_CLOSE_ACTIONS:
    ' If any error occurs during backup/hide, just close without blocking
    Resume Next
End Sub

