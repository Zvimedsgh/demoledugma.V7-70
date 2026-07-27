' ============================================================================
' install_modLevav.vbs
' Automatically replaces the modLevav module in LevavClaude.xlsm
' with the fixed version from modLevav.bas (all in the same folder).
'
' Usage: Double-click this file. Excel will open silently, swap the code,
'        save the workbook, and close.
'
' Requirements:
'   1. Excel must be installed.
'   2. "Trust access to the VBA project object model" must be enabled in Excel:
'      File > Options > Trust Center > Trust Center Settings > Macro Settings
'      > check "Trust access to the VBA project object model"
' ============================================================================

Option Explicit

Dim fso, scriptDir, xlsmPath, basPath
Dim xlApp, wb, vbProj, comp, i, foundIdx
Dim errMsg

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
xlsmPath = fso.BuildPath(scriptDir, "LevavClaude.xlsm")
basPath  = fso.BuildPath(scriptDir, "modLevav.bas")

If Not fso.FileExists(xlsmPath) Then
    MsgBox "Cannot find LevavClaude.xlsm in: " & scriptDir, vbCritical, "Install modLevav"
    WScript.Quit 1
End If
If Not fso.FileExists(basPath) Then
    MsgBox "Cannot find modLevav.bas in: " & scriptDir, vbCritical, "Install modLevav"
    WScript.Quit 1
End If

On Error Resume Next
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Cannot start Excel: " & Err.Description, vbCritical, "Install modLevav"
    WScript.Quit 1
End If
On Error Goto 0

xlApp.Visible = False
xlApp.DisplayAlerts = False
xlApp.AutomationSecurity = 3  ' msoAutomationSecurityForceDisable - no macros at open

On Error Resume Next
Set wb = xlApp.Workbooks.Open(xlsmPath)
If Err.Number <> 0 Then
    errMsg = "Cannot open workbook: " & Err.Description
    xlApp.Quit
    MsgBox errMsg, vbCritical, "Install modLevav"
    WScript.Quit 1
End If

Set vbProj = wb.VBProject
If Err.Number <> 0 Then
    errMsg = "Cannot access VBA project." & vbCrLf & vbCrLf & _
             "Enable: File > Options > Trust Center > Trust Center Settings > " & _
             "Macro Settings > 'Trust access to the VBA project object model'." & vbCrLf & vbCrLf & _
             "Error: " & Err.Description
    wb.Close False
    xlApp.Quit
    MsgBox errMsg, vbCritical, "Install modLevav"
    WScript.Quit 1
End If
On Error Goto 0

' Remove old modLevav if it exists
foundIdx = 0
For i = vbProj.VBComponents.Count To 1 Step -1
    Set comp = vbProj.VBComponents(i)
    If LCase(comp.Name) = "modlevav" Then
        vbProj.VBComponents.Remove comp
        foundIdx = 1
        Exit For
    End If
Next

' Import the new module from the .bas file
On Error Resume Next
vbProj.VBComponents.Import basPath
If Err.Number <> 0 Then
    errMsg = "Import failed: " & Err.Description
    wb.Close False
    xlApp.Quit
    MsgBox errMsg, vbCritical, "Install modLevav"
    WScript.Quit 1
End If
On Error Goto 0

' Save as xlsm (file format 52 = xlOpenXMLWorkbookMacroEnabled)
On Error Resume Next
wb.Save
If Err.Number <> 0 Then
    errMsg = "Save failed: " & Err.Description
    wb.Close False
    xlApp.Quit
    MsgBox errMsg, vbCritical, "Install modLevav"
    WScript.Quit 1
End If
On Error Goto 0

wb.Close True
xlApp.Quit

Set comp = Nothing
Set vbProj = Nothing
Set wb = Nothing
Set xlApp = Nothing

If foundIdx = 1 Then
    MsgBox "Done. modLevav was replaced with the fixed version in LevavClaude.xlsm.", _
           vbInformation, "Install modLevav"
Else
    MsgBox "Done. modLevav was added to LevavClaude.xlsm (no previous module with that name was found).", _
           vbInformation, "Install modLevav"
End If
