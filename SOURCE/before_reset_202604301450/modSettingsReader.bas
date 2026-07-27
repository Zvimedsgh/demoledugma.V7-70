Attribute VB_Name = "modSettingsReader"
Option Explicit

' ===== LevavChat v7 - Settings Reader =====
' Reads parameters only from Named Ranges on the workbook.
' No Hebrew sheet names in code. No fixed cell addresses.

Public Type LevavParams
    baseYear As Long
    currentYear As Long
    periodType As String
    PeriodValue As String
    DateType As String
End Type

Private Function GetNamedValue(ByVal rangeName As String) As Variant
    On Error GoTo ErrHandler
    
    GetNamedValue = ThisWorkbook.Names(rangeName).RefersToRange.Value
    Exit Function

ErrHandler:
    Err.Raise vbObjectError + 3001, _
              "GetNamedValue", _
              "Missing or invalid named range: " & rangeName
End Function

Public Function ReadLevavParams() As LevavParams
    Dim p As LevavParams
    
    p.baseYear = CLng(GetNamedValue("rngBaseYear"))
    p.currentYear = CLng(GetNamedValue("rngCurrentYear"))
    p.periodType = CStr(GetNamedValue("rngPeriodType"))
    p.PeriodValue = CStr(GetNamedValue("rngPeriodValue"))
    p.DateType = CStr(GetNamedValue("rngDateType"))
    
    ReadLevavParams = p
End Function

Public Sub Test_ReadLevavParams()
    Dim p As LevavParams
    Dim msg As String
    
    On Error GoTo ErrHandler
    
    p = ReadLevavParams()
    
    msg = "Settings check:" & vbCrLf & vbCrLf
    msg = msg & "BaseYear: " & p.baseYear & vbCrLf
    msg = msg & "CurrentYear: " & p.currentYear & vbCrLf
    msg = msg & "PeriodType: " & p.periodType & vbCrLf
    msg = msg & "PeriodValue: " & p.PeriodValue & vbCrLf
    msg = msg & "DateType: " & p.DateType & vbCrLf
    
    MsgBox msg, vbInformation, "Levav Settings"
    Exit Sub

ErrHandler:
    MsgBox "Settings error:" & vbCrLf & Err.Description, _
           vbCritical, "Levav Settings"
End Sub

