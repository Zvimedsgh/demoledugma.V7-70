Attribute VB_Name = "modLevav_clean"
Option Explicit

Public Sub RunLevav()
    NormalizeCurrentYear
    BuildReview
End Sub

Public Sub NormalizeCurrentYear()
    Dim rawYear As String
    Dim wsRaw As Worksheet, wsFixed As Worksheet
    Dim lastRow As Long, r As Long
    
    rawYear = Trim(CStr(ThisWorkbook.Names("rngCurrentYear").RefersToRange.Value)) ' 2025
    Set wsRaw = ThisWorkbook.Worksheets(rawYear)
    Set wsFixed = RecreateSheet(rawYear & "_fixed")
    
    ' fixed structure matching 2024
    wsFixed.Cells(1, "D").Value = "ID_NUMBER"
    wsFixed.Cells(1, "G").Value = "POLICY"
    wsFixed.Cells(1, "I").Value = "COMPANY_NAME"
    wsFixed.Cells(1, "K").Value = "BRANCH"
    wsFixed.Cells(1, "N").Value = "AGENT"
    wsFixed.Cells(1, "R").Value = "ACTION"
    wsFixed.Cells(1, "S").Value = "PREMIUM"
    wsFixed.Cells(1, "T").Value = "COMPANY_COMMISSION"
    
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, GetMappedColumnIndex("COMPANY_NAME")).End(xlUp).Row
    
    For r = 2 To lastRow
        wsFixed.Cells(r, "D").Value = wsRaw.Cells(r, GetMappedColumnIndex("ID_NUMBER")).Value
        wsFixed.Cells(r, "G").Value = wsRaw.Cells(r, GetMappedColumnIndex("POLICY")).Value
        wsFixed.Cells(r, "I").Value = wsRaw.Cells(r, GetMappedColumnIndex("COMPANY_NAME")).Value
        wsFixed.Cells(r, "R").Value = wsRaw.Cells(r, GetMappedColumnIndex("ACTION")).Value
        wsFixed.Cells(r, "S").Value = wsRaw.Cells(r, GetMappedColumnIndex("PREMIUM")).Value
        wsFixed.Cells(r, "T").Value = wsRaw.Cells(r, GetMappedColumnIndex("COMPANY_COMMISSION")).Value
    Next r
    
    wsFixed.DisplayRightToLeft = True
    wsFixed.Columns.AutoFit
End Sub

Public Sub BuildReview()
    Dim baseYear As String, currentYear As String
    Dim wsBase As Worksheet, wsCurrent As Worksheet
    Dim dict As Object
    
    baseYear = Trim(CStr(ThisWorkbook.Names("rngBaseYear").RefersToRange.Value))       ' 2024
    currentYear = Trim(CStr(ThisWorkbook.Names("rngCurrentYear").RefersToRange.Value)) ' 2025
    
    Set wsBase = ThisWorkbook.Worksheets(baseYear)
    Set wsCurrent = ThisWorkbook.Worksheets(currentYear & "_fixed")
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' fixed normalized structure
    modLevav_v7.ProcessCompanySheet wsBase, dict, 9, 19, 20, 4, 7, 18, 0
modLevav_v7.ProcessCompanySheet wsCurrent, dict, 9, 19, 20, 4, 7, 18, 1, 1
    
    WriteCompaniesOutput RecreateSheet("companies_comparison"), dict, baseYear, currentYear
End Sub
Private Function RecreateSheet(ByVal sheetName As String) As Worksheet
    Application.DisplayAlerts = False
    
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    
    Set RecreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    RecreateSheet.Name = sheetName
End Function
