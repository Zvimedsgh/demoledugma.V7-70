import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.052.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    text = f.read()

# Block 1: BuildReview
old_block_1 = """If isDemoMode Then
On Error Resume Next
Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)
On Error GoTo ERR_HANDLER
If Not wsSrc Is Nothing Then Set wbSrc = ThisWorkbook
End If

If wbSrc Is Nothing Then
srcPath = FindSourceFile(yearVal)
540         If srcPath = "" Then Err.Raise vbObjectError + 1003, "BuildReview", "SOURCE FILE NOT FOUND FOR YEAR: " & yearVal

Application.DisplayAlerts = True
550         Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
Application.DisplayAlerts = False
560         Set wsSrc = OpenDataSheet(wbSrc)
End If"""

new_block_1 = """If isDemoMode Then
On Error Resume Next
Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)
On Error GoTo ERR_HANDLER
If wsSrc Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & yearVal & " not found!"
Set wbSrc = ThisWorkbook
Else
srcPath = FindSourceFile(yearVal)
540         If srcPath = "" Then Err.Raise vbObjectError + 1003, "BuildReview", "SOURCE FILE NOT FOUND FOR YEAR: " & yearVal

Application.DisplayAlerts = True
550         Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
Application.DisplayAlerts = False
560         Set wsSrc = OpenDataSheet(wbSrc)
End If"""

# Block 2: ApplyCorrectionsAndBuildReports (yearVal)
old_block_2 = """If isDemoMode Then
On Error Resume Next
Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)
On Error GoTo ERR_HANDLER
If Not wsSrc Is Nothing Then Set wbSrc = ThisWorkbook
End If

If wbSrc Is Nothing Then
1050        srcPath = FindSourceFile(yearVal)
1060        If srcPath = "" Then Err.Raise vbObjectError + 2002, "ApplyCorrections", "SOURCE NOT FOUND: " & yearVal

1070        debugStep = "OPEN_SRC"
1080        Set wbSrc = SafeOpenWorkbook(srcPath)
If wbSrc Is Nothing Then Err.Raise 424, "OPEN_SRC", "Failed to open source file. Excel might be blocking it."
1090        Set wsSrc = OpenDataSheet(wbSrc)
End If"""

new_block_2 = """If isDemoMode Then
On Error Resume Next
Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)
On Error GoTo ERR_HANDLER
If wsSrc Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & yearVal & " not found!"
Set wbSrc = ThisWorkbook
Else
1050        srcPath = FindSourceFile(yearVal)
1060        If srcPath = "" Then Err.Raise vbObjectError + 2002, "ApplyCorrections", "SOURCE NOT FOUND: " & yearVal

1070        debugStep = "OPEN_SRC"
1080        Set wbSrc = SafeOpenWorkbook(srcPath)
If wbSrc Is Nothing Then Err.Raise 424, "OPEN_SRC", "Failed to open source file. Excel might be blocking it."
1090        Set wsSrc = OpenDataSheet(wbSrc)
End If"""


# Block 3: ApplyCorrectionsAndBuildReports (refYear)
old_block_3 = """If isDemoMode Then
On Error Resume Next
Set wsRef = ThisWorkbook.Worksheets("DATA_" & refYear)
On Error GoTo ERR_HANDLER
If Not wsRef Is Nothing Then Set wbRef = ThisWorkbook
End If

If wbRef Is Nothing Then
1100        refPath = FindSourceFile(refYear)
1110        If refPath = "" Then Err.Raise vbObjectError + 2003, "ApplyCorrections", "REF SOURCE NOT FOUND: " & refYear

1120        debugStep = "OPEN_REF"
1130        Set wbRef = SafeOpenWorkbook(refPath)
If wbRef Is Nothing Then Err.Raise 424, "OPEN_REF", "Failed to open ref file. Excel might be blocking it."
1140        Set wsRef = OpenDataSheet(wbRef)
End If"""

new_block_3 = """If isDemoMode Then
On Error Resume Next
Set wsRef = ThisWorkbook.Worksheets("DATA_" & refYear)
On Error GoTo ERR_HANDLER
If wsRef Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & refYear & " not found!"
Set wbRef = ThisWorkbook
Else
1100        refPath = FindSourceFile(refYear)
1110        If refPath = "" Then Err.Raise vbObjectError + 2003, "ApplyCorrections", "REF SOURCE NOT FOUND: " & refYear

1120        debugStep = "OPEN_REF"
1130        Set wbRef = SafeOpenWorkbook(refPath)
If wbRef Is Nothing Then Err.Raise 424, "OPEN_REF", "Failed to open ref file. Excel might be blocking it."
1140        Set wsRef = OpenDataSheet(wbRef)
End If"""

# Block 4: chartItems
old_block_4 = """chartItems = nItems
220     If chartItems > 15 Then chartItems = 15"""

new_block_4 = """chartItems = nItems
    If chartItems = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
220     If chartItems > 15 Then chartItems = 15"""

# Block 5: Chart delete
old_block_5 = """If imgDocs <> "" Then
tmpWs.ChartObjects.Delete"""

new_block_5 = """If imgDocs <> "" Then
On Error Resume Next
tmpWs.ChartObjects.Delete
On Error GoTo ERR_HANDLER"""

# Block 6: Chart delete 2
old_block_6 = """If imgIns <> "" Then
tmpWs.ChartObjects.Delete"""

new_block_6 = """If imgIns <> "" Then
On Error Resume Next
tmpWs.ChartObjects.Delete
On Error GoTo ERR_HANDLER"""

text = text.replace(old_block_1, new_block_1)
text = text.replace(old_block_2, new_block_2)
text = text.replace(old_block_3, new_block_3)
text = text.replace(old_block_4, new_block_4)
text = text.replace(old_block_5, new_block_5)
text = text.replace(old_block_6, new_block_6)

text = text.replace("V2_052", "V2_057")
text = text.replace("2.052", "2.057")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.057.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print("Created V2.057")
