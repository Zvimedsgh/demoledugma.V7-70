import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    text = f.read()

def find_block(text, search):
    idx = text.find(search)
    if idx == -1:
        print(f"NOT FOUND: {search[:50]}")
    else:
        print(f"FOUND: {search[:50]}")

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
find_block(text, old_block_1)

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
find_block(text, old_block_2)

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
find_block(text, old_block_3)

old_block_4 = """chartItems = nItems
220     If chartItems > 15 Then chartItems = 15"""
find_block(text, old_block_4)

old_block_5 = """If imgDocs <> "" Then
tmpWs.ChartObjects.Delete"""
find_block(text, old_block_5)

old_block_6 = """If imgIns <> "" Then
tmpWs.ChartObjects.Delete"""
find_block(text, old_block_6)

old_block_7 = """If sName = H_SET_PARAMS() Then IsProtectedSheet = True: Exit Function"""
find_block(text, old_block_7)
