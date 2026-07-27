import sys
import shutil

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.030.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# I want to change the src file logic in BuildReview:
old_logic = """        If isDemoMode Then
            yearVal = "2025"
            On Error Resume Next
            Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)
            On Error GoTo ERR_HANDLER
        End If

        If wsSrc Is Nothing Then
            srcPath = FindSourceFile(yearVal)
            If srcPath = "" Then Err.Raise vbObjectError + 1003, "BuildReview", "SOURCE FILE NOT FOUND FOR YEAR: " & yearVal
            
            Application.DisplayAlerts = True
            Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
            Application.DisplayAlerts = False
            Set wsSrc = OpenDataSheet(wbSrc)
        End If"""

# Wait, the exact text in V2.029 might be slightly different.
# Let's write a script to find the exact block and replace it.
