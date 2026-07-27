import sys
import shutil

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.030.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.031.bas'

shutil.copy(filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("Goren_Claude_V2_030", "Goren_Claude_V2_031")
content = content.replace("V2.030", "V2.031")
content = content.replace('APP_VERSION As String = "2.030"', 'APP_VERSION As String = "2.031"')

old_block = """        If isDemoMode Then
            yearVal = "2025"
            On Error Resume Next
            Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)
            On Error GoTo ERR_HANDLER
            If Not wsSrc Is Nothing Then Set wbSrc = ThisWorkbook
        End If

        If wbSrc Is Nothing Then"""

new_block = """        If isDemoMode Then
            yearVal = "2025"
        End If

        On Error Resume Next
        Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)
        If Err.Number <> 0 Then
            Err.Clear
            Set wsSrc = ThisWorkbook.Worksheets(ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & "_" & yearVal)
        End If
        On Error GoTo ERR_HANDLER
        
        If Not wsSrc Is Nothing Then Set wbSrc = ThisWorkbook

        If wbSrc Is Nothing Then"""

content = content.replace(old_block, new_block)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.031 successfully")
