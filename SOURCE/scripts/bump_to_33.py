import sys
import shutil

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.032.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.033.bas'

shutil.copy(filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("Goren_Claude_V2_032", "Goren_Claude_V2_033")
content = content.replace("V2.032", "V2.033")
content = content.replace('APP_VERSION As String = "2.032"', 'APP_VERSION As String = "2.033"')

old_func = """Private Function IsParameterSheet(ByVal sName As String) As Boolean
    Dim shtNames As Variant
    shtNames = Array(H_SET_FIELDMAP(), H_SET_BRANCHES(), H_SET_PARAMS(), H_SET_PERIODS(), H_SET_MESSAGES(), H_SET_PERMISSIONS(), H_SET_REASONS(), H_SET_CLIENTS())
    Dim i As Long
    For i = LBound(shtNames) To UBound(shtNames)
        If StrComp(sName, shtNames(i), vbTextCompare) = 0 Then
            IsParameterSheet = True
            Exit Function
        End If
    Next i
    IsParameterSheet = False
End Function"""

new_func = """Private Function IsParameterSheet(ByVal sName As String) As Boolean
    ' Safely checks if the sheet name starts with "הגדרות_"
    Dim prefix As String
    prefix = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514) & "_"
    If InStr(1, Trim$(sName), prefix, vbTextCompare) = 1 Then
        IsParameterSheet = True
    Else
        IsParameterSheet = False
    End If
End Function"""

content = content.replace(old_func, new_func)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.033 successfully")
