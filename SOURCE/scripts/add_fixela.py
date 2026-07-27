import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.108.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Insert ResetEvents at the top
insert_idx = content.find("Public Sub NavToIndex()")
if insert_idx != -1:
    content = content[:insert_idx] + """Public Sub ResetEvents()
    ' A utility macro ("Fixela") to reset application state if it ever gets stuck
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBoxU "המערכת שוחררה בהצלחה (Fixela)", vbInformation
End Sub

""" + content[insert_idx:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.108.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.108 updated with ResetEvents.")
