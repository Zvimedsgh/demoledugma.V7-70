import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.029.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.030.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace("Goren_Claude_V2_029", "Goren_Claude_V2_030")
content = content.replace("V2.029", "V2.030")
content = content.replace('APP_VERSION As String = "2.029"', 'APP_VERSION As String = "2.030"')

old_func = """Private Function IsProtectedSheet(ByVal sName As String, ByVal refBaseKeep As String, ByVal refBaseKeepOld As String) As Boolean
If sName = CONTROL_SHEET_NAME() Then IsProtectedSheet = True: Exit Function
If sName = MANAGEMENT_SHEET_NAME() Then IsProtectedSheet = True: Exit Function"""

new_func = """Private Function IsProtectedSheet(ByVal sName As String, ByVal refBaseKeep As String, ByVal refBaseKeepOld As String) As Boolean
If sName = CONTROL_SHEET_NAME() Then IsProtectedSheet = True: Exit Function
If sName = MANAGEMENT_SHEET_NAME() Then IsProtectedSheet = True: Exit Function
If UCase$(Left$(sName, 5)) = "DATA_" Then IsProtectedSheet = True: Exit Function
If InStr(1, sName, ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & "_", vbTextCompare) = 1 Then IsProtectedSheet = True: Exit Function ' "נתונים_"
"""

content = content.replace(old_func, new_func)

# Also wait, does it protect REVIEW_ and BASE_ from other years?
# The user might have basis_2024 and basis_2025, and BuildReview deletes ALL of them except refBaseKeep?
# Yes! Because we only want ONE current base and ONE ref base?
# Wait! In Button 2, ApplyCorrections reads TWO base sheets!
# If Button 1 deletes basis_2024 when generating basis_2025, Button 2 WON'T FIND basis_2024!
# Let me look at BuildReview line 739 again:
# refBaseKeep = H_BASE() & "_" & refYearStr
# Yes! It keeps basis_2024! But what if the user wants to run on 2026 vs 2025?
# If Button 1 is run on 2026, it keeps basis_2025!
# That's perfectly correct.
# BUT wait! If Button 1 is run on 2025, it creates basis_2025, and keeps basis_2024.
# Then Button 2 runs on 2025/2024 and reads basis_2025 and basis_2024.
# This works PERFECTLY.

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.030 successfully")
