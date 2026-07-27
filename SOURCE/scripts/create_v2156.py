import sys
import shutil

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.155.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.156.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update version to 2.156
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_155"', 'Attribute VB_Name = "Goren_Claude_V2_156"')
content = content.replace('VERSION: V2.155', 'VERSION: V2.156')
content = content.replace('APP_VERSION As String = "2.155"', 'APP_VERSION As String = "2.156"')

# 2. Fix the MsgBox gibberish
# Line 4766: "- " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " 'כן' (Yes) " & ChrW(1500) & ...
# Line 4767: "- " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " 'לא' (No) " & ChrW(1500) & ...
# Line 4768: "- " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " 'ביטול' (Cancel) " & ChrW(1499) & ...
# 'כן' = Chr(39) & ChrW(1499) & ChrW(1503) & Chr(39)
# 'לא' = Chr(39) & ChrW(1500) & ChrW(1488) & Chr(39)
# 'ביטול' = Chr(39) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500) & Chr(39)

old_yes = "\" 'כן' (Yes) \""
new_yes = " Chr(39) & ChrW(1499) & ChrW(1503) & Chr(39) & \" (Yes) \""
content = content.replace(old_yes, new_yes)

old_no = "\" 'לא' (No) \""
new_no = " Chr(39) & ChrW(1500) & ChrW(1488) & Chr(39) & \" (No) \""
content = content.replace(old_no, new_no)

old_cancel = "\" 'ביטול' (Cancel) \""
new_cancel = " Chr(39) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500) & Chr(39) & \" (Cancel) \""
content = content.replace(old_cancel, new_cancel)

# 3. Stop deleting result sheets at the end of Button 3
# Lines 4801-4807
cleanup_block = """DeleteSheetIfExists SHEET_COMPANIES()
DeleteSheetIfExists SHEET_BRANCH()
DeleteSheetIfExists SHEET_MAINBRANCH()
DeleteSheetIfExists SHEET_TELLERS()
DeleteSheetIfExists SHEET_AGENTS()
DeleteSheetIfExists SHEET_MONTHS()
DeleteSheetIfExists SHEET_SUMMARY()"""

content = content.replace(cleanup_block, "' (REMOVED: User requested that result sheets are NOT deleted after presentation)")

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.156 successfully!")
