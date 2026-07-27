import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.132.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "3970 wsMain.Range(\"F19\").Value" in line:
        lines[i] = "3970 wsMain.Range(\"F19\").ClearContents\n3975 wsMain.Range(\"A22\").Value = ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & \" v\" & APP_VERSION\n"
    elif "3980 wsMain.Range(\"F19\").Font.Size" in line:
        lines[i] = "3980 wsMain.Range(\"A22\").Font.Size = 10\n"
    elif "3990 wsMain.Range(\"F19\").Font.Color" in line:
        lines[i] = "3990 wsMain.Range(\"A22\").Font.Color = RGB(150, 150, 150)\n"
    elif "4000 wsMain.Range(\"F19\").Font.Bold" in line:
        lines[i] = "4000 wsMain.Range(\"A22\").Font.Bold = False\n"
    elif "4010 wsMain.Range(\"F19\").HorizontalAlignment" in line:
        lines[i] = "4010 wsMain.Range(\"A22\").HorizontalAlignment = -4108\n"

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Flexible credit replacement done.")
