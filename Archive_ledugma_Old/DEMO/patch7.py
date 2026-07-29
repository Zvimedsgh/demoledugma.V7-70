import re
from datetime import datetime

file_path = r"C:\ledugma\DEMO\modDemoReports_V9.30.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Replace MsgBox with MsgBoxU in ProVersionOnly
old_msgbox = 'MsgBox ChrW(1508) & ChrW(1506)'
new_msgbox = 'MsgBoxU ChrW(1508) & ChrW(1506)'
content = content.replace(old_msgbox, new_msgbox)

# Update version header
now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M:%S")

content = re.sub(r"VERSION: 9\.29 \(.*?\)", f"VERSION: 9.30 ({date_str})", content)
content = content.replace("VERSION 9.29", "VERSION 9.30")
content = content.replace("CHANGES IN 9.29:", "CHANGES IN 9.30:\n'   - Changed MsgBox to MsgBoxU in ProVersionOnly to fix question marks\n' CHANGES IN 9.29:")

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Done patch7")
