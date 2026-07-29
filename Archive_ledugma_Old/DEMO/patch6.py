import re
from datetime import datetime

file_path = r"C:\ledugma\DEMO\modDemoReports_V9.29.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Replace Data Validation error message from Hebrew ChrW to English string
old_err_msg = 'ChrW(1508) & ChrW(1506) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1500) & ChrW(1489) & ChrW(1491)'
new_err_msg = '"Available in FULL version only!"\n            .Validation.ErrorTitle = "PRO VERSION"'
new_err_msg2 = '"Available in FULL version only!"\n        .Validation.ErrorTitle = "PRO VERSION"'

# It appears twice (in SetupMainSheet and in ApplyDemoRestrictions)
# Wait, the indentation is different. I'll just replace the string.
content = content.replace(f'.Validation.ErrorMessage = {old_err_msg}', f'.Validation.ErrorMessage = {new_err_msg2}')

# Update version header
now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M:%S")

content = re.sub(r"VERSION: 9\.28 \(.*?\)", f"VERSION: 9.29 ({date_str})", content)
content = content.replace("VERSION 9.28", "VERSION 9.29")
content = content.replace("CHANGES IN 9.28:", "CHANGES IN 9.29:\n'   - Changed Data Validation error message to English to prevent question marks\n' CHANGES IN 9.28:")

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Done patch6")
