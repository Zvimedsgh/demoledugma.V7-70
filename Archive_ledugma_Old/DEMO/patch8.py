import re
from datetime import datetime

file_path = r"C:\ledugma\DEMO\modDemoReports_V9.31.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Revert English Data Validation message back to Hebrew
eng_msg = '"Available in FULL version only!"\n            .Validation.ErrorTitle = "PRO VERSION"'
eng_msg2 = '"Available in FULL version only!"\n        .Validation.ErrorTitle = "PRO VERSION"'
heb_msg = 'ChrW(1508) & ChrW(1506) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1500) & ChrW(1489) & ChrW(1491)'

# Note: In the file it might be .Validation.ErrorMessage = "Available in FULL version only!" ...
# Let's use regex to replace it robustly
content = re.sub(
    r'\.Validation\.ErrorMessage = "Available in FULL version only!"\s*\.Validation\.ErrorTitle = "PRO VERSION"',
    f'.Validation.ErrorMessage = {heb_msg}',
    content
)

# Update version header
now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M:%S")

content = re.sub(r"VERSION: 9\.30 \(.*?\)", f"VERSION: 9.31 ({date_str})", content)
content = content.replace("VERSION 9.30", "VERSION 9.31")
content = content.replace("CHANGES IN 9.30:", "CHANGES IN 9.31:\n'   - Reverted Data Validation error message back to Hebrew\n' CHANGES IN 9.30:")

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Done patch8")
