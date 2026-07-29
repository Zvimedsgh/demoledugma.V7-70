import re
from datetime import datetime

file_path = r"C:\ledugma\DEMO\modDemoReports_V7.95.bas"
out_path = r"C:\ledugma\DEMO\modDemoReports_V7.96.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Fix the syntax error: replacing "\'" with "'"
content = content.replace(r"\'", "'")

# Update version header and add Date/Time
now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M:%S")

# Replace version and add timestamp
content = content.replace("VERSION: 7.95", f"VERSION: 7.96 ({date_str})")
content = content.replace("CHANGES IN 7.95:", f"CHANGES IN 7.96:\n'   - Fixed syntax error with backslash in comments\n'   - Added Date and Time to version header\n' CHANGES IN 7.95:")

with open(out_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Fixed")
