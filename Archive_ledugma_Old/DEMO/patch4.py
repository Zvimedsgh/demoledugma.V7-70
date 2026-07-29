import re
from datetime import datetime

file_path = r"C:\ledugma\DEMO\modDemoReports_V9.27.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M:%S")

old_header = """' MODULE: modReports' VERSION: 7.93 (Fetch USD and EUR dynamically)
' PURPOSE: Complete reporting system - BuildReview + ApplyCorrectionsAndBuildReports
'          VERSION 7.93
' ============================================================================
' CHANGES IN 7.93:"""

new_header = f"""' MODULE: modReports' VERSION: 9.27 ({date_str})
' PURPOSE: Complete reporting system - BuildReview + ApplyCorrectionsAndBuildReports
'          VERSION 9.27
' ============================================================================
' CHANGES IN 9.27:
'   - Fixed BuildReview missing year override
'   - Added Date and Time to version header
'   - Fixed syntax error with backslash in comments
'   - Blocked UI years (2024/2025) and forced backend processing to 2019/2020
' CHANGES IN 7.93:"""

if old_header in content:
    content = content.replace(old_header, new_header)
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(content)
    print("Replaced successfully")
else:
    print("Old header not found")
