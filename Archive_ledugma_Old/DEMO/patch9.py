import re
from datetime import datetime

file_path = r"C:\ledugma\DEMO\modDemoReports_V9.32.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# We will just replace lines starting with ' MODULE: modReports until ' CHANGES IN 9.27:
now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M:%S")

new_header = f"""' MODULE: modReports' VERSION: 9.32 ({date_str})
' PURPOSE: Complete reporting system - BuildReview + ApplyCorrectionsAndBuildReports
'          VERSION 9.32
' ============================================================================
' CHANGES IN 9.32:
'   - Fixed Data Validation Msg (Hebrew restored) and ProVersion MsgBoxU
' CHANGES IN 9.28:"""

# Regex to match the exact block to replace
pattern = r"' MODULE: modReports.*?CHANGES IN 9\.28:"

if re.search(pattern, content, re.DOTALL):
    content = re.sub(pattern, new_header, content, flags=re.DOTALL)
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(content)
    print("Version header updated successfully")
else:
    print("Regex failed to find header block")
