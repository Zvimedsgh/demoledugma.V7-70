import re
from datetime import datetime

file_path = r"C:\ledugma\DEMO\modDemoReports_V9.27.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Fix BuildReview where yearVal is fetched using ThisWorkbook.Worksheets...
content = content.replace(
    '510     yearVal = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("rngCurrentYear").Value2))',
    '510     yearVal = "2020" \' Forced for demo\n        \' yearVal = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("rngCurrentYear").Value2))'
)

# Update version header
now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M:%S")
content = re.sub(r'VERSION: 7\.96 \(.*?\)', f'VERSION: 9.27 ({date_str})', content)
content = content.replace('VERSION 7.96', 'VERSION 9.27')
content = content.replace('CHANGES IN 7.96:', f'CHANGES IN 9.27:\n\'   - Fixed BuildReview missing year override\n\' CHANGES IN 7.96:')

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Done")
