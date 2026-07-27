import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.155.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix ExportTotalChart and ExportCompCharts DisplayAlerts
content = content.replace("'Application.DisplayAlerts = False\n    tmpWs.Delete\n    'Application.DisplayAlerts = True", "Application.DisplayAlerts = False\n    tmpWs.Delete\n    Application.DisplayAlerts = True")
content = content.replace("'Application.DisplayAlerts = False\n        If Not tmpWs Is Nothing Then tmpWs.Delete\n        'Application.DisplayAlerts = True", "Application.DisplayAlerts = False\n        If Not tmpWs Is Nothing Then tmpWs.Delete\n        Application.DisplayAlerts = True")

# There might be some indentation differences, let's use regex
content = re.sub(r"'\s*Application\.DisplayAlerts = False\s*\n(\s*(?:If Not tmpWs Is Nothing Then )?tmpWs\.Delete)\s*\n\s*'\s*Application\.DisplayAlerts = True", r"Application.DisplayAlerts = False\n\1\n    Application.DisplayAlerts = True", content)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("DisplayAlerts fixed!")
