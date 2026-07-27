import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.112.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

helpers = [
    "BuildBaseSheet",
    "BuildComparisonSheet",
    "BuildSummarySheet",
    "ExportTotalChart",
    "ExportCompCharts",
    "BuildTitleSlide",
    "BuildTotalSlideFromImage",
    "BuildChartSlide",
    "BuildTableSlide"
]

def comment_app_settings(func_body):
    lines = func_body.splitlines()
    for i in range(len(lines)):
        if re.search(r'^\s*Application\.(ScreenUpdating|EnableEvents|DisplayAlerts|Calculation)\s*=', lines[i]):
            lines[i] = "'" + lines[i] # Comment it out
    return "\n".join(lines)

for helper in helpers:
    pattern = r'(Private Sub|Public Sub)\s+' + helper + r'\b.*?(?=\n(?:Private Sub|Public Sub|Private Function|Public Function|Function|Sub)\s|\Z)'
    match = re.search(pattern, content, re.DOTALL)
    if match:
        old_body = match.group(0)
        new_body = comment_app_settings(old_body)
        content = content.replace(old_body, new_body)

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_112"', 'Attribute VB_Name = "Goren_Claude_V2_113"')
content = content.replace('VERSION: V2.112', 'VERSION: V2.113')
content = content.replace('APP_VERSION As String = "2.112"', 'APP_VERSION As String = "2.113"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.113.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.113 created. Helper macros modified.")
