import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.114.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

helpers = [
    "WriteComparisonHeaders",
    "LoadCheckedFields",
    "BuildArrays"
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
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_114"', 'Attribute VB_Name = "Goren_Claude_V2_115"')
content = content.replace('VERSION: V2.114', 'VERSION: V2.115')
content = content.replace('APP_VERSION As String = "2.114"', 'APP_VERSION As String = "2.115"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.115.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.115 created. Fixed remaining helper macros.")
