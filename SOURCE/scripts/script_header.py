import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.016.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = re.sub(r'Attribute VB_Name = "Goren_Claude2_\d+"', 'Attribute VB_Name = "Goren_Claude2_016"', content)
content = re.sub(r"' VERSION: V2\.\d+", "' VERSION: V2.016", content)
content = re.sub(r"' DATE: \d{4}-\d{2}-\d{2} \d{2}:\d{2}", "' DATE: 2026-07-01 16:53", content)
content = re.sub(r'Private Const APP_VERSION As String = "2\.\d+"', 'Private Const APP_VERSION As String = "2.016"', content)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
print("Updated header to V2.016")
