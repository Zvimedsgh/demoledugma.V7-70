import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.103.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern = re.compile(r'Public Sub BuildPresentation\(\).*?End Sub', re.DOTALL)
match = pattern.search(content)
if match:
    sub_content = match.group(0)
    for i, line in enumerate(sub_content.splitlines()):
        if "levavName =" in line or "GetActiveAgencyName" in line:
            print(f"[{i+1}] {line.strip()}")
            
pattern2 = re.compile(r'Private Sub BuildTitleSlide\(.*?End Sub', re.DOTALL)
match2 = pattern2.search(content)
if match2:
    sub_content2 = match2.group(0)
    for i, line in enumerate(sub_content2.splitlines()):
        if "GetActiveAgencyName" in line or "ChrW" in line:
            print(f"[TitleSlide {i+1}] {line.strip()}")

