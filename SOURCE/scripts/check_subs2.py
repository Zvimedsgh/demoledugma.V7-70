import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.099.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern = re.compile(r'Public Sub SendForReview\(.*?(?=Public Sub|End Sub).*?End Sub', re.DOTALL)
match = pattern.search(content)
if match:
    print("SendForReview found. Length:", len(match.group(0)))

pattern = re.compile(r'Public Sub ApplyCorrectionsAndBuildReports\(.*?(?=Public Sub|End Sub).*?End Sub', re.DOTALL)
match = pattern.search(content)
if match:
    print("ApplyCorrectionsAndBuildReports found. Length:", len(match.group(0)))
