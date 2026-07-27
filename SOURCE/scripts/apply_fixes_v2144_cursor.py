import sys, re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern_goto = re.compile(r"(' Select first result\s*Application\.Goto wsSearch\.Cells\(4, 1\))")

if pattern_goto.search(content):
    content = pattern_goto.sub(r"' Stay in search cell\nApplication.Goto wsSearch.Cells(2, 1)", content)
    print("Updated cursor to stay in A2.")
else:
    print("Could not find Goto 4,1.")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
