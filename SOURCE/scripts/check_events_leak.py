import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.107.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

funcs = re.findall(r'(Public Sub \w+\(.*?\).*?End Sub)', content, re.DOTALL)
for f in funcs:
    if "Application.EnableEvents = False" in f:
        # Check if there is an Exit Sub
        exit_subs = re.finditer(r'\n[ \t\d]*Exit Sub', f)
        for es in exit_subs:
            # Check if there is EnableEvents = True before it
            idx = es.start()
            prev_code = f[:idx]
            # Check the last 15 lines before Exit Sub
            last_lines = prev_code.splitlines()[-15:]
            if not any("EnableEvents = True" in line for line in last_lines):
                name = re.search(r'Public Sub (\w+)', f).group(1)
                print(f"Missing EnableEvents=True in {name} before Exit Sub!")
                for l in last_lines[-5:]:
                    print("   ", l.strip())

