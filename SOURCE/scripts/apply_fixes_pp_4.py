import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.121.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    # Only remove the specific lines at the old location
    if i == 4707 and "910     Set ppPres = Nothing" in line:
        continue
    if i == 4708 and "920     Set ppApp = Nothing" in line:
        continue
    new_lines.append(line)

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Removed duplicate Set Nothing lines.")
