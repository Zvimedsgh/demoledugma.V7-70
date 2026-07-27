import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.114.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def get_function_name(idx):
    for i in range(idx, -1, -1):
        m = re.search(r'(Sub|Function)\s+([A-Za-z0-9_]+)', lines[i])
        if m:
            return m.group(2)
    return "Unknown"

line_nums = [728, 783, 1165, 1184, 1403, 1635, 1658, 2468, 2659, 2676, 2692, 4343, 4353, 4370, 4636, 4665, 4684, 5571, 5591, 5621, 5635, 5712, 5728, 5929, 6479, 6543, 7043, 7064, 7092, 7113, 7130, 7239, 7403, 7407, 8155]

funcs = set()
for l in line_nums:
    f_name = get_function_name(l-1)
    print(f"Line {l}: {f_name}")
    funcs.add(f_name)

print("Unique functions:", ", ".join(funcs))
