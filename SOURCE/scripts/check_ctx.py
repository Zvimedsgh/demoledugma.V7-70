import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.112.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def print_context(line_num):
    print(f"Context for line {line_num}:")
    for i in range(max(0, line_num-5), min(len(lines), line_num+6)):
        print(f"[{i+1}] {lines[i].strip()}")
    print("-" * 20)

print_context(2371)
print_context(2522)
