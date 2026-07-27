import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

print("Block 1:")
for i, line in enumerate(lines):
    if "If isDemoMode Then" in line and "BuildReview" in "".join(lines[max(0, i-50):i]):
        for j in range(i, i+15):
            print(f"{lines[j].rstrip()}")
        break

print("-" * 20)
print("Block 5:")
for i, line in enumerate(lines):
    if "tmpWs.ChartObjects.Delete" in line:
        for j in range(i-2, i+3):
            print(f"{lines[j].rstrip()}")
        break

