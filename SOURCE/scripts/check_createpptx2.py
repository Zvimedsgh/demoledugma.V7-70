import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    text = f.read()

if "CreatePPTX" in text:
    print("Found CreatePPTX")
else:
    print("No CreatePPTX")

