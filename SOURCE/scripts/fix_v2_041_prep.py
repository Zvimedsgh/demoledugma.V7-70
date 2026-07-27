import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040_20260702_1507.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Fix Default DEMO_MODE to YES (כן)
old_demo_default = """wsParams.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "DEMO_MODE"
2030  wsParams.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = ChrW(1500) & ChrW(1488) 'לא"""
new_demo_default = """wsParams.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "DEMO_MODE"
2030  wsParams.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = ChrW(1499) & ChrW(1503) 'כן"""
# Wait, let me check the exact lines for demo default
