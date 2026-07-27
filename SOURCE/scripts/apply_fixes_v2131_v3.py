import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.131.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range(\"D15\").Left, wsMain.Range(\"D15\").Top, 350, 40)" in line:
        lines[i] = line.replace("350", "500")
        print("Replaced 350 with 500")
    if "shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500)" in line and "054-6677396" in line and not "- 054-6677396" in line:
        lines[i] = line.replace("054-6677396", "- 054-6677396")
        print("Replaced whatsapp text")

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

