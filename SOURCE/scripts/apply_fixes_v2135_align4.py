import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.135.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "shpDemoMsg.TextFrame2.TextRange.Font.Size = 17" in line and "3688" not in line:
        # Wait, the lines above show it IS [3688] shpDemoMsg.TextFrame2.TextRange.Font.Size = 17
        pass

    if "shpDemoMsg.TextFrame2.TextRange.Font.Size = 17" in line:
        # Check if next line is already alignment
        if "Alignment = 2" not in lines[i+1]:
            lines.insert(i+1, "shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter\n")

    if "3687 shpDemoMsg" in line and not line.startswith("3687 "):
        # fix the weird line numbering prefix
        lines[i] = line.replace("3687 shpDemoMsg", "shpDemoMsg")

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

