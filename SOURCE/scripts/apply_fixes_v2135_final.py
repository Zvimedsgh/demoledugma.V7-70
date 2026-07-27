import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.135.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "shpDemoMsg.TextFrame2.TextRange.Text =" in line and "Whatsapp" in line:
        lines[i] = "3687 shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & \" \" & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & \" \" & ChrW(1513) & ChrW(1500) & ChrW(1495) & \" Whatsapp \" & ChrW(8207) & ChrW(1500) & \"-054-6677396\"\n"
        print("Updated WhatsApp string")
    if "3688 shpDemoMsg.TextFrame2.TextRange.Font.Size = 17" in line:
        lines.insert(i+1, "36885 shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter\n")
        print("Added center alignment")
        break

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

