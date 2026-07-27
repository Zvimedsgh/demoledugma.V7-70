import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.103.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

def repl_slide_title(match):
    return match.group(0).replace('ChrW(1500) & ChrW(1489) & ChrW(1489)', 'GetActiveAgencyName()')

content = re.sub(r'BuildChartSlide.*?ChrW\(1489\).*?paramsSubtitle', repl_slide_title, content)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.103 updated with slide titles.")
