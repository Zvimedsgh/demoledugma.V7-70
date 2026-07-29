with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
subs = ['ExportTotalChart', 'ExportCompCharts', 'BuildTitleSlide', 'BuildTotalSlideFromImage', 'BuildChartSlide', 'BuildTableSlide']

for sub in subs:
    match = re.search(fr'(Sub {sub}\b.*?)End Sub', text, re.DOTALL | re.IGNORECASE)
    if match:
        sub_text = match.group(1)
        if 'ThisWorkbook.Protect' in sub_text:
            print(f"{sub} HAS Protect!")
            for line in sub_text.split('\n'):
                if 'ThisWorkbook.Protect' in line:
                    print(f"  {line.strip()}")
        else:
            print(f"{sub} is safe.")
