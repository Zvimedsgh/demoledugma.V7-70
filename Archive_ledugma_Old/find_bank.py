import re
import json

lines = []
with open(r'C:\Users\Zvi\.gemini\antigravity\brain\a19cc612-509d-4585-8e80-d112928a73e4\.system_generated\steps\822\content.md', encoding='utf-8') as f:
    for line in f:
        if re.search(r'(iban|swift|ro[0-9]{2}|חשבון|בנק)', line, re.IGNORECASE):
            lines.append(line.strip())

with open(r'c:\ledugma\bank_out.json', 'w', encoding='utf-8') as out:
    json.dump(lines, out, ensure_ascii=False, indent=2)
