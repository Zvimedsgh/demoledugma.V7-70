lines = open(r'C:\Users\Zvi\.gemini\antigravity\brain\a19cc612-509d-4585-8e80-d112928a73e4\.system_generated\steps\822\content.md', encoding='utf-8').read().splitlines()
out = []
for line in lines:
    if any(word in line.lower() for word in ['pdf', 'doc', 'קובץ', 'מסמך', 'download', 'href', 'jotform', 'attachment', 'form', 'invoice', 'חשבונית', 'קבלה', 'proforma']):
        out.append(line.strip())
import json
json.dump(out, open('c:/ledugma/site_docs_out.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
