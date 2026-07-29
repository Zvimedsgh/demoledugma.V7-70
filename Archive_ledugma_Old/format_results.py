import json

with open('c:/ledugma/search_results.json', 'r', encoding='utf-8') as f:
    results = json.load(f)

md_content = '# תוצאות חיפוש: קורס רומנית / פינקו ברקן\n\n'
md_content += f'נמצאו בסך הכל **{len(results)}** מיילים ברחבי התיבה והארכיון שלך.\n\n'
md_content += '| תאריך | שולח | נושא | תיקייה |\n'
md_content += '|---|---|---|---|\n'

for r in results:
    date_short = r['Date'][:16]
    subject = r.get('Subject', '').replace('|', '-')
    sender = r.get('Sender', '').replace('|', '-')
    md_content += f"| {date_short} | {sender} | {subject} | {r.get('Folder', '')} |\n"

artifact_path = r'C:\Users\Zvi\.gemini\antigravity\brain\a19cc612-509d-4585-8e80-d112928a73e4\search_results.md'
with open(artifact_path, 'w', encoding='utf-8') as f:
    f.write(md_content)
print('Artifact created.')
