import json
import re
with open('c:/ledugma/iban_fast_full.json', encoding='utf-8') as f:
    d = json.load(f)
out = []
for x in d:
    if 'b1' in x['Subject'].lower() or 'רומנית' in x['Subject'].lower():
        body = x['Body']
        match = re.search(r'(.{0,100})(iban|swift|ro[0-9a-z]{10,20}|account|בנק)(.{0,300})', body.lower(), re.DOTALL)
        if match:
            idx = body.lower().find(match.group(2))
            start = max(0, idx - 100)
            end = min(len(body), idx + 300)
            out.append({"Subj": x['Subject'], "Snippet": body[start:end]})
json.dump(out, open('c:/ledugma/filter2_out.json','w',encoding='utf-8'), ensure_ascii=False, indent=2)
