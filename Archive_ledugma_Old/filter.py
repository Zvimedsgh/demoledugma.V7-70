import json
with open('c:/ledugma/iban_fast.json', encoding='utf-8') as f:
    d = json.load(f)
out = []
for x in d:
    if 'b1' in x['Subject'].lower() or 'רומנית' in x['Subject'].lower():
        out.append({"Subj": x['Subject'], "Body": x['Body'][:400]})
json.dump(out, open('c:/ledugma/filter_out.json','w',encoding='utf-8'), ensure_ascii=False, indent=2)
