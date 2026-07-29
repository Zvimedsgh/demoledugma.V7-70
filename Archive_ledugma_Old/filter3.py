import json
with open('c:/ledugma/iban_fast_full.json', encoding='utf-8') as f:
    d = json.load(f)

for x in d:
    if 'b1' in x['Subject'].lower() or 'רומנית' in x['Subject'].lower():
        if '11:40' in x['Body'] or '11:18' in x['Body'] or '12:35' in x['Body'] or 'swift' in x['Body'].lower():
            with open('c:/ledugma/thread.txt', 'w', encoding='utf-8') as out:
                out.write("=====================\n")
                out.write(x['Subject'] + "\n")
                out.write(x['Body'] + "\n")
                out.write("=====================\n")
            break
