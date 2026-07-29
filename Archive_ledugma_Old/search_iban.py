import win32com.client
import json

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
res = []

for f in outlook.Folders:
    for folder in f.Folders:
        try:
            for item in folder.Items:
                if getattr(item, 'Class', 0) == 43:
                    body = getattr(item, 'Body', '').lower()
                    subj = getattr(item, 'Subject', '').lower()
                    
                    if ('iban' in body or 'swift' in body or 'ro' in body or 'account' in body) and ('רומנית' in body or 'רומנית' in subj or 'b1' in body or 'b1' in subj):
                        res.append({
                            'Subject': getattr(item, 'Subject', ''),
                            'Body': getattr(item, 'Body', '')[:1000]
                        })
        except Exception:
            pass

with open('c:/ledugma/iban_results.json', 'w', encoding='utf-8') as out:
    json.dump(res, out, ensure_ascii=False, indent=2)
