import win32com.client
import json

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
res = []

for folder_id in [6, 5]:
    try:
        folder = outlook.GetDefaultFolder(folder_id)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        
        count = 0
        for item in items:
            count += 1
            if count > 2000: break
            
            if getattr(item, 'Class', 0) == 43:
                body = getattr(item, 'Body', '').lower()
                subj = getattr(item, 'Subject', '').lower()
                
                if ('רומנית' in body or 'רומנית' in subj or 'b1' in body or 'b1' in subj):
                    res.append({
                        'Subject': getattr(item, 'Subject', ''),
                        'Body': getattr(item, 'Body', '')
                    })
    except Exception as e:
        pass

with open('c:/ledugma/iban_fast_full.json', 'w', encoding='utf-8') as out:
    json.dump(res, out, ensure_ascii=False, indent=2)
