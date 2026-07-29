import win32com.client
import os
import json

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')

os.makedirs('c:/ledugma/attachments', exist_ok=True)
results = []

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
                subj = getattr(item, 'Subject', '').lower()
                
                if 'b1' in subj or 'רומנית' in subj:
                    for att in item.Attachments:
                        if att.FileName.lower().endswith(('.png', '.jpg', '.gif', '.bmp', '.jpeg')):
                            continue
                        results.append({"Subject": item.Subject, "Attachment": att.FileName})
                        path = os.path.join('c:/ledugma/attachments', att.FileName)
                        try:
                            att.SaveAsFile(path)
                        except Exception as e:
                            results.append({"Error": str(e)})
    except Exception as e:
        pass

with open('c:/ledugma/attachments_res.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
