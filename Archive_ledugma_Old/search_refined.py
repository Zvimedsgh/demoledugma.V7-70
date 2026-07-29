import win32com.client
import pythoncom
import os

def refine_search():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    required = ["רומנית"]
    payment_terms = ["תשלום", "העברה", "חשבון", "בנק", "שקלים", "לשלם", "פרטי"]
    
    results = []
    folders = []
    try:
        for account in outlook.Folders:
            for folder in account.Folders:
                folders.append(folder)
    except:
        pass
        
    for folder in folders:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            count = 0
            for msg in items:
                count += 1
                if count > 5000: break
                if msg.Class != 43: continue
                
                subject = getattr(msg, 'Subject', '') or ''
                sender = getattr(msg, 'SenderName', '') or ''
                body = getattr(msg, 'Body', '') or ''
                to_field = getattr(msg, 'To', '') or ''
                text = (subject + " " + sender + " " + body).lower()
                
                has_req = any(req in text for req in required)
                has_pay = any(pay in text for pay in payment_terms)
                
                if has_req and has_pay:
                    snippet = body[:250].replace('\n', ' ')
                    results.append({
                        "EntryID": getattr(msg, 'EntryID', ''),
                        "Date": str(getattr(msg, 'ReceivedTime', ''))[:16],
                        "Sender": sender,
                        "To": to_field,
                        "Subject": subject,
                        "Snippet": snippet
                    })
        except:
            pass
            
    html = f"""<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>תוצאות חיפוש: תשלום קורס רומנית</title>
    <style>
        body {{ font-family: Arial; padding: 20px; background: #f0f2f5; }}
        .email {{ background: white; padding: 15px; margin-bottom: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .btn {{ background: #4F46E5; color: white; padding: 8px 15px; border: none; border-radius: 4px; cursor: pointer; text-decoration: none; display: inline-block; }}
        .btn:hover {{ background: #4338CA; }}
        .snippet {{ color: #666; font-size: 0.9em; margin: 10px 0; background: #fafafa; padding: 10px; border-radius: 4px; border: 1px solid #eee; }}
    </style>
</head>
<body>
    <h2>מצאנו {len(results)} מיילים רלוונטיים:</h2>
"""
    for r in results:
        html += f"""
    <div class="email">
        <strong>מאת:</strong> {r['Sender']} &nbsp;|&nbsp; 
        <strong>אל:</strong> {r['To']} &nbsp;|&nbsp; 
        <strong>תאריך:</strong> <span dir="ltr">{r['Date']}</span><br>
        <strong>נושא:</strong> {r['Subject']} <br>
        <p class="snippet">{r['Snippet']}...</p>
        <a class="btn" href="http://127.0.0.1:5004/open/{r['EntryID']}" target="_blank">👁️ פתח באאוטלוק</a>
    </div>
"""
    html += """
</body></html>"""

    with open("c:/ledugma/search_results.html", "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == '__main__':
    refine_search()
