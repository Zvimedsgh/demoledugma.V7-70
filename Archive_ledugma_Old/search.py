import win32com.client
import pythoncom
import json

def search_emails():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # Terms to search for
    terms = ["רומנית", "פינקו", "ברקן"]
    
    results = []
    
    # Try to find all main folders in all accounts
    folders = []
    try:
        for account in outlook.Folders:
            for folder in account.Folders:
                folders.append(folder)
    except Exception as e:
        print(f"Error getting folders: {e}")
        
    print(f"Searching in {len(folders)} folders...")
    
    for folder in folders:
        print(f"Searching folder: {folder.Name}")
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True) # Newest first
            
            # Limit to 5000 emails per folder to avoid taking forever
            count = 0
            for msg in items:
                count += 1
                if count > 5000:
                    break
                    
                if msg.Class != 43: # olMail
                    continue
                    
                subject = getattr(msg, 'Subject', '') or ''
                sender = getattr(msg, 'SenderName', '') or ''
                body = getattr(msg, 'Body', '') or ''
                
                text_to_search = (subject + " " + sender + " " + body).lower()
                
                # Check if any term is in the email
                matched_terms = [t for t in terms if t in text_to_search]
                
                # Prioritize emails that have at least one term
                if matched_terms:
                    results.append({
                        "Folder": folder.Name,
                        "Date": str(getattr(msg, 'ReceivedTime', '')),
                        "Sender": sender,
                        "Subject": subject,
                        "Matched": ", ".join(matched_terms)
                    })
        except Exception as e:
            # Some folders don't have items or are restricted
            pass

    with open("search_results.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=4)
        
    print(f"\nFound {len(results)} matching emails.")
    for r in results:
        print(f"[{r['Folder']}] {r['Date']} | {r['Sender']}: {r['Subject']}")

if __name__ == '__main__':
    search_emails()
