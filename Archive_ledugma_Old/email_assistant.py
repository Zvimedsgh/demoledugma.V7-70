import win32com.client
import json
import re

def get_recent_emails(num_emails=10):
    try:
        import pythoncom
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        # 6 is the olFolderInbox enum
        inbox = outlook.GetDefaultFolder(6)
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True) # Sort by received time, descending
        
        email_data = []
        
        count = 0
        for message in messages:
            # Check if the item is a MailItem (Class == 43)
            if message.Class == 43:
                try:
                    sender = message.SenderName
                except Exception:
                    sender = "Unknown"
                    
                has_unsubscribe = False
                try:
                    pa = message.PropertyAccessor
                    headers = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
                    if headers and re.search(r"List-Unsubscribe:\s*", headers, re.IGNORECASE):
                        has_unsubscribe = True
                except Exception:
                    pass

                try:
                    categories = getattr(message, 'Categories', '') or ''
                    if 'לטיפול' in categories or 'להמתין' in categories or 'ארכיון' in categories:
                        continue
                except:
                    pass

                try:
                    email_info = {
                        "EntryID": message.EntryID,
                        "Sender": sender,
                        "Subject": message.Subject,
                        "ReceivedTime": str(message.ReceivedTime),
                        "BodySnippet": message.Body[:200].replace('\n', ' ').replace('\r', ' ') + "..." if message.Body else "",
                        "HasUnsubscribe": has_unsubscribe
                    }
                    email_data.append(email_info)
                    count += 1
                except Exception as e:
                    # Some items might have restrictions, just skip them
                    pass
                    
                if count >= num_emails:
                    break
                    
        return email_data
    except Exception as e:
        print(f"Failed to connect to Outlook or retrieve messages: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    print("Fetching recent emails from Outlook...")
    emails = get_recent_emails(20)
    with open('emails.json', 'w', encoding='utf-8') as f:
        json.dump(emails, f, indent=4, ensure_ascii=False)
    print("Saved to emails.json")
