import win32com.client
import sys
import json
import pythoncom
import re
import webbrowser
import urllib.parse

def process_emails(actions_list):
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        processed_count = 0
        failed_count = 0
        for item in actions_list:
            entry_id = item.get("EntryID")
            action = item.get("Action")
            
            if not entry_id or action == "keep":
                continue
                
            try:
                # GetItemFromID retrieves the specific item
                message = outlook.GetItemFromID(entry_id)
                if action == "delete":
                    message.Delete() # Moves to Deleted Items
                elif action == "handle":
                    message.Categories = "לטיפול"
                    message.Save()
                elif action == "wait":
                    message.Categories = "להמתין"
                    message.Save()
                elif action == "archive":
                    message.Categories = "ארכיון"
                    message.Save()
                elif action == "unsubscribe":
                    try:
                        pa = message.PropertyAccessor
                        headers = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
                        if headers:
                            match = re.search(r"List-Unsubscribe:\s*(.+)", headers, re.IGNORECASE)
                            if match:
                                unsub_data = match.group(1)
                                mailto_match = re.search(r"<mailto:([^>]+)>", unsub_data)
                                http_match = re.search(r"<(https?[^>]+)>", unsub_data)
                                
                                if mailto_match:
                                    mailto_url = mailto_match.group(1)
                                    email_addr = mailto_url.split('?')[0]
                                    subject_str = "Unsubscribe"
                                    if "?subject=" in mailto_url:
                                        subject_str = mailto_url.split('?subject=')[1].split('&')[0]
                                        subject_str = urllib.parse.unquote(subject_str)
                                    
                                    # Send email via Outlook
                                    app = win32com.client.Dispatch("Outlook.Application")
                                    mail = app.CreateItem(0)
                                    mail.To = email_addr
                                    mail.Subject = subject_str
                                    mail.Body = "Please unsubscribe me from this list."
                                    mail.Send()
                                    print(f"Sent unsubscribe email to {email_addr}")
                                elif http_match:
                                    http_url = http_match.group(1)
                                    webbrowser.open(http_url)
                                    print(f"Opened unsubscribe link: {http_url}")
                    except Exception as inner_e:
                        print(f"Failed to extract or perform unsubscribe for {entry_id}: {inner_e}")
                    
                    # Delete the email after unsubscribe attempt
                    message.Delete()
                    
                processed_count += 1
            except Exception as e:
                print(f"Failed to process item {entry_id}: {e}")
                failed_count += 1
                
        return processed_count, failed_count
    except Exception as e:
        print(f"Failed to connect to Outlook: {e}")
        return 0, 0

def open_email(entry_id):
    try:
        import subprocess
        import os
        ps_script = os.path.abspath('open_email.ps1')
        subprocess.Popen(['powershell', '-WindowStyle', 'Hidden', '-ExecutionPolicy', 'Bypass', '-File', ps_script, '-EntryID', entry_id], 
                         creationflags=subprocess.CREATE_NO_WINDOW)
        return True
    except Exception as e:
        print(f"Failed to open email: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
        with open(filepath, 'r', encoding='utf-8') as f:
            actions_list = json.load(f)
        process_emails(actions_list)
    else:
        print("Please provide a path to a JSON file containing a list of actions.")
