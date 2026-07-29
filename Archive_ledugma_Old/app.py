import os
import json
import subprocess

def kill_existing_on_port(port):
    """Kill any existing process listening on the given port before we start."""
    try:
        result = subprocess.run(
            f'netstat -ano | findstr ":{port}" | findstr "LISTENING"',
            capture_output=True, text=True, shell=True
        )
        pids = set()
        my_pid = os.getpid()
        for line in result.stdout.strip().split('\n'):
            if line.strip():
                parts = line.strip().split()
                pid = int(parts[-1])
                if pid != my_pid and pid != 0:
                    pids.add(pid)
        for pid in pids:
            try:
                subprocess.run(f'taskkill /F /PID {pid}', shell=True, capture_output=True)
                print(f"Killed old process on port {port}: PID {pid}")
            except:
                pass
    except:
        pass

kill_existing_on_port(5004)

from flask import Flask, render_template, request, jsonify, send_file
import email_assistant
import email_manager

app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

RULES_FILE = 'rules.json'
AD_KEYWORDS = ["פרסומת", "ניוזלטר", "newsletter", "promotions", "מבצע", "buyme", "tripadvisor", "שופרסל", "alm", "משרדיה", "everywear", "wix.com", "אמיגו"]
HANDLE_KEYWORDS = ["חשבונית", "invoice", "receipt", "קבלה", "הזמנה", "order", "דחוף", "urgent"]

def load_rules():
    if os.path.exists(RULES_FILE):
        try:
            with open(RULES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_rules(rules):
    with open(RULES_FILE, 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=4)

def get_recommendation(email, rules):
    sender = email.get('Sender', '').strip()
    subject = email.get('Subject', '').lower()
    has_unsub = email.get('HasUnsubscribe', False)
    
    # 1. Check user defined rules by exact sender name
    if sender in rules:
        return rules[sender]
        
    sender_lower = sender.lower()
    # 2. Check general keywords
    for kw in AD_KEYWORDS:
        if kw in sender_lower or kw in subject:
            return "unsubscribe" if has_unsub else "delete"
            
    for kw in HANDLE_KEYWORDS:
        if kw in sender_lower or kw in subject:
            return "handle"
            
    return "keep"

@app.route('/')
def index():
    try:
        rules = load_rules()
        emails = email_assistant.get_recent_emails(100)
        for em in emails:
            em['recommendation'] = get_recommendation(em, rules)
        
        response = app.make_response(render_template('index.html', emails=emails))
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"Error loading emails: {e}"

@app.route('/search_results')
def search_results():
    try:
        return send_file('C:\\ledugma\\search_results.html')
    except Exception as e:
        return str(e), 500

@app.route('/process', methods=['POST'])
def process_emails():
    data = request.json
    actions_list = data.get('actions', [])
    if not actions_list:
        return jsonify({"success": False, "message": "No actions provided."}), 400
        
    try:
        # Save rules if requested
        rules = load_rules()
        rules_changed = False
        
        for item in actions_list:
            if item.get("SaveRule") and item.get("Sender") and item.get("Action"):
                rules[item["Sender"].strip()] = item["Action"]
                rules_changed = True
                
        if rules_changed:
            save_rules(rules)

        # Process the emails (only pass Action and EntryID to email_manager)
        processed, failed = email_manager.process_emails(actions_list)
        if failed > 0:
            return jsonify({"success": True, "message": f"טופלו {processed} מיילים בהצלחה. {failed} מיילים לא נמצאו (ייתכן שכבר טופלו או נמחקו)."})
        return jsonify({"success": True, "message": f"כל {processed} המיילים טופלו בהצלחה!"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

@app.route('/open/<path:entry_id>')
def open_email_route(entry_id):
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        message = outlook.GetItemFromID(entry_id)
        message.Display()
        return "<script>alert('האימייל נפתח באאוטלוק! נא לבדוק בשורת המשימות (Taskbar).'); window.close();</script>"
    except Exception as e:
        return f"Failed to open email: {e}", 500

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5004)
