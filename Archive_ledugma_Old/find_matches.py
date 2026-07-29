import json

keywords = ['מעוין', 'מעויין', 'diamond', 'מדריך', 'הוראות', 'הפעלה', 'תפעול']

with open(r'C:\Users\Zvi\.gemini\antigravity\brain\9a1929d9-eeae-443b-95e3-8045d9c9ea22\.system_generated\logs\transcript_full.jsonl', 'r', encoding='utf-8') as f, open(r'c:\ledugma\matches.txt', 'w', encoding='utf-8') as out:
    for line in f:
        try:
            data = json.loads(line)
            if data.get('type') == 'USER_INPUT':
                content = data.get('content', '')
                if any(k in content for k in keywords):
                    out.write(content + '\n---\n')
        except:
            pass
