import sys
import json

log_file = r'C:\Users\Zvi\.gemini\antigravity\brain\acccdbfb-db6c-42d6-9b9b-4c9a939b557a\.system_generated\logs\transcript.jsonl'
with open(log_file, 'r', encoding='utf-8') as f:
    for line in f:
        try:
            data = json.loads(line)
            if data.get('type') == 'USER_INPUT' or data.get('type') == 'PLANNER_RESPONSE':
                content = data.get('content', '')
                if 'תפעול' in content or 'עושים' in content or 'כפתורים' in content:
                    print(f"[{data.get('type')}] {content[:200]}...")
        except Exception:
            pass
