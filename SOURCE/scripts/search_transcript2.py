import sys
import os
import json

log_file = r'C:\Users\Zvi\.gemini\antigravity\brain\acccdbfb-db6c-42d6-9b9b-4c9a939b557a\.system_generated\logs\transcript.jsonl'
with open(log_file, 'r', encoding='utf-8') as f:
    for line in f:
        try:
            data = json.loads(line)
            if data.get('type') == 'USER_INPUT' or data.get('type') == 'PLANNER_RESPONSE':
                content = data.get('content', '')
                if 'הוראות_תפעול' in content or 'תפעול' in content or 'כפתור 1' in content:
                    print(f"[{data.get('type')}] {content[:150]}...")
        except Exception:
            pass
