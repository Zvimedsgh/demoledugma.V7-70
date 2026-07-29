lines = open(r'C:\Users\Zvi\.gemini\antigravity\brain\a19cc612-509d-4585-8e80-d112928a73e4\.system_generated\steps\822\content.md', encoding='utf-8').read().splitlines()
for i, line in enumerate(lines):
    if "BIC" in line or "SWIFT" in line:
        for j in range(max(0, i-5), min(len(lines), i+6)):
            print(lines[j].strip())
