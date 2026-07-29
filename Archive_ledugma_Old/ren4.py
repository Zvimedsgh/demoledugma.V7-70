import os

target = "הגיליון והחוברת לא מוגנים ואין הסתר או בטל נסתרה אני שוב לא מוצא את הג".encode('utf-8')
replacement = "Demo Reports System - Final Version".encode('utf-8')

# Ensure same length
if len(replacement) < len(target):
    replacement = replacement + b' ' * (len(target) - len(replacement))
elif len(replacement) > len(target):
    replacement = replacement[:len(target)]

paths = [
    r"C:\Users\Zvi\.gemini\antigravity\agyhub_summaries_proto.pb",
    r"C:\Users\Zvi\.gemini\antigravity\conversations\9a1929d9-eeae-443b-95e3-8045d9c9ea22.db"
]

for path in paths:
    try:
        with open(path, 'rb') as f:
            data = f.read()
        
        if target in data:
            new_data = data.replace(target, replacement)
            with open(path, 'wb') as f:
                f.write(new_data)
            print(f"Successfully replaced in {path}")
        else:
            print(f"Target not found in {path}")
    except Exception as e:
        print(f"Error processing {path}: {e}")
