import re
with open(r'C:\Users\Zvi\.gemini\antigravity\brain\d87047a2-5e99-4d1c-9993-702c3c3db5ae\Operation_Instructions.md', 'r', encoding='utf-8') as f:
    text = f.read()

# Replace blockquotes with github alerts
# Example:
# > [!NOTE]
# > **הודעה 1 - אזהרה**
# Becomes:
# **הערה חשובה:**
# **הודעה 1 - אזהרה**

text = text.replace("> [!NOTE]\n>", "**הערה:**\n")
text = text.replace("> [!WARNING]\n>", "**אזהרה חשובה:**\n")
text = text.replace("> [!TIP]\n>", "**טיפ:**\n")
text = text.replace("> [!SUCCESS]\n>", "**הצלחה:**\n")

text = text.replace("> [!NOTE]", "**הערה:**")
text = text.replace("> [!WARNING]", "**אזהרה חשובה:**")
text = text.replace("> [!TIP]", "**טיפ:**")
text = text.replace("> [!SUCCESS]", "**הצלחה:**")

# Remove remaining blockquote markers
text = re.sub(r'^>\s?', '', text, flags=re.MULTILINE)

with open(r'C:\Users\Zvi\.gemini\antigravity\brain\d87047a2-5e99-4d1c-9993-702c3c3db5ae\Operation_Instructions.md', 'w', encoding='utf-8') as f:
    f.write(text)

print("Markdown artifacts cleaned up.")
