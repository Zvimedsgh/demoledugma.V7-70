import sys

# Target string: "להתקנה והדרכה שלח WhatsApp ל – 054-6677396"
chars = []
for c in "להתקנה והדרכה שלח WhatsApp ל – 054-6677396":
    if ord(c) > 127:
        chars.append(f"ChrW({ord(c)})")
    else:
        chars.append(f'"{c}"')

# Optimize adjacent strings
optimized = []
current_str = ""
for c in chars:
    if c.startswith('"'):
        current_str += c[1:-1]
    else:
        if current_str:
            optimized.append(f'"{current_str}"')
            current_str = ""
        optimized.append(c)
if current_str:
    optimized.append(f'"{current_str}"')

final_vba_string = " & ".join(optimized)
print(final_vba_string)

