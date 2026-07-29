text = "להדרכה וסיוע שלח ווטסאפ לטלפון 054-6677396"
out = []
for c in text:
    if 'א' <= c <= 'ת':
        out.append(f"ChrW({ord(c)})")
    else:
        out.append(f'"{c}"')
print(" & ".join(out).replace('" & "', ''))
