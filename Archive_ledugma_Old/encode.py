def hebrew_to_chraw(text):
    out = []
    for c in text:
        if '\u05d0' <= c <= '\u05ea':
            out.append(f'ChrW({ord(c)})')
        elif c == ' ':
            out.append('ChrW(32)')
        elif c == '\n':
            out.append('vbCrLf')
        elif c == '-':
            out.append('ChrW(45)')
        elif c == ':':
            out.append('ChrW(58)')
        elif c == '!':
            out.append('ChrW(33)')
        elif c.isdigit():
            out.append(f'ChrW({ord(c)})')
        else:
            out.append(f'"{c}"')
    return ' & '.join(out)

text = "יש להתחיל כאן!\nהקש להורדת הוראות התקנה\nלתמיכה בוואטסאפ: 054-6677396"
print(hebrew_to_chraw(text))
