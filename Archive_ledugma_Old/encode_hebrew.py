text = "ראה גיליון הוראות התקנה, לעזרה בהתקנה והדרכה שלח WhatsApp לטלפון 054-6677396."
output = ""
for char in text:
    if '\u0590' <= char <= '\u05EA':
        output += f'ChrW({ord(char)}) & '
    elif char == ' ':
        output += '" " & '
    elif char == ',':
        output += '"," & '
    elif char == '-':
        output += '"-" & '
    elif char == '.':
        output += '"." & '
    elif char == ':':
        output += '":" & '
    else:
        output += f'"{char}" & '

print(output[:-3])
