import re, json
with open('c:/ledugma/search_results.html', encoding='utf-8') as f:
    html = f.read()
snippets = re.findall(r'<p class="snippet">(.*?)</p>', html, flags=re.DOTALL)
json.dump(snippets, open('c:/ledugma/res.json','w',encoding='utf-8'), ensure_ascii=False)
