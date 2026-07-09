import json

with open(r'C:\Users\prasa\Desktop\TS-Staff-System\staff_18_clean.json', encoding='utf-8') as f:
    data = json.load(f)

def fix(s):
    return s.replace('ï¿½', '-').replace('–', '-').replace('—', '-')

template_js = fix(json.dumps(data['template'], ensure_ascii=False))
duties_js = fix(json.dumps(data['duties'], ensure_ascii=False))

with open(r'C:\Users\prasa\Desktop\TS-Staff-System\final_vars.txt', 'w', encoding='utf-8') as f:
    f.write('DASHBOARD_STAFF=' + template_js + '\n')
    f.write('DUTIES_MAP=' + duties_js + '\n')

print('Done. Template chars:', len(template_js), 'Duties chars:', len(duties_js))
