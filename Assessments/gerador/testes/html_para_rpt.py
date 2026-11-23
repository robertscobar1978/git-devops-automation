# Reconstrói um .rpt (formato "Resultados em arquivo" do SSMS) a partir do HTML de exemplo,
# para testar o gerador sem um .rpt real.
import re, sys, html

src, dst = sys.argv[1], sys.argv[2]
h = open(src, encoding='utf-8').read()


def raw(cell):
    cls = re.search(r'class="(\w+)"', cell.split('>')[0])
    v = html.unescape(re.sub(r'<[^>]+>', '', re.sub(r'^<td[^>]*>', '', cell)))
    if v == '—':
        return 'NULL', False
    if v in ('Sim', 'Não') and not cls:
        return ('1' if v == 'Sim' else '0'), True
    m = re.match(r'^(\d\d)/(\d\d)/(\d{4}) (\d\d:\d\d:\d\d)$', v)
    if m:
        return f'{m[3]}-{m[2]}-{m[1]} {m[4]}.000', False
    if cls and cls.group(1) == 'n' and re.match(r'^-?[\d.]+,\d+$', v):
        return v.replace('.', '').replace(',', '.'), True
    return v, bool(cls and cls.group(1) == 'n')


def rs_txt(cols, rows, num):
    w = [max([len(c)] + [len(r[k]) for r in rows]) for k, c in enumerate(cols)]
    out = [' '.join(c.ljust(w[k]) for k, c in enumerate(cols)), ' '.join('-' * x for x in w)]
    for r in rows:
        out.append(' '.join((r[k].rjust(w[k]) if num[k] else r[k].ljust(w[k])) for k in range(len(cols))))
    return '\n'.join(out) + '\n\n'


out = []
# Dashboard
dash = [('01', 'IDENTIFICAÇÃO', 'Servidor', 'PENINSULA\\SQL2017'), ('03', 'IDENTIFICAÇÃO', 'SQL Server', 'SQL Server 2017'),
        ('10', 'TOPOLOGIA', 'FCI / Cluster', 'NÃO'), ('11', 'TOPOLOGIA', 'Always On / HADR', 'NÃO HABILITADO'),
        ('12', 'TOPOLOGIA', 'Availability Groups', '0'), ('13', 'TOPOLOGIA', 'Réplicas', '0'),
        ('14', 'TOPOLOGIA', 'Linked Servers', '0'), ('17', 'RECURSOS', 'Databases', '23'), ('19', 'RECURSOS', 'SQL Agent Jobs', '45639')]
out.append(rs_txt(['ID', 'Categoria', 'Indicador', 'Valor', 'Status', 'Observacao'],
                  [[a, b, c, d, 'INFO', 'x'] for a, b, c, d in dash], [0] * 6))
# Checklist
chk = re.search(r'<section id="checklist">.*?<tbody>(.*?)</tbody>', h, re.S).group(1)
rows, cat = [], ''
for tr in re.findall(r'<tr.*?</tr>', chk, re.S):
    if 'class="grp"' in tr:
        cat = html.unescape(re.sub('<[^>]+>', '', tr))
        continue
    tds = re.findall(r'<td.*?</td>', tr, re.S)
    vals = [html.unescape(re.sub('<[^>]+>', '', t)) for t in tds]
    vals = ['NULL' if v == '—' else v for v in vals]
    rows.append([str(len(rows) + 1), cat] + vals)
out.append(rs_txt(['ID', 'Categoria', 'Indicador', 'Valor', 'Status', 'Observacao'], rows, [1, 0, 0, 0, 0, 0]))
# Detalhes, na ordem do id
secs = {int(m.group(1)): m.group(0) for m in re.finditer(r'<details id="s(\d+)".*?</details>', h, re.S)}
for k in sorted(secs):
    d = secs[k]
    if 'class="note">' in d and '<table' not in d:
        nota = html.unescape(re.search(r'<div class="note">(.*?)</div>', d).group(1))
        out.append(rs_txt(['Information'], [[nota]], [0]))
        continue
    if 'kv' in d:
        kv = re.findall(r'<tr><th>([^<]*)</th>(<td>.*?</td>)</tr>', d)
        cols = [a for a, b in kv]
        vals = [raw(b) for a, b in kv]
        out.append(rs_txt(cols, [[v for v, n in vals]], [n for v, n in vals]))
        continue
    cols = re.findall(r'<th>(.*?)</th>', d)
    rows, num = [], [0] * len(cols)
    for tr in re.findall(r'<tr>(<td.*?)</tr>', d, re.S):
        if 'class="empty"' in tr:
            continue
        cells = [raw(c) for c in re.findall(r'<td[^>]*>.*?</td>', tr, re.S)]
        rows.append([v for v, n in cells])
        num = [a or b for a, b in zip(num, [n for v, n in cells])]
    out.append(rs_txt([html.unescape(c) for c in cols], rows, num))
open(dst, 'w', encoding='utf-16').write(''.join(out))
print('ok', len(out))
