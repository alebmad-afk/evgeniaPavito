#!/usr/bin/env python3
"""Выгружает раздел «Заблокировано» текущего отчёта в CSV — для обращения в поддержку."""
import csv, json, os, re, subprocess, time

TOKF = os.environ.get('AVITO_TOKEN_FILE', '.avito_access_token')
OUT = os.environ.get('OUT', 'reports/blocked.csv')


def api(path):
    tok = open(TOKF).read().strip()
    for _ in range(8):
        r = subprocess.run(['curl', '-sS', '--retry', '3', '--retry-all-errors',
                            '-H', 'Authorization: Bearer ' + tok,
                            'https://api.avito.ru' + path], capture_output=True)
        try:
            return json.loads(r.stdout)
        except Exception:
            time.sleep(4)
    return None


rows, page = [], 1
while True:
    d = api('/autoload/v4/uploads/current/items?sections=error_blocked&perPage=20&page=%d' % page)
    if not d:
        break
    items = d.get('items', [])
    for i in items:
        msgs = [m for m in i.get('messages', []) if m.get('type') == 'error']
        m = msgs[0] if msgs else {}
        rows.append({
            'avito_id': i.get('avito_id'), 'url': i.get('url'),
            'код': m.get('code', ''),
            'причина': re.sub('<[^>]+>', ' ', m.get('title', '')).strip(),
            'подробности': re.sub('<[^>]+>', ' ', m.get('description', '')).replace('\n', ' ').strip()[:300],
        })
    if len(items) < 20:
        break
    page += 1
    time.sleep(1.8)

os.makedirs(os.path.dirname(OUT) or '.', exist_ok=True)
with open(OUT, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0]) if rows else ['avito_id'], delimiter=';')
    w.writeheader(); w.writerows(rows)
print('заблокированных строк:', len(rows), '->', OUT)
import collections
print('по кодам:', collections.Counter(r['код'] for r in rows))
