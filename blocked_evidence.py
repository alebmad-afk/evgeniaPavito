#!/usr/bin/env python3
"""Доказательная база по блокировкам 2015: с чем сочли дублем и чем отличается.

    AVITO_TOKEN_FILE=... python3 blocked_evidence.py
"""
import csv, difflib, json, os, re, subprocess, time

TOKF = os.environ.get('AVITO_TOKEN_FILE', '.avito_access_token')


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


def all_items():
    """id → (заголовок, адрес, ссылка) по всем объявлениям аккаунта."""
    out = {}
    for st in ('active', 'blocked', 'old'):
        page = 1
        while True:
            d = api('/core/v1/items?per_page=99&page=%d&status=%s' % (page, st))
            if not d:
                break
            res = d.get('resources', [])
            for r in res:
                out[str(r['id'])] = (r.get('title', ''), r.get('address', ''),
                                     r.get('url', ''), st)
            if len(res) < 99:
                break
            page += 1
            time.sleep(2.5)
        time.sleep(2.5)
    return out


def blocked_pairs():
    """(id заблокированного, id того, с чем сочли дублем)."""
    out, page = [], 1
    while True:
        d = api('/autoload/v4/uploads/current/items'
                '?sections=error_blocked&perPage=20&page=%d' % page)
        if not d:
            break
        items = d.get('items', [])
        for i in items:
            for m in i.get('messages', []):
                if m.get('code') == 2015:
                    g = re.search(r'avito\.ru/items/(\d+)', m.get('description', ''))
                    out.append((str(i['avito_id']), g.group(1) if g else None))
        if len(items) < 20:
            break
        page += 1
        time.sleep(1.8)
    return out


def main():
    texts = json.load(open('data/ads_texts.json', encoding='utf-8'))
    body = {t.strip().split('\n')[0].strip(): '\n'.join(t.strip().split('\n')[1:])
            for t in texts.values()}

    items = all_items()
    pairs = blocked_pairs()
    print('заблокированных: %d, объявлений в аккаунте известно: %d'
          % (len(pairs), len(items)))

    rows, stats = [], {'разные города': 0, 'один город': 0,
                       'разные заголовки': 0, 'одинаковые заголовки': 0}
    for bid, oid in pairs:
        bt, ba, bu, _ = items.get(bid, ('?', '?', '', ''))
        ot, oa, ou, _ = items.get(oid, ('?', '?', '', '')) if oid else ('нет данных',) * 4
        same_city = ba == oa
        sim = difflib.SequenceMatcher(None, body.get(bt, ''), body.get(ot, '')).ratio()
        stats['один город' if same_city else 'разные города'] += 1
        stats['одинаковые заголовки' if bt == ot else 'разные заголовки'] += 1
        rows.append({
            'заблокировано': bt, 'город': ba, 'ссылка': bu,
            'сочтено дублем': ot, 'город 2': oa, 'ссылка 2': ou,
            'город совпадает': 'да' if same_city else 'нет',
            'заголовок совпадает': 'да' if bt == ot else 'нет',
            'похожесть описаний': '%.0f%%' % (sim * 100),
        })

    out = 'reports/blocked_evidence.csv'
    with open(out, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]), delimiter=';')
        w.writeheader(); w.writerows(rows)
    print(out)
    print(stats)
    sims = [float(r['похожесть описаний'].rstrip('%')) for r in rows]
    if sims:
        print('похожесть описаний: мин %.0f%%, среднее %.0f%%, макс %.0f%%'
              % (min(sims), sum(sims) / len(sims), max(sims)))


if __name__ == '__main__':
    main()
