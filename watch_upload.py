#!/usr/bin/env python3
"""Мониторит текущую выгрузку до терминального статуса и печатает разбивку.

    AVITO_TOKEN_FILE=/путь/к/токену python3 watch_upload.py [ГГГГ-ММ-ДД]

Дата — граница «свежести»: терминальным считается только прогон, начатый
не раньше неё, иначе скрипт сразу увидит прошлую завершённую выгрузку.
"""
import json, os, subprocess, sys, time

TOKF = os.environ.get('AVITO_TOKEN_FILE', '.avito_access_token')
SINCE = sys.argv[1] if len(sys.argv) > 1 else time.strftime('%Y-%m-%d')
TERMINAL = {'success', 'success_warning', 'error'}


def api(path):
    tok = open(TOKF).read().strip()
    for _ in range(6):
        r = subprocess.run(['curl', '-sS', '--retry', '3', '--retry-all-errors', '-H',
                            'Authorization: Bearer ' + tok,
                            'https://api.avito.ru' + path], capture_output=True)
        try:
            return json.loads(r.stdout)
        except Exception:
            time.sleep(5)
    return None


def flat(sections, out=None):
    out = {} if out is None else out
    for s in sections:
        out[s['slug']] = s['count']
        flat(s.get('sections', []), out)
    return out


prev = None
for i in range(90):
    d = api('/autoload/v4/uploads?per_page=1&page=1')
    if not d or not d.get('uploads'):
        time.sleep(60); continue
    u = d['uploads'][0]
    cur = (u['upload_id'], u['status'], json.dumps(flat(u['stats']['sections']), sort_keys=True))
    if cur != prev:
        prev = cur
        print('%s | %s | %s | обработано %s | %s' % (
            time.strftime('%H:%M:%S'), u['upload_id'], u['status'],
            u['stats'].get('count'), flat(u['stats']['sections'])), flush=True)
    if u['status'] in TERMINAL and u['started_at'][:10] >= SINCE:
        print('ТЕРМИНАЛЬНЫЙ СТАТУС', flush=True)
        break
    time.sleep(60)
