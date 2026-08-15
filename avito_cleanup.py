#!/usr/bin/env python3
"""Чистка объявлений Авито по числу контактов.

Режимы:
  --dry-run (по умолчанию)  собрать объявления + контакты, записать CSV-отчёт,
                            ничего не менять
  --execute --report X.csv  выполнить действия (DELETE/ACTIVATE) из отчёта

Креды: переменные окружения AVITO_CLIENT_ID / AVITO_CLIENT_SECRET (или .env).
"""
import argparse
import csv
import json
import logging
import os
import sys
import time
from datetime import date, timedelta
from pathlib import Path

import requests

API = "https://api.avito.ru"
REPORT_DIR = Path(__file__).parent / "reports"
STATS_WINDOW_DAYS = 269  # заявленная глубина статистики — 270 дней
STATS_BATCH = 200
ITEMS_PAGE_PAUSE = 2.6   # лимит 25 req/min на списке объявлений
ACTION_PAUSE = 1.0

log = logging.getLogger("avito_cleanup")


def load_env():
    env_file = Path(__file__).parent / ".env"
    if env_file.exists():
        for line in env_file.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, _, v = line.partition("=")
                os.environ.setdefault(k.strip(), v.strip())


class AvitoClient:
    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self.http = requests.Session()
        self.token = None
        self.token_exp = 0.0

    def _auth(self):
        r = self.http.post(f"{API}/token", data={
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
        }, timeout=30)
        r.raise_for_status()
        d = r.json()
        self.token = d["access_token"]
        self.token_exp = time.time() + d.get("expires_in", 86400) - 120
        log.info("получен новый токен")

    def request(self, method, path, params=None, json_body=None):
        """Запрос с ретраями на 429/5xx и рефрешем токена на 401/403."""
        last = None
        for attempt in range(6):
            if not self.token or time.time() > self.token_exp:
                self._auth()
            r = self.http.request(
                method, API + path, params=params, json=json_body,
                headers={"Authorization": f"Bearer {self.token}"}, timeout=60)
            last = r
            if r.status_code == 429:
                wait = float(r.headers.get("Retry-After") or 2 ** attempt * 3)
                log.warning("429 на %s, ждём %ss", path, wait)
                time.sleep(min(wait, 120))
                continue
            if r.status_code in (401, 403) and attempt == 0:
                # Авито отдаёт 403 на просроченный токен
                self.token = None
                continue
            if r.status_code >= 500:
                time.sleep(2 ** attempt)
                continue
            return r
        return last


def fetch_all_items(cl, statuses):
    items, page = {}, 1
    while True:
        r = cl.request("GET", "/core/v1/items",
                       params={"per_page": 99, "page": page, "status": statuses})
        r.raise_for_status()
        res = r.json().get("resources", [])
        if not res:
            break
        for it in res:
            items[it["id"]] = it
        log.info("страница %d: +%d объявлений (всего %d)", page, len(res), len(items))
        page += 1
        time.sleep(ITEMS_PAGE_PAUSE)
    return list(items.values())


def stats_window(cl, user_id, ids, d_from, d_to):
    """Сумма uniqContacts по каждому id за окно. None => окно отвергнуто API."""
    r = cl.request("POST", f"/stats/v1/accounts/{user_id}/items", json_body={
        "itemIds": ids,
        "dateFrom": d_from.isoformat(),
        "dateTo": d_to.isoformat(),
        "fields": ["uniqContacts"],
        "periodGrouping": "month",
    })
    if r.status_code == 400:
        log.warning("окно %s..%s отвергнуто API: %s", d_from, d_to, r.text[:200])
        return None
    r.raise_for_status()
    out = {}
    for it in r.json().get("result", {}).get("items", []):
        out[it["itemId"]] = sum(s.get("uniqContacts") or 0 for s in it.get("stats", []))
    return out


def fetch_contacts(cl, user_id, item_ids):
    """Контакты за максимально доступную глубину, окнами по 270 дней назад.

    Возвращает (totals: dict id->int, coverage_from: date).
    """
    today = date.today()
    totals = {i: 0 for i in item_ids}
    coverage_from = today - timedelta(days=STATS_WINDOW_DAYS)

    batches = [item_ids[i:i + STATS_BATCH] for i in range(0, len(item_ids), STATS_BATCH)]

    # окно 0 — последние 270 дней
    for n, b in enumerate(batches, 1):
        got = stats_window(cl, user_id, b, coverage_from, today)
        if got is None:
            raise RuntimeError("API отверг даже базовое окно статистики")
        for i, v in got.items():
            totals[i] += v
        log.info("статистика (окно 0): батч %d/%d готов", n, len(batches))
        time.sleep(1)

    # идём в прошлое, пока API принимает окна и пока есть ненулевые данные
    d_to = coverage_from - timedelta(days=1)
    empty_windows = 0
    while empty_windows < 2 and d_to.year >= 2010:
        d_from = d_to - timedelta(days=STATS_WINDOW_DAYS)
        window_sum = 0
        rejected = False
        for b in batches:
            got = stats_window(cl, user_id, b, d_from, d_to)
            if got is None:
                rejected = True
                break
            for i, v in got.items():
                totals[i] += v
                window_sum += v
            time.sleep(1)
        if rejected:
            log.info("глубина статистики ограничена API: покрыто с %s", coverage_from)
            break
        coverage_from = d_from
        log.info("окно %s..%s: +%d контактов", d_from, d_to, window_sum)
        empty_windows = empty_windows + 1 if window_sum == 0 else 0
        d_to = d_from - timedelta(days=1)

    return totals, coverage_from


def fetch_start_time(cl, user_id, item_id):
    r = cl.request("GET", f"/core/v1/accounts/{user_id}/items/{item_id}/")
    if r.ok:
        return (r.json().get("start_time") or "")[:10]
    return ""


def check_autoload(cl):
    r = cl.request("GET", "/autoload/v2/profile")
    if r.ok:
        return r.json()
    return {"error": r.status_code, "body": r.text[:300]}


def classify(item, contacts):
    if contacts <= 1:
        return "DELETE"
    if item["status"] == "active":
        return "KEEP"
    return "ACTIVATE"


def write_report(path, rows):
    path.parent.mkdir(exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=[
            "item_id", "title", "status", "price", "url", "total_contacts",
            "stats_from", "stats_to", "start_time", "older_than_stats",
            "planned_action", "action_result", "action_error"])
        w.writeheader()
        w.writerows(rows)


def dry_run(cl, args):
    r = cl.request("GET", "/core/v1/accounts/self")
    r.raise_for_status()
    user_id = r.json()["id"]
    log.info("аккаунт: %s (id=%s)", r.json().get("name"), user_id)

    items = fetch_all_items(cl, args.statuses)
    log.info("всего объявлений (%s): %d", args.statuses, len(items))
    if not items:
        log.info("объявлений нет — отчёт не нужен")
        return

    totals, coverage_from = fetch_contacts(cl, user_id, [it["id"] for it in items])

    rows = []
    delete_candidates = [it for it in items if totals.get(it["id"], 0) <= 1]
    log.info("кандидатов на удаление: %d — запрашиваю даты размещения", len(delete_candidates))
    start_times = {}
    for n, it in enumerate(delete_candidates, 1):
        start_times[it["id"]] = fetch_start_time(cl, user_id, it["id"])
        if n % 100 == 0:
            log.info("даты размещения: %d/%d", n, len(delete_candidates))
        time.sleep(0.15)

    today = date.today()
    for it in items:
        contacts = totals.get(it["id"], 0)
        st = start_times.get(it["id"], "")
        older = ""
        if it["id"] in start_times:
            older = "yes" if (st and st < coverage_from.isoformat()) else "no"
        rows.append({
            "item_id": it["id"],
            "title": it.get("title", ""),
            "status": it.get("status", ""),
            "price": it.get("price", ""),
            "url": it.get("url", ""),
            "total_contacts": contacts,
            "stats_from": coverage_from.isoformat(),
            "stats_to": today.isoformat(),
            "start_time": st,
            "older_than_stats": older,
            "planned_action": classify(it, contacts),
            "action_result": "",
            "action_error": "",
        })

    rows.sort(key=lambda r: (r["planned_action"], -int(r["total_contacts"])))
    report = REPORT_DIR / f"report_{today.isoformat()}_{int(time.time())}.csv"
    write_report(report, rows)

    n = lambda a: sum(1 for r in rows if r["planned_action"] == a)
    log.info("=== СВОДКА ===")
    log.info("DELETE:   %d", n("DELETE"))
    log.info("ACTIVATE: %d", n("ACTIVATE"))
    log.info("KEEP:     %d", n("KEEP"))
    log.info("покрытие статистики: с %s по %s", coverage_from, today)
    log.info("старше покрытия статистики (среди DELETE): %d",
             sum(1 for r in rows if r["older_than_stats"] == "yes"))
    log.info("отчёт: %s", report)

    log.info("автовыгрузка: %s", json.dumps(check_autoload(cl), ensure_ascii=False)[:400])


# --- execute -----------------------------------------------------------------

DELETE_CANDIDATE_ENDPOINTS = [
    ("DELETE", "/core/v1/accounts/{uid}/items/{iid}/"),
    ("DELETE", "/core/v1/accounts/{uid}/items/{iid}"),
    ("POST", "/core/v1/accounts/{uid}/items/{iid}/delete"),
    ("POST", "/core/v1/accounts/{uid}/items/{iid}/deleted"),
    ("DELETE", "/core/v1/items/{iid}"),
    ("POST", "/core/v1/items/{iid}/delete"),
]
ACTIVATE_CANDIDATE_ENDPOINTS = [
    ("POST", "/core/v1/accounts/{uid}/items/{iid}/activate"),
    ("POST", "/core/v1/accounts/{uid}/items/{iid}/restore"),
    ("POST", "/core/v1/items/{iid}/activate"),
]


def item_status(cl, uid, iid):
    r = cl.request("GET", f"/core/v1/accounts/{uid}/items/{iid}/")
    return r.json().get("status") if r.ok else None


def probe_endpoint(cl, uid, iid, candidates, want_statuses):
    """Пробует кандидатов на одном объявлении; возвращает (method, path) или None."""
    for method, tpl in candidates:
        path = tpl.format(uid=uid, iid=iid)
        r = cl.request(method, path)
        log.info("проба %s %s -> %s %s", method, path, r.status_code, r.text[:150])
        if r.status_code in (200, 204):
            time.sleep(2)
            st = item_status(cl, uid, iid)
            log.info("статус объявления %s после пробы: %s", iid, st)
            if st in want_statuses or st is None:
                return (method, tpl)
        time.sleep(ACTION_PAUSE)
    return None


def load_state(path):
    if path.exists():
        return json.loads(path.read_text())
    return {}


def execute(cl, args):
    r = cl.request("GET", "/core/v1/accounts/self")
    r.raise_for_status()
    uid = r.json()["id"]

    with open(args.report, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    to_delete = [r for r in rows if r["planned_action"] == "DELETE"]
    to_activate = [r for r in rows if r["planned_action"] == "ACTIVATE"]
    log.info("из отчёта: DELETE=%d, ACTIVATE=%d", len(to_delete), len(to_activate))

    state_path = REPORT_DIR / "state.json"
    state = load_state(state_path) if args.resume else {}

    if not args.yes:
        ans = input(f"Удалить {len(to_delete)} объявлений НАВСЕГДА? Введите DELETE: ")
        if ans.strip() != "DELETE":
            log.info("отменено")
            return

    # проба на одном объявлении из каждого списка
    del_ep = act_ep = None
    if to_delete:
        del_ep = probe_endpoint(cl, uid, to_delete[0]["item_id"],
                                DELETE_CANDIDATE_ENDPOINTS,
                                {"removed", "not_found", "archived"})
        log.info("рабочий эндпоинт удаления: %s", del_ep)
    if to_activate:
        act_ep = probe_endpoint(cl, uid, to_activate[0]["item_id"],
                                ACTIVATE_CANDIDATE_ENDPOINTS, {"active"})
        log.info("рабочий эндпоинт активации: %s", act_ep)

    if not del_ep and not act_ep:
        log.error("ни один эндпоинт не сработал — публичный API не поддерживает "
                  "удаление/активацию. Дальше: автовыгрузка (план В).")
        return

    def run(batch, ep, action):
        ok = fail = 0
        for row in batch:
            iid = row["item_id"]
            if state.get(iid, {}).get("result") == "ok":
                continue
            method, tpl = ep
            r = cl.request(method, tpl.format(uid=uid, iid=iid))
            res = "ok" if r.status_code in (200, 204) else f"http_{r.status_code}"
            state[iid] = {"action": action, "result": res,
                          "error": "" if res == "ok" else r.text[:200]}
            state_path.write_text(json.dumps(state, ensure_ascii=False, indent=1))
            ok += res == "ok"
            fail += res != "ok"
            if (ok + fail) % 25 == 0:
                log.info("%s: %d ok / %d fail", action, ok, fail)
            time.sleep(ACTION_PAUSE)
        log.info("%s завершено: %d ok / %d fail", action, ok, fail)

    if del_ep:
        run(to_delete, del_ep, "DELETE")
    else:
        log.error("удаление недоступно через API — только автовыгрузка/кабинет")
    if act_ep:
        run(to_activate, act_ep, "ACTIVATE")
    else:
        log.error("активация недоступна через API — только автовыгрузка/кабинет")

    for row in rows:
        st = state.get(row["item_id"])
        if st:
            row["action_result"] = st["result"]
            row["action_error"] = st.get("error", "")
    out = Path(args.report).with_name(Path(args.report).stem + "_executed.csv")
    write_report(out, rows)
    log.info("итоговый отчёт: %s", out)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--execute", action="store_true")
    p.add_argument("--report", help="CSV-отчёт для --execute")
    p.add_argument("--statuses", default="active,old")
    p.add_argument("--resume", action="store_true")
    p.add_argument("--yes", action="store_true", help="не спрашивать подтверждение")
    args = p.parse_args()

    REPORT_DIR.mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.StreamHandler(),
                  logging.FileHandler(REPORT_DIR / "run.log", encoding="utf-8")])

    load_env()
    cid, csec = os.environ.get("AVITO_CLIENT_ID"), os.environ.get("AVITO_CLIENT_SECRET")
    if not cid or not csec:
        sys.exit("нет AVITO_CLIENT_ID / AVITO_CLIENT_SECRET")
    cl = AvitoClient(cid, csec)

    if args.execute:
        if not args.report:
            sys.exit("--execute требует --report path/to/report.csv")
        execute(cl, args)
    else:
        dry_run(cl, args)


if __name__ == "__main__":
    main()
