#!/usr/bin/env python3
"""Сборка feed.xlsx: 30 объявлений × 10 городов МСК на ID уже снятых объявлений.

Источники:
  * официальный Excel-экспорт автозагрузки Авито (все объявления аккаунта
    со служебными колонками AvitoId / AvitoStatus);
  * `data/ads_texts.json` — тексты 30 объявлений с Яндекс.Диска клиента
    (первая строка файла = заголовок, остальное = описание);
  * фото на привязанном Яндекс.Диске: `avito_feed/porshneva/cover_NN.jpg`
    (уникальная обложка) + `common_02..10.jpg` (общие для всех объявлений).

Логика:
  * строки со статусом «Активно» переносятся в фид БЕЗ изменений — иначе
    отсутствие строки снимет объявление с публикации;
  * под 300 комбинаций (объявление × город) берутся ID строк со статусом
    «Снято с публикации»: размещение по ним уже оплачено, поэтому возврат
    в актив не тратит месячный лимит (в отличие от «Истёк срок публикации»);
  * приоритет при раздаче ID — строки, у которых Address уже совпадает
    с целевым городом, чтобы менять гео как можно реже;
  * служебные колонки экспорта (AvitoStatus, AvitoDateEnd, EMail, CompanyName,
    ImageNames) из фида удаляются — они ломают обработку;
  * пустые обязательные поля услуг заполняются (иначе ошибка 1073).

Использование:
    python3 build_city_feed.py <экспорт_авито.xlsx> [feed.xlsx]
"""
import html
import json
import os
import sys

import openpyxl

from cities import MSK_TOP10

SERVICE_SHEET = "Деловые услуги-Бухгалтерия, фин"
SKIP_PREFIX = ("Инструкция", "Спр-")
# колонки экспорта, которых не должно быть в фиде
DROP_COLUMNS = ("AvitoStatus", "AvitoDateEnd", "EMail", "CompanyName", "ImageNames")
# обязательные поля услуг, пустые в экспорте Авито (ошибка 1073)
REQUIRED_DEFAULTS = {"Consultations": "Есть", "WorkWithContract": "Да",
                     "Prepayment": "Нет", "Place": "Удалённо"}

DISK_DIR = "yandex_disk://avito_feed/porshneva"
COMMON_PHOTOS = ["%s/common_%02d.jpg" % (DISK_DIR, i) for i in range(2, 11)]

STATUS_ACTIVE = "Активно"
STATUS_ARCHIVED = "Снято с публикации"


def read_sheet(ws):
    """Строки листа как список (номер_строки, dict) — служебные строки 3-4 пропускаются."""
    header = [c.value for c in ws[2]]
    out = []
    for ri in range(5, ws.max_row + 1):
        if ws.cell(ri, 1).value in (None, ""):
            continue
        out.append((ri, dict(zip(header, [c.value for c in ws[ri]]))))
    return header, out


def make_description(title, body):
    """Текст с Диска → HTML, который принимает Авито (<strong>, <br/>)."""
    parts = ["<strong>%s</strong>" % html.escape(title)]
    for block in [b.strip() for b in body.split("\n\n") if b.strip()]:
        lines = []
        for line in block.split("\n"):
            line = html.escape(line.strip())
            # строки-подзаголовки вида «Что берём на себя:» — жирным
            if line.endswith(":") and len(line) < 60:
                line = "<strong>%s</strong>" % line
            lines.append(line)
        parts.append("<br/>".join(lines))
    return "<br/><br/>".join(parts)


def load_ads(path):
    """30 объявлений: ключ папки → (порядковый номер, заголовок, описание, фото)."""
    raw = json.load(open(path, encoding="utf-8"))
    ads = []
    for num, key in enumerate(sorted(raw), 1):
        lines = raw[key].strip().split("\n")
        title = lines[0].strip()
        body = "\n".join(lines[1:]).strip()
        if not title or not body:
            sys.exit("ОШИБКА: пустой текст объявления %s" % key)
        if len(title) > 50:
            sys.exit("ОШИБКА: заголовок длиннее 50 символов: %s" % key)
        photos = ["%s/cover_%02d.jpg" % (DISK_DIR, num)] + COMMON_PHOTOS
        ads.append({"key": key, "num": num, "title": title,
                    "description": make_description(title, body),
                    "images": " | ".join(photos)})
    return ads


def assign_ids(archived, ads, cities):
    """(объявление × город) → строка экспорта. Сначала строки с уже нужным гео."""
    pool = sorted(archived, key=lambda t: str(t[1]["Id"]))
    by_address = {}
    for item in pool:
        by_address.setdefault(str(item[1]["Address"]), []).append(item)
    used = set()
    plan = []
    # проход 1: для каждой пары берём строку, уже стоящую в нужном городе
    for _, address, _ in cities:
        bucket = by_address.get(address, [])
        for ad in ads:
            match = next((i for i in bucket if id(i) not in used), None)
            if match is None:
                plan.append((ad, address, None))
            else:
                used.add(id(match))
                plan.append((ad, address, match))
    # проход 2: остальным раздаём свободные ID подряд
    free = (i for i in pool if id(i) not in used)
    filled = []
    for ad, address, row in plan:
        if row is None:
            row = next(free, None)
            if row is None:
                sys.exit("ОШИБКА: снятых объявлений меньше, чем нужно комбинаций")
        filled.append((ad, address, row))
    return filled


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else "feed.xlsx"
    here = os.path.dirname(os.path.abspath(__file__))
    ads = load_ads(os.path.join(here, "data", "ads_texts.json"))

    wb = openpyxl.load_workbook(src)
    ws = wb[SERVICE_SHEET]
    header, rows = read_sheet(ws)

    active = [(ri, r) for ri, r in rows if r["AvitoStatus"] == STATUS_ACTIVE]
    archived = [(ri, r) for ri, r in rows if r["AvitoStatus"] == STATUS_ARCHIVED]
    print("в экспорте: активных %d, снятых с публикации %d, всего услуг %d"
          % (len(active), len(archived), len(rows)))

    need = len(ads) * len(MSK_TOP10)
    plan = assign_ids(archived, ads, MSK_TOP10)
    assert len(plan) == need

    col = {name: i + 1 for i, name in enumerate(header) if name}
    kept_rows = {ri for ri, _ in active}
    same_geo = 0
    for ad, address, (ri, row) in plan:
        if str(row["Address"]) == address:
            same_geo += 1
        ws.cell(ri, col["Title"]).value = ad["title"]
        ws.cell(ri, col["Description"]).value = ad["description"]
        ws.cell(ri, col["Address"]).value = address
        ws.cell(ri, col["ImageUrls"]).value = ad["images"]
        kept_rows.add(ri)
    print("подготовлено %d комбинаций (30 объявлений × %d городов), "
          "гео не менялось у %d" % (need, len(MSK_TOP10), same_geo))

    # лишние строки листа услуг — удалить (они уже сняты, отсутствие в фиде ничего не меняет)
    for ri in sorted({ri for ri, _ in rows} - kept_rows, reverse=True):
        ws.delete_rows(ri)

    # прочие листы с объявлениями чистим полностью: вакансия и квартира сейчас сняты,
    # возврат каждой из них — отдельное платное решение хозяина
    for name in wb.sheetnames:
        if name.startswith(SKIP_PREFIX) or name == SERVICE_SHEET:
            continue
        other = wb[name]
        _, other_rows = read_sheet(other)
        for ri in sorted((ri for ri, _ in other_rows), reverse=True):
            other.delete_rows(ri)
        if other_rows:
            print("лист %r: убрано %d строк" % (name, len(other_rows)))

    # обязательные поля услуг + удаление служебных колонок
    filled = 0
    for name in wb.sheetnames:
        if name.startswith(SKIP_PREFIX):
            continue
        sheet = wb[name]
        head = [c.value for c in sheet[2]]
        for col_name, value in REQUIRED_DEFAULTS.items():
            if col_name not in head:
                continue
            ci = head.index(col_name) + 1
            for ri in range(5, sheet.max_row + 1):
                if sheet.cell(ri, 1).value in (None, ""):
                    continue
                if sheet.cell(ri, ci).value in (None, ""):
                    sheet.cell(ri, ci).value = value
                    filled += 1
        for col_name in DROP_COLUMNS:
            head = [c.value for c in sheet[2]]
            if col_name in head:
                sheet.delete_cols(head.index(col_name) + 1)
    print("заполнено пустых обязательных полей: %d" % filled)
    print("удалены служебные колонки: %s" % ", ".join(DROP_COLUMNS))

    wb.save(dst)
    verify(dst, ads, active, need)


def verify(dst, ads, active, need):
    wb = openpyxl.load_workbook(dst)
    ws = wb[SERVICE_SHEET]
    header, rows = read_sheet(ws)
    for col_name in DROP_COLUMNS:
        assert col_name not in header, col_name
    ids = [str(r["Id"]) for _, r in rows]
    assert len(ids) == len(set(ids)), "дубли Id в фиде"
    assert len(rows) == need + len(active), "неверное число строк"
    titles = {ad["title"] for ad in ads}
    per_city = {}
    for _, r in rows:
        if r["Title"] in titles:
            per_city.setdefault(r["Address"], set()).add(r["Title"])
        for field in ("Id", "AvitoId", "Title", "Description", "Address", "ImageUrls",
                      "Category", "ServiceType", "ServiceSubtype", "Consultations",
                      "WorkWithContract", "Prepayment", "Place"):
            assert r.get(field) not in (None, ""), "пустое поле %s у Id=%s" % (field, r["Id"])
    assert len(per_city) == len(MSK_TOP10), "городов в фиде: %d" % len(per_city)
    for address, found in sorted(per_city.items()):
        assert len(found) == len(ads), "%s: %d объявлений" % (address, len(found))
    print("проверка пройдена: %s — %d строк (%d новых + %d активных без изменений), "
          "%d городов по %d объявлений"
          % (dst, len(rows), need, len(active), len(per_city), len(ads)))


if __name__ == "__main__":
    main()
