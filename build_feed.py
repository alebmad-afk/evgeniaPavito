#!/usr/bin/env python3
"""Сборка отфильтрованного фида автовыгрузки Авито.

Берёт исходный XML-фид и reports/ad_id_map.json (карта avito_id -> ad_id +
действия из отчёта), удаляет из фида все <Ad>, чей Id соответствует
объявлениям с планируемым действием DELETE. Остальные (KEEP/ACTIVATE и
незнакомые отчёту) остаются как есть.

Использование:
    python3 build_feed.py исходный_фид.xml feed.xml
"""
import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    src, dst = sys.argv[1], sys.argv[2]

    data = json.load(open(Path(__file__).parent / "reports" / "ad_id_map.json"))
    ad_action = {}  # ad_id (из фида) -> действие
    for avito_id, ad_id in data["map"].items():
        ad_action[str(ad_id)] = data["actions"].get(avito_id, "")

    tree = ET.parse(src)
    root = tree.getroot()
    ads = root.findall("Ad")
    print(f"в исходном фиде: {len(ads)} объявлений")

    counts = {"DELETE": 0, "KEEP": 0, "ACTIVATE": 0, "не в отчёте": 0}
    for ad in ads:
        el = ad.find("Id")
        ad_id = (el.text or "").strip() if el is not None else ""
        action = ad_action.get(ad_id, "не в отчёте")
        counts[action] = counts.get(action, 0) + 1
        if action == "DELETE":
            root.remove(ad)

    print("разбор:", counts)
    kept = root.findall("Ad")
    print(f"в новом фиде: {len(kept)} объявлений")

    # какие KEEP/ACTIVATE ad_id из карты не нашлись в фиде
    expected = {a for a, act in ad_action.items() if act in ("KEEP", "ACTIVATE")}
    present = set()
    for ad in kept:
        el = ad.find("Id")
        if el is not None and el.text:
            present.add(el.text.strip())
    missing = expected - present
    if missing:
        print(f"ВНИМАНИЕ: {len(missing)} KEEP/ACTIVATE ad_id нет в фиде: "
              f"{sorted(missing)[:10]}...")

    tree.write(dst, encoding="utf-8", xml_declaration=True)
    ET.parse(dst)  # валидация результата
    print(f"записан и провалидирован: {dst}")


if __name__ == "__main__":
    main()
