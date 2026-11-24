"""Stage-specific A–E methodology. G, parking and infrastructure are outside this module."""
from copy import deepcopy
from math import isfinite
import pandas as pd

RATE = "Ставка, руб/ед."
SECTION_COLS = ["Секция / стилобат", "Количество", "S этажа, м2", "Этажей", "Периметр, м", "Сблокированные торцы, м2", "Стилобат (1/0)"]
DETAIL_COLS = ["Категория", "Тип / материал / размеры", "Объем", RATE, "Аналог", "Дата цен", "Состав работ / комментарий"]
# Internal detail suffixes preserve the corporate parent code.
DETAIL_SPEC = {
    "Кладка наружных стен": ("B.20.10.01", "м² фасада"),
    "Отделка фасада": ("B.20.10.02", "м² фасада"),
    "Ограждения балконов и лоджий": ("B.20.10.03", "м² ограждения"),
    "Корзины кондиционеров": ("B.20.10.04", "шт"),
    "Экраны": ("B.20.10.05", "шт"),
    "Входные группы по площади": ("B.20.10.06", "м² отделки входных групп"),
    "Входные группы без площади": ("B.20.10.07", "шт"),
    "Террасы первого этажа": ("B.20.10.08", "м² террасы"),
    "Террасы верхних этажей": ("B.20.10.09", "м² террасы"),
}
MP_ITEMS = [
 ("A.10", "Фундаменты", "A", "руб/м² застройки", "MP_FOOTPRINT", 0),
 ("A.20", "Подземная часть", "A", "руб/м² застройки секций", "MP_UNDERGROUND", 0),
 ("B.10", "Несущие конструкции", "B", "руб/м² надземной площади", "MP_ABOVE", 0),
 ("B.20", "Фасады с остеклением (комплексно)", "B", "руб/м² фасада", "MP_FACADE", 0),
 ("B.30", "Кровля", "B", "руб/м² застройки", "MP_FOOTPRINT", 0),
 ("C.10", "Общестроительные работы", "C", "руб/м² ПП", "NSA", 0),
 ("C.20", "Лестницы", "C", "руб/м² ПП", "NSA", 0),
 ("C.30", "Предчистовая отделка квартир, чистовая МОП", "C", "руб/м² ПП", "NSA", 0),
 ("C.40", "Чистовая отделка квартир", "C", "руб/м² ПП", "NSA", 0),
 ("D.10", "Лифты", "D", "руб/м² ПП", "NSA", 0),
 ("D.20", "Системы ВК, газоснабжение", "D", "руб/м² ПП", "NSA", 0),
 ("D.30", "Отопление, вентиляция, ИТП и котельная", "D", "руб/м² ПП", "NSA", 0),
 ("D.40", "Противопожарные системы", "D", "руб/м² ПП", "NSA", 0),
 ("D.50", "Электрические сети и оборудование", "D", "руб/м² ПП", "NSA", 0),
 ("E.10", "Технологическое оборудование", "E", "руб/м² ПП", "NSA", 0),
 ("E.20", "Мебель и аксессуары МОП", "E", "руб/секцию", "MP_SECTIONS", 0),
]
EXTRA_PARAMS = [
 ("Дверей МОП, шт", "Двери МОП", "шт"),
 ("Дверей технических помещений, шт", "Двери техпомещений", "шт"),
 ("Дверей входных групп, шт", "Двери входных групп", "шт"),
 ("Площадь холлов МОП, м2", "Холлы МОП", "м²"),
 ("ИТП, комплектов на этап", "ИТП на этап", "компл."),
 ("Котельная, комплектов на этап", "Котельная на этап", "компл."),
]

def number(v):
    try:
        n = float(v)
        return n if isfinite(n) else 0.0
    except (TypeError, ValueError):
        return 0.0


def stage_catalog(ep_items):
    result = []
    for code, name, group, unit, basis, default in ep_items:
        if code == "B.20.10":
            for category, (child, child_unit) in DETAIL_SPEC.items():
                result.append((child, category, "B", "руб/" + child_unit, child, 0))
        elif code == "B.20.30":
            result.append((code, "Квартирные входные двери", group, "руб/дверь", "APT_COUNT", 0))
            for suffix, label, key in [("01", "Двери МОП", "Дверей МОП, шт"), ("02", "Двери технических помещений", "Дверей технических помещений, шт"), ("03", "Двери входных групп", "Дверей входных групп, шт")]:
                result.append((code + "." + suffix, label, group, "руб/дверь", key, 0))
        elif code == "D.30.30":
            result.extend([
                ("D.30.20.01", "ИТП — комплект на этап", "D", "руб/комплект", "ИТП, комплектов на этап", 0),
                ("D.30.20.02", "Котельная — комплект на этап", "D", "руб/комплект", "Котельная, комплектов на этап", 0),
            ])
            result.append((code, name, group, unit, basis, default))
        elif code == "E.20.10":
            result.append((code, name, group, "руб/м² холла", "Площадь холлов МОП, м2", 0))
        elif code == "E.20.20":
            continue  # legacy additional design/fitout is not a second furniture charge
        else:
            result.append((code, name, group, unit, basis, default))
    return result


def rates_frame(items):
    return pd.DataFrame([{"Код": c, "Статья затрат": n, "Группа": g, "Единица измерения": u, "_basis": b,
                          RATE: 0.0, "Применимость": "Уточнить", "Аналог": "", "Дата цен": "", "Состав работ / комментарий": ""}
                         for c, n, g, u, b, _ in items])


def normalise_rates(old, items, stage):
    frame = rates_frame(items)
    if old is None or "Код" not in old:
        return frame
    old = old.drop_duplicates("Код", keep="last").set_index("Код")
    for i, row in frame.iterrows():
        code = row["Код"]
        if code in old.index:
            for col in (RATE, "Применимость", "Аналог", "Дата цен", "Состав работ / комментарий"):
                if col in old:
                    val = old.at[code, col]
                    frame.at[i, col] = number(val) if col == RATE else (str(val) if pd.notna(val) else ("Уточнить" if col == "Применимость" else ""))
    return frame


def geometry(rows):
    totals = dict(MP_FOOTPRINT=0.0, MP_UNDERGROUND=0.0, MP_ABOVE=0.0, MP_FACADE=0.0, MP_SECTIONS=0.0)
    issues = []
    for i, row in enumerate(rows):
        n, area, floors, perimeter, shared, podium = [number(row.get(k)) for k in SECTION_COLS[1:]]
        if any(v < 0 for v in (n, area, floors, perimeter, shared)) or n % 1 or floors % 1 or podium not in (0, 1):
            issues.append(f"Строка {i + 1}: количество и этажность — целые неотрицательные; стилобат — 0 или 1.")
        gross = n * perimeter * floors * 3.0
        if shared > gross:
            issues.append(f"Строка {i + 1}: площадь сблокированных торцов больше площади фасада.")
        if n > 0 and (area <= 0 or floors <= 0 or perimeter <= 0):
            issues.append(f"Строка {i + 1}: заполните площадь этажа, этажность и периметр.")
        totals["MP_FOOTPRINT"] += n * area
        totals["MP_UNDERGROUND"] += n * area * (1 - podium)
        totals["MP_ABOVE"] += n * area * floors
        totals["MP_FACADE"] += max(0.0, gross - shared)
        totals["MP_SECTIONS"] += n * (1 - podium)
    return totals, issues


def detail_totals(rows):
    amounts, qty, issues = {}, {}, []
    entries = set()
    for i, row in enumerate(rows):
        category = row.get("Категория")
        if category not in DETAIL_SPEC:
            if category or number(row.get("Объем")) or number(row.get(RATE)):
                issues.append(f"Строка {i + 1}: выберите категорию работ.")
            continue
        code, unit = DETAIL_SPEC[category]
        q, rate = number(row.get("Объем")), number(row.get(RATE))
        if q < 0 or rate < 0 or (unit == "шт" and q % 1):
            issues.append(f"Строка {i + 1}: проверьте объём и ставку; штучные объёмы должны быть целыми.")
        amounts[code] = amounts.get(code, 0) + q * rate
        qty[code] = qty.get(code, 0) + q
        if q > 0 and rate <= 0:
            issues.append(f"Строка {i + 1}: не заполнена ставка СЭО.")
        if rate > 0 and q <= 0:
            issues.append(f"Строка {i + 1}: не заполнен объём.")
        if q > 0 and code in ("B.20.10.06", "B.20.10.07"):
            entries.add(code)
    if len(entries) > 1:
        issues.append("Входные группы заданы по площади и поштучно: проверьте, что это разные группы.")
    return amounts, qty, issues


def legacy_issues(old):
    if old is None or "Код" not in old:
        return []
    rates = old.set_index("Код")[RATE]
    result = []
    for code, reason in [("B.20.10", "комплексная ставка фасада заменена детализацией по видам работ"),
                         ("E.20.10", "единица мебели МОП изменена с секции на м² холла"),
                         ("E.20.20", "дополнительная дизайн-отделка исключена из повторного начисления мебели")]:
        if number(rates.get(code)):
            result.append(f"{code}: {reason}; прежнее значение сохранено в архиве, требуется новая ставка.")
    return result


def effective_rates(frame):
    result = pd.to_numeric(frame.set_index("Код")[RATE], errors="coerce").fillna(0.0)
    if "Применимость" in frame:
        off = frame.loc[frame["Применимость"] == "Не требуется", "Код"]
        result.loc[result.index.isin(off)] = 0.0
    return result
