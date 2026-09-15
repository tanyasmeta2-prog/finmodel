# -*- coding: utf-8 -*-
"""
Финансовая модель девелоперского проекта (Streamlit + openpyxl)
=================================================================
Универсальный расчет экономики проекта (от 1 до N блоков) на уровне
Валовой прибыли и Валовой рентабельности (без налогов и кредитов).

Версия 4 — добавлена вкладка "СМР по методике": детальный расчет
себестоимости коробки жилого блока по укрупненным видам работ (коды
затрат A–E из действующей методики компании), вместо одной блендовой
ставки за м2. Расчет по методике АВТОМАТИЧЕСКИ заменяет прямые затраты
жилого блока в основной модели.

Единая таблица ТЭП, где каждая строка (урбан-блок) имеет "Тип блока":
  • "Жилой блок" — жилье, коммерция 1 эт., кладовые, подземный паркинг +
    детальная себестоимость коробки по методике (вкладка 2). Именно на
    эти строки распределяется пул косвенных расходов площадки,
    пропорционально их суммарной продаваемой площади (NSA).
  • "Наземный/Многоуровневый паркинг" — отдельно стоящий паркинг,
    считается по прямой ставке за 1 машиноместо, без методики по коробке
    и без аллокации косвенных расходов.

Экспорт в Excel строится ДВУМЯ листами с живыми формулами:
  1. "Экономика проекта" — сводная таблица по блокам.
  2. "СМР по методике" — каталог расценок + расчет себестоимости коробки
     по видам работ на каждый жилой блок, с перекрестными ссылками на
     лист 1 (единый источник входных площадей/объемов).

Запуск:
    streamlit run app.py
"""

import io
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ======================================================================
# 0. ЦВЕТОВАЯ ПАЛИТРА (фиксированный порядок слотов — единая для веб- и
#    Excel-графиков)
# ======================================================================
PALETTE = {
    "blue": "2a78d6", "orange": "eb6834", "aqua": "1baf7a", "yellow": "eda100",
    "magenta": "e87ba4", "violet": "4a3aa7", "red": "e34948", "green": "008300",
}
COLOR_REVENUE = PALETTE["blue"]
COLOR_COST = PALETTE["orange"]
COLOR_POSITIVE = PALETTE["blue"]
COLOR_NEGATIVE = PALETTE["red"]
COLOR_DIRECT_COST = PALETTE["blue"]
COLOR_ALLOC_COST = PALETTE["violet"]
PIE_COLORS = [PALETTE["blue"], PALETTE["orange"], PALETTE["aqua"], PALETTE["yellow"], PALETTE["magenta"]]
GROUP_COLORS = [PALETTE["blue"], PALETTE["orange"], PALETTE["aqua"], PALETTE["yellow"], PALETTE["magenta"]]

# ======================================================================
# 1. НАСТРОЙКИ СТРАНИЦЫ И СЦЕНАРИИ
# ======================================================================
st.set_page_config(page_title="Финмодель девелоперского проекта", layout="wide")

TYPE_RESIDENTIAL = "Жилой блок"
TYPE_PARKING = "Наземный/Многоуровневый паркинг"
BLOCK_TYPES = [TYPE_RESIDENTIAL, TYPE_PARKING]

SCENARIOS = {
    "Базовый": {"revenue": 1.00, "cost": 1.00},
    "Стресс": {"revenue": 0.85, "cost": 1.15},
    "Оптимистичный": {"revenue": 1.10, "cost": 0.95},
}

# ======================================================================
# 2. КАТАЛОГ РАСЦЕНОК ПО МЕТОДИКЕ (действующая методика компании,
#    активные статьи расчета себестоимости "коробки" — коды A-E).
#    Статьи со статусом "на данный момент не используем" в методике
#    в каталог не включены. Наружные работы (G) и прочие затраты по
#    СМР (Z, генподряд/непредвиденные и т.п.) остаются в пуле косвенных
#    расходов (сайдбар) либо в ставках паркингов — они не привязаны к
#    коробке конкретного жилого блока.
# ======================================================================
# Ключи базы расчета — либо имя колонки исходных данных блока, либо
# служебный код NSA/VOL_TOTAL, вычисляемый на лету.
BASIS_COLUMN = {
    "AREA_FOOTPRINT": "Площадь застройки, м2",
    "VOL_BELOW": "Объем здания ниже 0, м3",
    "VOL_ABOVE": "Объем здания выше 0, м3",
    "AREA_FACADE": "Площадь фасада, м2",
    "AREA_GLAZING_WINDOWS": "Площадь остекления окон, м2",
    "AREA_GLAZING_BALCONY": "Площадь остекления лоджий, м2",
    "APT_COUNT": "Кол-во квартир, шт",
    "ENTRANCE_COUNT": "Кол-во подъездов, шт",
    "ELEVATOR_COUNT": "Кол-во лифтов, шт",
}
PARAM_COLS = list(BASIS_COLUMN.values())  # 9 доп. параметров — новые колонки таблицы блоков

# (код, наименование статьи, группа, единица измерения, ключ базы, ставка по умолчанию)
ITEMS = [
    ("A.10.10", "Типовые фундаменты (ростверк)", "A", "руб/м² площади застройки", "AREA_FOOTPRINT", 8000),
    ("A.10.20", "Специализированные работы (сваи)", "A", "руб/м² площади застройки", "AREA_FOOTPRINT", 6000),
    ("A.10.30", "Фундаментная плита", "A", "руб/м² площади застройки", "AREA_FOOTPRINT", 12000),
    ("A.20.10", "Земляные работы", "A", "руб/м³ объема ниже 0", "VOL_BELOW", 2500),
    ("A.20.20", "Конструкции подземной части", "A", "руб/м³ объема ниже 0", "VOL_BELOW", 18000),
    ("A.—", "Кладовые (перегородки, отделка)", "A", "руб/м² кладовых", "STORAGE_AREA", 15000),
    ("B.10.10", "Несущий каркас и плиты перекрытий", "B", "руб/м³ объема выше 0", "VOL_ABOVE", 9500),
    ("B.20.10", "Наружные стены и фасады", "B", "руб/м² фасада", "AREA_FACADE", 14000),
    ("B.20.20", "Заполнение оконных проемов", "B", "руб/м² остекления окон", "AREA_GLAZING_WINDOWS", 18000),
    ("B.20.30", "Заполнение дверных проемов", "B", "руб/квартиру", "APT_COUNT", 45000),
    ("B.20.40", "Остекление лоджий и балконов", "B", "руб/м² остекления лоджий", "AREA_GLAZING_BALCONY", 16000),
    ("B.30.10", "Кровельные покрытия", "B", "руб/м² площади застройки", "AREA_FOOTPRINT", 4500),
    ("C.10.10", "Перегородки", "C", "руб/м² NSA", "NSA", 3200),
    ("C.20.10", "Конструкции лестниц", "C", "руб/м³ объема здания", "VOL_TOTAL", 21000),
    ("C.20.20", "Отделка лестниц", "C", "руб/м³ объема здания", "VOL_TOTAL", 6500),
    ("C.30.10", "Отделка стен (МОП/предчистовая)", "C", "руб/м² NSA", "NSA", 2800),
    ("C.30.20", "Отделка полов", "C", "руб/м² NSA", "NSA", 2200),
    ("C.30.30", "Отделка потолков", "C", "руб/м² NSA", "NSA", 1600),
    ("D.10.10", "Лифты", "D", "руб/лифт", "ELEVATOR_COUNT", 3800000),
    ("D.20.10", "Сантехническое оборудование", "D", "руб/квартиру", "APT_COUNT", 25000),
    ("D.20.20", "Водоснабжение", "D", "руб/м² NSA", "NSA", 900),
    ("D.20.30", "Хоз.быт. канализация", "D", "руб/м² NSA", "NSA", 850),
    ("D.20.40", "Ливневая канализация", "D", "руб/подъезд", "ENTRANCE_COUNT", 180000),
    ("D.20.90", "Технологические трубопроводы, газоснабжение", "D", "руб/м² NSA", "NSA", 600),
    ("D.30.30", "Система отопления", "D", "руб/м² NSA", "NSA", 1400),
    ("D.30.40", "Вентиляция и кондиционирование", "D", "руб/м² NSA", "NSA", 1100),
    ("D.40.40", "Система АПС", "D", "руб/м² NSA", "NSA", 450),
    ("D.50.10", "Силовое оборудование и сети", "D", "руб/м² NSA", "NSA", 1800),
    ("D.50.20", "Сети освещения", "D", "руб/м² NSA", "NSA", 350),
    ("D.50.30", "Слаботочные сети", "D", "руб/м² NSA", "NSA", 400),
    ("E.20.10", "Встроенная мебель МОП", "E", "руб/подъезд", "ENTRANCE_COUNT", 650000),
    ("E.20.20", "Дизайн-отделка холлов МОП", "E", "руб/подъезд", "ENTRANCE_COUNT", 900000),
]
GROUP_LABELS = {
    "A": "Подземная часть", "B": "Конструкции", "C": "Внутренние работы",
    "D": "Инженерные системы", "E": "Оборудование",
}

DEFAULT_RATES_DF = pd.DataFrame(
    [{"Код": c, "Статья затрат": n, "Группа": g, "Единица измерения": u, "_basis": b, "Ставка, руб/ед.": r}
     for c, n, g, u, b, r in ITEMS]
)

# ----------------------------------------------------------------------
# Наружные работы (код G) и прочие затраты, связанные с СМР (код Z) —
# ПРОЕКТНЫЙ уровень (считаются на весь проект, а не по блокам), из той
# же методики. Статьи со статусом "на данный момент не используем" не
# включены. G.20.20 «Автостоянки» тоже не включена — паркинг уже считается
# отдельно по блокам (ставки СМР подземного/наземного м/м), включение
# этой статьи задвоило бы затраты.
# ----------------------------------------------------------------------
# Статьи G на площадь участка — количество НЕ вводится по каждой статье отдельно,
# а берется ОДНО значение площади участка (лист «Исходные данные», ТЭП по мастер-плану).
ITEMS_G_AREA = [
    # (код, статья затрат, единица измерения, ставка по умолчанию)
    ("G.10.10", "Очистка площадки", "площадь участка, га", 250_000),
    ("G.10.20", "Разборка и вывоз сооружений", "площадь участка, га", 180_000),
    ("G.10.30", "Земляные работы по подготовке площадки", "площадь участка, га", 650_000),
    ("G.10.50", "Инженерная подготовка территории", "площадь участка, га", 400_000),
    ("G.20.10", "Дороги", "площадь участка, га", 3_500_000),
    ("G.20.40", "МАФ, детские площадки, ограждение", "площадь участка, га", 1_200_000),
    ("G.20.50", "Озеленение", "площадь участка, га", 900_000),
    ("G.40.20", "Наружное освещение", "площадь участка, га", 350_000),
]
DEFAULT_G_AREA_DF = pd.DataFrame(
    [{"Код": c, "Статья затрат": n, "Единица измерения": u, "Ставка, руб/ед.": r} for c, n, u, r in ITEMS_G_AREA]
)

# Статьи G по сетям/кабелям — своей ТЭП-базы (длины) в методике нет, кол-во вводится вручную.
ITEMS_G_LENGTH = [
    # (код, статья затрат, единица измерения, кол-во по умолчанию, ставка по умолчанию)
    ("G.30.10", "Водоснабжение (наружные сети)", "п.м. трубопровода", 600, 18_000),
    ("G.30.20", "Хоз.быт. канализация (наружные сети)", "п.м. трубопровода", 600, 20_000),
    ("G.30.30", "Ливневая канализация (наружные сети)", "п.м. трубопровода", 500, 16_000),
    ("G.30.40", "Сети теплоснабжения", "п.м. трубопровода", 550, 22_000),
    ("G.30.50", "Сети холодоснабжения", "п.м. трубопровода", 0, 22_000),
    ("G.30.60", "Топливоснабжение (газ)", "п.м. трубопровода", 0, 15_000),
    ("G.40.10", "Сети электроснабжения, ТП", "п.м. кабеля", 400, 25_000),
]
DEFAULT_G_LENGTH_DF = pd.DataFrame(
    [{"Код": c, "Статья затрат": n, "Единица измерения": u, "Кол-во": q, "Ставка, руб/ед.": r}
     for c, n, u, q, r in ITEMS_G_LENGTH]
)

ITEMS_Z_PCT = [
    # (код, статья затрат, ставка по умолчанию — доля от (СМР коробки + G))
    ("Z.10.10", "Услуги генподрядчика", 0.06),
    ("Z.20.10", "Непредвиденные расходы по объекту", 0.07),
]
DEFAULT_Z_PCT_DF = pd.DataFrame(
    [{"Код": c, "Статья затрат": n, "Ставка, доля от СМР+G": r} for c, n, r in ITEMS_Z_PCT]
)

ITEMS_Z_FIXED = [
    # (код, статья затрат, сумма по умолчанию, руб) — по методике формульной базы
    # нет, сумма берется "по объектам-аналогам" — вводится вручную.
    ("Z.10.20", "Контроль качества выполнения работ", 15_000_000),
    ("Z.10.30", "Временные здания и сооружения на площадке", 20_000_000),
    ("Z.10.40", "Коммунальные услуги на площадке в период строительства", 10_000_000),
    ("Z.10.50", "Услуги сторонних организаций (межевание, кадастр, ТУ)", 8_000_000),
    ("Z.10.60", "Проектирование, изыскания, авторский надзор", 45_000_000),
    ("Z.150", "Технологическое присоединение", 25_000_000),
]
DEFAULT_Z_FIXED_DF = pd.DataFrame(
    [{"Код": c, "Статья затрат": n, "Сумма, руб": s} for c, n, s in ITEMS_Z_FIXED]
)


def get_basis_values(basis_key: str, df: pd.DataFrame, nsa: np.ndarray) -> np.ndarray:
    """Возвращает массив значений базы расчета (площадь/объем/шт) по каждой строке блоков."""
    if basis_key == "NSA":
        return nsa
    if basis_key == "VOL_TOTAL":
        return (df["Объем здания ниже 0, м3"] + df["Объем здания выше 0, м3"]).to_numpy(dtype=float)
    if basis_key == "STORAGE_AREA":
        return df["S кладовых, м2"].to_numpy(dtype=float)
    return df[BASIS_COLUMN[basis_key]].to_numpy(dtype=float)


# ----------------------------------------------------------------------
# Лист «Исходные данные (ТЭП)» — из методики, отдельный лист «Перечень ТЭП
# для МП и ЭП». Включены ТОЛЬКО числовые поля, которые реально участвуют
# в расчете (остальные пункты методики текстовые/описательные — тип
# вентиляции, СТУ и т.п. — и в модели не используются).
# Поля ЭП (по каждому жилому блоку) передаются в соответствующие колонки
# таблицы блоков — там они становятся нередактируемыми.
# ----------------------------------------------------------------------
TEP_TO_MAIN_COL = {
    "Кол-во секций, шт": "Кол-во подъездов, шт",
    "Кол-во квартир, шт": "Кол-во квартир, шт",
    "Общая площадь квартир (с летними, с коэф.), м2": "S квартир, м2",
    "Площадь коммерции, м2": "S коммерции 1 эт., м2",
    "Площадь застройки, м2": "Площадь застройки, м2",
    "Площадь кладовых в доме, м2": "S кладовых, м2",
    "Площадь остекления окон, м2": "Площадь остекления окон, м2",
    "Кол-во лифтов, шт": "Кол-во лифтов, шт",
}
TEP_COLS = list(TEP_TO_MAIN_COL.keys())


# ======================================================================
# 3. КОЛОНКИ ЕДИНОЙ ТАБЛИЦЫ ТЭП И ТЕСТОВЫЕ ДАННЫЕ
# ======================================================================
NUMERIC_COLS = [
    "S квартир, м2", "S коммерции 1 эт., м2", "S кладовых, м2",
    "Подземный паркинг, м/м", "Наземный/Многоуровневый паркинг, м/м",
    "Цена жилья, руб/м2", "Цена коммерции, руб/м2", "Цена кладовых, руб/м2",
    "Цена подземного м/м, руб", "Цена наземного м/м, руб",
    "Ставка СМР подземного м/м, руб", "Ставка СМР наземного м/м, руб",
] + PARAM_COLS  # + 9 доп. параметров для расчета СМР коробки по методике
ALL_COLS = ["Название блока", "Тип блока"] + NUMERIC_COLS


def generate_default_table() -> pd.DataFrame:
    """Тестовый набор: 13 жилых урбан-блоков + 2 блока-паркинга (детерминированно)."""
    rng = np.random.default_rng(11)
    rows = []

    for i in range(1, 14):
        s_apt = int(rng.integers(6000, 16000) // 100 * 100)
        s_c1 = int(rng.integers(150, 700) // 10 * 10)
        s_storage = int(rng.integers(50, 400) // 10 * 10)
        parking_u = int(rng.integers(60, 170) // 5 * 5)

        # -- Доп. параметры для методики СМР, выведенные из площадей блока --
        footprint = int(rng.integers(500, 1200))
        vol_below = int(footprint * rng.uniform(4.0, 6.0))
        nsa_est = s_apt + s_c1 + s_storage
        vol_above = int(nsa_est * rng.uniform(3.2, 3.8))
        area_facade = int(footprint * rng.uniform(2.0, 3.0))
        area_glz_win = int(nsa_est * rng.uniform(0.15, 0.25))
        apt_count = max(1, int(s_apt / rng.uniform(45, 65)))
        area_glz_balcony = int(apt_count * rng.uniform(3, 6))
        entrance_count = max(1, int(apt_count / rng.uniform(60, 100)))
        elevator_count = max(1, int(round(entrance_count * rng.uniform(2, 3))))

        rows.append({
            "Название блока": f"УБ {i}", "Тип блока": TYPE_RESIDENTIAL,
            "S квартир, м2": s_apt, "S коммерции 1 эт., м2": s_c1, "S кладовых, м2": s_storage,
            "Подземный паркинг, м/м": parking_u, "Наземный/Многоуровневый паркинг, м/м": 0,
            "Цена жилья, руб/м2": int(rng.integers(150, 221) * 1000),
            "Цена коммерции, руб/м2": int(rng.integers(200, 321) * 1000),
            "Цена кладовых, руб/м2": int(rng.integers(40, 91) * 1000),
            "Цена подземного м/м, руб": int(rng.integers(700, 1001) * 1000),
            "Цена наземного м/м, руб": 0,
            "Ставка СМР подземного м/м, руб": int(rng.integers(650, 821) * 1000),
            "Ставка СМР наземного м/м, руб": 0,
            "Площадь застройки, м2": footprint,
            "Объем здания ниже 0, м3": vol_below,
            "Объем здания выше 0, м3": vol_above,
            "Площадь фасада, м2": area_facade,
            "Площадь остекления окон, м2": area_glz_win,
            "Площадь остекления лоджий, м2": area_glz_balcony,
            "Кол-во квартир, шт": apt_count,
            "Кол-во подъездов, шт": entrance_count,
            "Кол-во лифтов, шт": elevator_count,
        })

    for name, count in [("УБ-Паркинг 1", 1494), ("УБ-Паркинг 2", 1437)]:
        row = {col: 0 for col in NUMERIC_COLS}
        row.update({
            "Название блока": name, "Тип блока": TYPE_PARKING,
            "Наземный/Многоуровневый паркинг, м/м": count,
            "Цена наземного м/м, руб": 450000,
            "Ставка СМР наземного м/м, руб": 320000,
        })
        rows.append(row)

    return pd.DataFrame(rows)[ALL_COLS]


# ======================================================================
# 4. ПАСПОРТ ПРОЕКТА
# ======================================================================
st.title("Финансовая модель девелоперского проекта")

pass_col1, pass_col2 = st.columns(2)
with pass_col1:
    project_name = st.text_input("Наименование проекта", value="ЖК «Пример»")
with pass_col2:
    project_city = st.text_input("Город", value="Самара")

st.caption("Валовая прибыль и валовая рентабельность. Налоги и кредиты не учитываются.")

# ======================================================================
# 5. БОКОВАЯ ПАНЕЛЬ — СЦЕНАРИЙ И ПУЛ КОСВЕННЫХ РАСХОДОВ
# ======================================================================
with st.sidebar:
    st.header("Сценарий расчета")
    scenario_name = st.selectbox("Выберите сценарий", list(SCENARIOS.keys()))
    rev_factor = SCENARIOS[scenario_name]["revenue"]
    cost_factor = SCENARIOS[scenario_name]["cost"]

    st.header("Пул косвенных расходов проекта, руб")
    st.caption(
        "Земля/сети/благоустройство/социалка/soft costs — распределяется ТОЛЬКО на «Жилые "
        "блоки» пропорц. NSA. Наружные работы (G) и прочие затраты по СМР (Z) считаются "
        "детально на вкладке «СМР по методике» и добавляются к этому пулу автоматически — "
        "если поля «Сети» и «Благоустройство» ниже уже их учитывают, скорректируйте суммы, "
        "чтобы не задвоить."
    )
    cost_land = st.number_input("Земля", min_value=0.0, value=500_000_000.0, step=1_000_000.0)
    cost_infra = st.number_input("Сети / Инфраструктура", min_value=0.0, value=300_000_000.0, step=1_000_000.0)
    cost_landscape = st.number_input("Благоустройство / Дороги", min_value=0.0, value=150_000_000.0, step=1_000_000.0)
    cost_social = st.number_input("Социальные объекты (школы/сады)", min_value=0.0, value=400_000_000.0, step=1_000_000.0)
    cost_soft = st.number_input("Прочие Soft Costs", min_value=0.0, value=100_000_000.0, step=1_000_000.0)
    indirect_pool_sidebar = cost_land + cost_infra + cost_landscape + cost_social + cost_soft
    st.caption(f"Итого по этим статьям: {indirect_pool_sidebar:,.0f} руб (без G/Z)".replace(",", " "))

# ======================================================================
# 6. ЕДИНАЯ ТАБЛИЦА ТЭП + КАТАЛОГ РАСЦЕНОК (session_state)
# ======================================================================
if "blocks_df" not in st.session_state:
    st.session_state.blocks_df = generate_default_table()
if "block_rates" not in st.session_state:
    st.session_state.block_rates = {}  # имя жилого блока -> DataFrame ставок (32 статьи)
if "g_area_df" not in st.session_state:
    st.session_state.g_area_df = DEFAULT_G_AREA_DF.copy()
if "g_length_df" not in st.session_state:
    st.session_state.g_length_df = DEFAULT_G_LENGTH_DF.copy()
if "z_pct_df" not in st.session_state:
    st.session_state.z_pct_df = DEFAULT_Z_PCT_DF.copy()
if "z_fixed_df" not in st.session_state:
    st.session_state.z_fixed_df = DEFAULT_Z_FIXED_DF.copy()
if "site_area_ga" not in st.session_state:
    st.session_state.site_area_ga = 3.5
if "tep_store" not in st.session_state:
    st.session_state.tep_store = {}  # имя жилого блока -> {поле ТЭП: значение}

tab_main, tab_smr, tab_tep = st.tabs(
    ["📊 Финансовая модель", "🏗️ СМР по методике", "📋 Исходные данные (ТЭП)"]
)

# ------------------------------------------------------------------
# ВКЛАДКА 1: основная таблица ТЭП
# ------------------------------------------------------------------
with tab_main:
    st.subheader("ТЭП проекта (жилые блоки и блоки-паркинги в одной таблице)")
    st.caption(
        "Себестоимость коробки жилого блока считается на вкладке «СМР по методике» — "
        "здесь задаются площади, цены и доп. параметры для этого расчета. Колонки, "
        "выделенные на вкладке «Исходные данные» (кол-во квартир, площади и т.п.), "
        "вводятся там и здесь не редактируются."
    )

    tep_linked_cols = set(TEP_TO_MAIN_COL.values())
    column_config = {
        "Название блока": st.column_config.TextColumn(required=True),
        "Тип блока": st.column_config.SelectboxColumn(options=BLOCK_TYPES, required=True),
    }
    for col in NUMERIC_COLS:
        column_config[col] = st.column_config.NumberColumn(
            min_value=0, format="%.0f", disabled=(col in tep_linked_cols)
        )

    edited = st.data_editor(
        st.session_state.blocks_df,
        num_rows="dynamic",
        use_container_width=True,
        key="blocks_editor",
        column_config=column_config,
    )
    st.session_state.blocks_df = edited

blocks = edited.copy()
for col in NUMERIC_COLS:
    blocks[col] = pd.to_numeric(blocks[col], errors="coerce").fillna(0.0)
blocks["Название блока"] = blocks["Название блока"].fillna("").astype(str)
blocks["Тип блока"] = blocks["Тип блока"].fillna(TYPE_RESIDENTIAL)
blocks = blocks[~((blocks["Название блока"] == "") & (blocks[NUMERIC_COLS].sum(axis=1) == 0))]
blocks = blocks.reset_index(drop=True)
N_ROWS = len(blocks)

is_res = (blocks["Тип блока"] == TYPE_RESIDENTIAL)
is_park = (blocks["Тип блока"] == TYPE_PARKING)

# ------------------------------------------------------------------
# ВКЛАДКА 3: исходные данные (ТЭП) по методике — источник для части
# параметров жилых блоков (кол-во квартир, площади и т.п., ЭП) и площади
# участка проекта (МП), которая используется в разделе «Наружные работы» (G).
# ------------------------------------------------------------------
res_block_names = list(blocks.loc[is_res, "Название блока"])
st.session_state.tep_store = {k: v for k, v in st.session_state.tep_store.items() if k in res_block_names}
for _i, _name in enumerate(blocks["Название блока"]):
    if is_res[_i] and _name not in st.session_state.tep_store:
        st.session_state.tep_store[_name] = {
            tep_col: float(blocks.at[_i, main_col]) for tep_col, main_col in TEP_TO_MAIN_COL.items()
        }

with tab_tep:
    st.subheader("Исходные данные по мастер-плану (МП)")
    st.caption(
        "Площадь участка — единая на проект. Используется в расчете раздела «Наружные "
        "работы» (G) на вкладке «СМР по методике»."
    )
    st.session_state.site_area_ga = st.number_input(
        "Площадь участка, га", min_value=0.0, value=float(st.session_state.site_area_ga), step=0.1,
    )

    st.subheader("Исходные данные по эскизному проекту (ЭП) — по каждому жилому блоку")
    st.caption(
        "Значения передаются в таблицу блоков на вкладке «Финансовая модель» (там эти "
        "колонки нередактируемые) и используются в расчете себестоимости коробки. Остальные "
        "параметры блока (цены, паркинг, объемы, фасад, лоджии) по-прежнему задаются на "
        "вкладке «Финансовая модель»."
    )
    if res_block_names:
        tep_df_view = pd.DataFrame(
            [{"Название блока": name, **st.session_state.tep_store[name]} for name in res_block_names]
        )
        tep_column_config = {"Название блока": st.column_config.TextColumn(disabled=True)}
        for _col in TEP_COLS:
            tep_column_config[_col] = st.column_config.NumberColumn(min_value=0, format="%.0f")
        tep_edited = st.data_editor(
            tep_df_view,
            use_container_width=True,
            num_rows="fixed",
            key="tep_editor",
            column_config=tep_column_config,
        )
        for _, _row in tep_edited.iterrows():
            st.session_state.tep_store[_row["Название блока"]] = {
                col: float(pd.to_numeric(_row[col], errors="coerce") or 0.0) for col in TEP_COLS
            }
    else:
        st.info("Добавьте хотя бы один «Жилой блок» на вкладке «Финансовая модель», чтобы ввести данные ЭП.")

# Данные ЭП — источник истины для соответствующих колонок таблицы блоков
for _i, _name in enumerate(blocks["Название блока"]):
    if is_res[_i] and _name in st.session_state.tep_store:
        for tep_col, main_col in TEP_TO_MAIN_COL.items():
            blocks.at[_i, main_col] = st.session_state.tep_store[_name][tep_col]
st.session_state.blocks_df = blocks[ALL_COLS].copy()

# NSA нужна и для аллокации, и как база нескольких статей методики
nsa = np.where(is_res, blocks["S квартир, м2"] + blocks["S коммерции 1 эт., м2"] + blocks["S кладовых, м2"], 0.0)
blocks["NSA, м2"] = nsa

# ------------------------------------------------------------------
# ВКЛАДКА 2: каталог расценок (ИНДИВИДУАЛЬНО по каждому жилому блоку) +
# расчет себестоимости коробки по методике
# ------------------------------------------------------------------
item_codes_master = list(DEFAULT_RATES_DF["Код"])
code_to_basis = dict(zip(DEFAULT_RATES_DF["Код"], DEFAULT_RATES_DF["_basis"]))
code_to_group = dict(zip(DEFAULT_RATES_DF["Код"], DEFAULT_RATES_DF["Группа"]))
code_to_name = dict(zip(DEFAULT_RATES_DF["Код"], DEFAULT_RATES_DF["Статья затрат"]))

# res_block_names уже посчитан выше (синхронизация с вкладкой «Исходные данные»).
# Убираем ставки блоков, которых больше нет в таблице (удалены/переименованы),
# и заводим ставки по умолчанию для новых блоков.
st.session_state.block_rates = {k: v for k, v in st.session_state.block_rates.items() if k in res_block_names}
for _name in res_block_names:
    if _name not in st.session_state.block_rates:
        st.session_state.block_rates[_name] = DEFAULT_RATES_DF.copy()

with tab_smr:
    st.subheader("Ставки СМР по видам работ — индивидуально по каждому урбан-блоку")
    st.caption(
        "У каждого жилого блока может быть своя себестоимость коробки — выберите блок и при "
        "необходимости скорректируйте его ставки. Код, группа, единица и база расчета едины "
        "по методике, редактируется только ставка. Новый блок получает ставки по умолчанию. "
        "Названия блоков должны быть уникальны, иначе ставки будут общими на все блоки с "
        "одинаковым названием."
    )
    if res_block_names:
        selected_block = st.selectbox("Урбан-блок", res_block_names, key="smr_block_selector")
        block_rates_edited = st.data_editor(
            st.session_state.block_rates[selected_block],
            use_container_width=True,
            num_rows="fixed",
            key=f"rates_editor_{selected_block}",
            column_config={
                "Код": st.column_config.TextColumn(disabled=True),
                "Статья затрат": st.column_config.TextColumn(disabled=True),
                "Группа": st.column_config.TextColumn(disabled=True),
                "Единица измерения": st.column_config.TextColumn(disabled=True),
                "_basis": None,  # служебная колонка — скрыта
                "Ставка, руб/ед.": st.column_config.NumberColumn(min_value=0, format="%.0f"),
            },
        )
        st.session_state.block_rates[selected_block] = block_rates_edited
    else:
        st.info("Добавьте хотя бы один «Жилой блок» на вкладке «Финансовая модель», чтобы задать ставки СМР.")

# -- Расчет себестоимости коробки по методике: у КАЖДОГО блока — свой каталог ставок --
qty_by_basis = {b: get_basis_values(b, blocks, nsa) for b in set(code_to_basis.values())}

item_cost_matrix = {code: np.zeros(N_ROWS) for code in item_codes_master}  # код -> np.array по блокам
block_rate_series = {}  # имя блока -> Series(код -> ставка) — для итогов/детализации
for i in range(N_ROWS):
    if not is_res[i]:
        continue
    name = blocks["Название блока"].iloc[i]
    block_df = st.session_state.block_rates.get(name, DEFAULT_RATES_DF)
    rate_series = pd.to_numeric(block_df.set_index("Код")["Ставка, руб/ед."], errors="coerce").fillna(0.0)
    block_rate_series[name] = rate_series
    for code in item_codes_master:
        item_cost_matrix[code][i] = qty_by_basis[code_to_basis[code]][i] * rate_series.get(code, 0.0)

smr_korobka_raw = np.sum(list(item_cost_matrix.values()), axis=0) if item_cost_matrix and N_ROWS > 0 else np.zeros(N_ROWS)
blocks["Себестоимость коробки (методика)"] = smr_korobka_raw
blocks["Эффективная ставка коробки, руб/м2"] = np.where(nsa > 0, smr_korobka_raw / np.where(nsa > 0, nsa, 1), 0.0)

with tab_smr:
    st.markdown("**Себестоимость коробки по блокам (сумма по статьям x количество)**")
    smr_result_df = pd.DataFrame({
        "Название блока": blocks["Название блока"],
        "Тип блока": blocks["Тип блока"],
        "NSA, м2": blocks["NSA, м2"],
        "Себестоимость коробки, руб": blocks["Себестоимость коробки (методика)"],
        "Эффективная ставка, руб/м2": blocks["Эффективная ставка коробки, руб/м2"],
    })
    st.dataframe(
        smr_result_df, use_container_width=True,
        column_config={
            "Себестоимость коробки, руб": st.column_config.NumberColumn(format="%.0f"),
            "Эффективная ставка, руб/м2": st.column_config.NumberColumn(format="%.0f"),
        },
    )

    smr_chart_col1, smr_chart_col2 = st.columns(2)
    with smr_chart_col1:
        res_only = smr_result_df[smr_result_df["Тип блока"] == TYPE_RESIDENTIAL]
        fig_smr_bar = go.Figure(
            go.Bar(x=res_only["Название блока"], y=res_only["Себестоимость коробки, руб"], marker_color=f"#{COLOR_DIRECT_COST}")
        )
        fig_smr_bar.update_layout(title="Себестоимость коробки по блокам (методика)", margin=dict(t=60, b=40))
        st.plotly_chart(fig_smr_bar, use_container_width=True)

    with smr_chart_col2:
        group_totals = {}
        for code in item_codes_master:
            g = GROUP_LABELS[code_to_group[code]]
            group_totals[g] = group_totals.get(g, 0.0) + item_cost_matrix[code].sum()
        fig_group_pie = go.Figure(
            go.Pie(
                labels=list(group_totals.keys()), values=list(group_totals.values()),
                marker=dict(colors=[f"#{c}" for c in GROUP_COLORS]), hole=0.35,
            )
        )
        fig_group_pie.update_layout(title="Структура СМР коробки по группам работ", margin=dict(t=60, b=20))
        st.plotly_chart(fig_group_pie, use_container_width=True)

    with st.expander("Детализация по каждой статье и блоку"):
        detail_df = pd.DataFrame(item_cost_matrix, index=blocks["Название блока"]).T
        detail_df.insert(0, "Статья затрат", [code_to_name[c] for c in detail_df.index])
        st.dataframe(detail_df, use_container_width=True)

# ------------------------------------------------------------------
# ВКЛАДКА 2 (продолжение): наружные работы (G) и прочие затраты, связанные
# с СМР (Z) — считаются на весь проект, не по блокам, из той же методики.
# ------------------------------------------------------------------
with tab_smr:
    st.divider()
    st.subheader("Наружные работы (код G, на весь проект)")
    st.caption(
        f"Статьи на площадь участка считаются от {st.session_state.site_area_ga:.1f} га "
        "(вкладка «Исходные данные», МП) — кол-во там не редактируется, только ставка. "
        "Автостоянки (G.20.20) не включены — паркинг уже учтен по блокам."
    )
    g_area_edited = st.data_editor(
        st.session_state.g_area_df,
        use_container_width=True,
        num_rows="fixed",
        key="g_area_editor",
        column_config={
            "Код": st.column_config.TextColumn(disabled=True),
            "Статья затрат": st.column_config.TextColumn(disabled=True),
            "Единица измерения": st.column_config.TextColumn(disabled=True),
            "Ставка, руб/ед.": st.column_config.NumberColumn(min_value=0, format="%.0f"),
        },
    )
    st.session_state.g_area_df = g_area_edited

g_area_df = st.session_state.g_area_df.copy()
g_area_df["Ставка, руб/ед."] = pd.to_numeric(g_area_df["Ставка, руб/ед."], errors="coerce").fillna(0.0)
g_area_df["Кол-во"] = float(st.session_state.site_area_ga)
g_area_df["Сумма, руб"] = g_area_df["Кол-во"] * g_area_df["Ставка, руб/ед."]

with tab_smr:
    st.dataframe(
        g_area_df[["Код", "Статья затрат", "Сумма, руб"]], use_container_width=True,
        column_config={"Сумма, руб": st.column_config.NumberColumn(format="%.0f")},
    )
    st.markdown("**Сети и кабели (своей ТЭП-базы нет — кол-во вводится вручную)**")
    g_length_edited = st.data_editor(
        st.session_state.g_length_df,
        use_container_width=True,
        num_rows="fixed",
        key="g_length_editor",
        column_config={
            "Код": st.column_config.TextColumn(disabled=True),
            "Статья затрат": st.column_config.TextColumn(disabled=True),
            "Единица измерения": st.column_config.TextColumn(disabled=True),
            "Кол-во": st.column_config.NumberColumn(min_value=0, format="%.1f"),
            "Ставка, руб/ед.": st.column_config.NumberColumn(min_value=0, format="%.0f"),
        },
    )
    st.session_state.g_length_df = g_length_edited

g_length_df = st.session_state.g_length_df.copy()
g_length_df["Кол-во"] = pd.to_numeric(g_length_df["Кол-во"], errors="coerce").fillna(0.0)
g_length_df["Ставка, руб/ед."] = pd.to_numeric(g_length_df["Ставка, руб/ед."], errors="coerce").fillna(0.0)
g_length_df["Сумма, руб"] = g_length_df["Кол-во"] * g_length_df["Ставка, руб/ед."]

g_total = float(g_area_df["Сумма, руб"].sum()) + float(g_length_df["Сумма, руб"].sum())

with tab_smr:
    st.metric("Итого наружные работы (G)", f"{g_total:,.0f} руб".replace(",", " "))

    st.subheader("Прочие затраты, связанные с СМР (код Z, на весь проект)")
    st.caption(
        "Часть статей считается как % от себестоимости СМР (коробка по блокам + наружные "
        "работы G). Остальные статьи по методике не имеют формульной базы — сумма берется "
        "«по объектам-аналогам» и вводится напрямую."
    )

z_pct_base = float(smr_korobka_raw.sum()) + g_total

with tab_smr:
    st.caption(f"База для % статей (СМР коробки + G): {z_pct_base:,.0f} руб".replace(",", " "))
    z_pct_edited = st.data_editor(
        st.session_state.z_pct_df,
        use_container_width=True,
        num_rows="fixed",
        key="z_pct_editor",
        column_config={
            "Код": st.column_config.TextColumn(disabled=True),
            "Статья затрат": st.column_config.TextColumn(disabled=True),
            "Ставка, доля от СМР+G": st.column_config.NumberColumn(min_value=0, max_value=1, format="%.3f"),
        },
    )
    st.session_state.z_pct_df = z_pct_edited

z_pct_df = st.session_state.z_pct_df.copy()
z_pct_df["Ставка, доля от СМР+G"] = pd.to_numeric(z_pct_df["Ставка, доля от СМР+G"], errors="coerce").fillna(0.0)
z_pct_df["Сумма, руб"] = z_pct_base * z_pct_df["Ставка, доля от СМР+G"]

with tab_smr:
    st.markdown("**Статьи с прямым вводом суммы (нет формульной базы по методике)**")
    z_fixed_edited = st.data_editor(
        st.session_state.z_fixed_df,
        use_container_width=True,
        num_rows="fixed",
        key="z_fixed_editor",
        column_config={
            "Код": st.column_config.TextColumn(disabled=True),
            "Статья затрат": st.column_config.TextColumn(disabled=True),
            "Сумма, руб": st.column_config.NumberColumn(min_value=0, format="%.0f"),
        },
    )
    st.session_state.z_fixed_df = z_fixed_edited

z_fixed_df = st.session_state.z_fixed_df.copy()
z_fixed_df["Сумма, руб"] = pd.to_numeric(z_fixed_df["Сумма, руб"], errors="coerce").fillna(0.0)

z_total = float(z_pct_df["Сумма, руб"].sum()) + float(z_fixed_df["Сумма, руб"].sum())

with tab_smr:
    st.metric("Итого прочие затраты, связанные с СМР (Z)", f"{z_total:,.0f} руб".replace(",", " "))
    st.caption(
        f"Итого G + Z: {(g_total + z_total):,.0f} руб — автоматически добавляется в пул "
        "косвенных расходов проекта (сайдбар) и распределяется на жилые блоки пропорц. NSA."
        .replace(",", " ")
    )

# ======================================================================
# 7. РАСЧЕТ ЭКОНОМИКИ (единая логика на всю таблицу, ветвление по типу)
# ======================================================================
indirect_pool_total = indirect_pool_sidebar + g_total + z_total

total_nsa = nsa.sum()
share = np.zeros_like(nsa, dtype=float)
if total_nsa > 0:
    share = nsa / total_nsa
blocks["Доля аллокации"] = share

indirect_pool_scenario = indirect_pool_total * cost_factor
blocks["Аллоцированные затраты"] = blocks["Доля аллокации"] * indirect_pool_scenario

# Прямые затраты: жилой блок = себестоимость коробки (по методике) + подземный
# паркинг блока; блок-паркинг = наземный/многоуровневый паркинг по своей ставке.
direct_res = (smr_korobka_raw + blocks["Подземный паркинг, м/м"] * blocks["Ставка СМР подземного м/м, руб"]) * cost_factor
direct_park = blocks["Наземный/Многоуровневый паркинг, м/м"] * blocks["Ставка СМР наземного м/м, руб"] * cost_factor
blocks["Прямые затраты"] = np.where(is_res, direct_res, direct_park)

blocks["Полные затраты"] = blocks["Прямые затраты"] + blocks["Аллоцированные затраты"]

revenue_res = (
    blocks["S квартир, м2"] * blocks["Цена жилья, руб/м2"]
    + blocks["S коммерции 1 эт., м2"] * blocks["Цена коммерции, руб/м2"]
    + blocks["S кладовых, м2"] * blocks["Цена кладовых, руб/м2"]
    + blocks["Подземный паркинг, м/м"] * blocks["Цена подземного м/м, руб"]
) * rev_factor
revenue_park = blocks["Наземный/Многоуровневый паркинг, м/м"] * blocks["Цена наземного м/м, руб"] * rev_factor
blocks["Выручка"] = np.where(is_res, revenue_res, revenue_park)

blocks["Валовая прибыль"] = blocks["Выручка"] - blocks["Полные затраты"]
blocks["Рентабельность"] = np.where(blocks["Выручка"] > 0, blocks["Валовая прибыль"] / blocks["Выручка"], 0.0)

revenue_components = {
    "Жилье": (blocks.loc[is_res, "S квартир, м2"] * blocks.loc[is_res, "Цена жилья, руб/м2"] * rev_factor).sum(),
    "Коммерция": (blocks.loc[is_res, "S коммерции 1 эт., м2"] * blocks.loc[is_res, "Цена коммерции, руб/м2"] * rev_factor).sum(),
    "Кладовые": (blocks.loc[is_res, "S кладовых, м2"] * blocks.loc[is_res, "Цена кладовых, руб/м2"] * rev_factor).sum(),
    "Подземные м/м": (blocks.loc[is_res, "Подземный паркинг, м/м"] * blocks.loc[is_res, "Цена подземного м/м, руб"] * rev_factor).sum(),
    "Наземные/Многоур. паркинги": (
        blocks.loc[is_park, "Наземный/Многоуровневый паркинг, м/м"] * blocks.loc[is_park, "Цена наземного м/м, руб"] * rev_factor
    ).sum(),
}

# ======================================================================
# 8. КОНСОЛИДИРОВАННЫЕ ПОКАЗАТЕЛИ ПРОЕКТА
# ======================================================================
total_revenue = blocks["Выручка"].sum()
total_cost = blocks["Полные затраты"].sum()
total_profit = total_revenue - total_cost
avg_margin = (total_profit / total_revenue) if total_revenue > 0 else 0.0

no_residential_warning = (blocks.shape[0] > 0) and (total_nsa == 0) and (indirect_pool_total > 0)

# ======================================================================
# 9. ВЕБ-ИНТЕРФЕЙС (вкладка 1) — МЕТРИКИ И ИНТЕРАКТИВНАЯ ГРАФИКА
# ======================================================================
with tab_main:
    st.divider()
    st.subheader(f"«{project_name}», {project_city} — сценарий «{scenario_name}»")

    if no_residential_warning:
        st.warning(
            "В таблице нет ни одного «Жилого блока» — пул косвенных расходов не на что "
            "распределить, поэтому он не учтен в итогах проекта."
        )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Выручка проекта", f"{total_revenue:,.0f} руб".replace(",", " "))
    m2.metric("Затраты проекта", f"{total_cost:,.0f} руб".replace(",", " "))
    m3.metric("Валовая прибыль", f"{total_profit:,.0f} руб".replace(",", " "))
    m4.metric("Средняя рентабельность", f"{avg_margin * 100:.1f} %")

    st.markdown("**Экономика проекта по блокам**")
    tbl_cols = [
        "Название блока", "Тип блока", "NSA, м2", "Доля аллокации",
        "Прямые затраты", "Аллоцированные затраты", "Полные затраты",
        "Выручка", "Валовая прибыль", "Рентабельность",
    ]
    st.dataframe(
        blocks[tbl_cols], use_container_width=True,
        column_config={
            "Доля аллокации": st.column_config.NumberColumn(format="percent"),
            "Рентабельность": st.column_config.NumberColumn(format="percent"),
            "Прямые затраты": st.column_config.NumberColumn(format="%.0f"),
            "Аллоцированные затраты": st.column_config.NumberColumn(format="%.0f"),
            "Полные затраты": st.column_config.NumberColumn(format="%.0f"),
            "Выручка": st.column_config.NumberColumn(format="%.0f"),
            "Валовая прибыль": st.column_config.NumberColumn(format="%.0f"),
        },
    )

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        fig_rev_cost = go.Figure()
        fig_rev_cost.add_bar(name="Выручка", x=blocks["Название блока"], y=blocks["Выручка"], marker_color=f"#{COLOR_REVENUE}")
        fig_rev_cost.add_bar(name="Полные затраты", x=blocks["Название блока"], y=blocks["Полные затраты"], marker_color=f"#{COLOR_COST}")
        fig_rev_cost.update_layout(
            title="Выручка vs Затраты по блокам", barmode="group",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0), margin=dict(t=60, b=40),
        )
        st.plotly_chart(fig_rev_cost, use_container_width=True)

    with chart_col2:
        margin_colors = [f"#{COLOR_POSITIVE}" if v >= 0 else f"#{COLOR_NEGATIVE}" for v in blocks["Рентабельность"]]
        fig_margin = go.Figure(
            go.Bar(
                x=blocks["Название блока"], y=blocks["Рентабельность"] * 100, marker_color=margin_colors,
                text=[f"{v * 100:.1f}%" for v in blocks["Рентабельность"]], textposition="outside",
            )
        )
        fig_margin.update_layout(title="Валовая рентабельность по блокам, %", yaxis_title="%", margin=dict(t=60, b=40))
        fig_margin.add_hline(y=0, line_color="#898781", line_width=1)
        st.plotly_chart(fig_margin, use_container_width=True)

    chart_col3, chart_col4 = st.columns(2)
    with chart_col3:
        fig_pie = go.Figure(
            go.Pie(
                labels=list(revenue_components.keys()), values=list(revenue_components.values()),
                marker=dict(colors=[f"#{c}" for c in PIE_COLORS]), hole=0.35,
            )
        )
        fig_pie.update_layout(title="Структура выручки проекта (Итого)", margin=dict(t=60, b=20))
        st.plotly_chart(fig_pie, use_container_width=True)

    with chart_col4:
        fig_cost_structure = go.Figure()
        fig_cost_structure.add_bar(name="Прямые затраты", x=blocks["Название блока"], y=blocks["Прямые затраты"], marker_color=f"#{COLOR_DIRECT_COST}")
        fig_cost_structure.add_bar(name="Аллоцированные затраты", x=blocks["Название блока"], y=blocks["Аллоцированные затраты"], marker_color=f"#{COLOR_ALLOC_COST}")
        fig_cost_structure.update_layout(
            title="Структура затрат по блокам", barmode="stack",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0), margin=dict(t=60, b=40),
        )
        st.plotly_chart(fig_cost_structure, use_container_width=True)

# ======================================================================
# 10. ЭКСПОРТ В EXCEL (openpyxl) — ЖИВЫЕ ФОРМУЛЫ, 2 ЛИСТА, ВСТРОЕННЫЕ ГРАФИКИ
# ======================================================================
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=13)
PARAM_FONT = Font(bold=True)
TOTAL_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
TOTAL_FONT = Font(bold=True)
MONEY_FMT = "#,##0"
PERCENT_FMT = "0.0%"

SHEET1_NAME = "Экономика проекта"
SHEET2_NAME = "СМР по методике"

EXCEL_INPUT_HEADERS = [
    "Название блока", "Тип блока",
    "S квартир, м2", "S коммерции 1 эт., м2", "S кладовых, м2",
    "Подземный паркинг, м/м", "Наземный/Многоуровневый паркинг, м/м",
    "Цена жилья, руб/м2", "Цена коммерции, руб/м2", "Цена кладовых, руб/м2",
    "Цена подземного м/м, руб", "Цена наземного м/м, руб",
    "Ставка СМР подземного м/м, руб", "Ставка СМР наземного м/м, руб",
] + PARAM_COLS
EXCEL_CALC_HEADERS = [
    "NSA, м2", "Доля аллокации", "Аллоцированные затраты",
    "Себестоимость коробки (методика)", "Прямые затраты",
    "Полные затраты", "Выручка", "Валовая прибыль", "Рентабельность",
]
EXCEL_HEADERS = EXCEL_INPUT_HEADERS + EXCEL_CALC_HEADERS
N_COLS_EXCEL = len(EXCEL_HEADERS)
COL = {name: get_column_letter(i + 1) for i, name in enumerate(EXCEL_HEADERS)}


def style_header_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_cell(cell, number_format=None, bold=False, fill=None):
    cell.border = BORDER
    if number_format:
        cell.number_format = number_format
    if bold:
        cell.font = TOTAL_FONT
    if fill:
        cell.fill = fill


def autosize(ws, n_cols, width=15):
    for c in range(1, n_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def build_excel_report() -> bytes:
    """Собирает Excel-отчет с живыми формулами и встроенными графиками.
    Лист 1 ссылается на Лист 2 для себестоимости коробки жилых блоков —
    изменение ставок в каталоге (Лист 2) пересчитывает весь отчет."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = SHEET1_NAME

    ws1["A1"] = f"Проект: {project_name}"
    ws1["A1"].font = TITLE_FONT
    ws1["A2"] = f"Город: {project_city}"
    ws1["A2"].font = TITLE_FONT
    ws1["A3"] = f"Сценарий: {scenario_name}"
    ws1["A3"].font = TITLE_FONT
    ws1["A4"] = "Коэфф. выручки (сценарий):"
    ws1["B4"] = rev_factor
    ws1["A4"].font = PARAM_FONT
    ws1["D4"] = "Коэфф. затрат (сценарий):"
    ws1["E4"] = cost_factor
    ws1["D4"].font = PARAM_FONT
    ws1["G4"] = "Пул косвенных расходов, руб (база):"
    ws1["G4"].font = PARAM_FONT
    # H4 заполняется формулой ниже, после того как на Листе 2 посчитаны
    # ИТОГО G и ИТОГО Z (пул = сайдбар-статьи + G + Z, живая ссылка на Лист 2).

    header_row1 = 6
    first_row1 = header_row1 + 1
    last_row1 = first_row1 + N_ROWS - 1

    for j, h in enumerate(EXCEL_HEADERS, start=1):
        ws1.cell(row=header_row1, column=j, value=h)
    style_header_row(ws1, header_row1, N_COLS_EXC
