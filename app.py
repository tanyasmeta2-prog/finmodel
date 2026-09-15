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


def get_basis_values(basis_key: str, df: pd.DataFrame, nsa: np.ndarray) -> np.ndarray:
    """Возвращает массив значений базы расчета (площадь/объем/шт) по каждой строке блоков."""
    if basis_key == "NSA":
        return nsa
    if basis_key == "VOL_TOTAL":
        return (df["Объем здания ниже 0, м3"] + df["Объем здания выше 0, м3"]).to_numpy(dtype=float)
    if basis_key == "STORAGE_AREA":
        return df["S кладовых, м2"].to_numpy(dtype=float)
    return df[BASIS_COLUMN[basis_key]].to_numpy(dtype=float)


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
    st.caption("Земля/сети/благоустройство/социалка — распределяется ТОЛЬКО на «Жилые блоки» пропорц. NSA")
    cost_land = st.number_input("Земля", min_value=0.0, value=500_000_000.0, step=1_000_000.0)
    cost_infra = st.number_input("Сети / Инфраструктура", min_value=0.0, value=300_000_000.0, step=1_000_000.0)
    cost_landscape = st.number_input("Благоустройство / Дороги", min_value=0.0, value=150_000_000.0, step=1_000_000.0)
    cost_social = st.number_input("Социальные объекты (школы/сады)", min_value=0.0, value=400_000_000.0, step=1_000_000.0)
    cost_soft = st.number_input("Прочие Soft Costs", min_value=0.0, value=100_000_000.0, step=1_000_000.0)
    indirect_pool_total = cost_land + cost_infra + cost_landscape + cost_social + cost_soft
    st.caption(f"Итого пул косвенных расходов: {indirect_pool_total:,.0f} руб".replace(",", " "))

# ======================================================================
# 6. ЕДИНАЯ ТАБЛИЦА ТЭП + КАТАЛОГ РАСЦЕНОК (session_state)
# ======================================================================
if "blocks_df" not in st.session_state:
    st.session_state.blocks_df = generate_default_table()
if "rates_df" not in st.session_state:
    st.session_state.rates_df = DEFAULT_RATES_DF.copy()

tab_main, tab_smr = st.tabs(["📊 Финансовая модель", "🏗️ СМР по методике"])

# ------------------------------------------------------------------
# ВКЛАДКА 1: основная таблица ТЭП
# ------------------------------------------------------------------
with tab_main:
    st.subheader("ТЭП проекта (жилые блоки и блоки-паркинги в одной таблице)")
    st.caption(
        "Себестоимость коробки жилого блока считается на вкладке «СМР по методике» — "
        "здесь задаются площади, цены и доп. параметры для этого расчета."
    )

    column_config = {
        "Название блока": st.column_config.TextColumn(required=True),
        "Тип блока": st.column_config.SelectboxColumn(options=BLOCK_TYPES, required=True),
    }
    for col in NUMERIC_COLS:
        column_config[col] = st.column_config.NumberColumn(min_value=0, format="%.0f")

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

# NSA нужна и для аллокации, и как база нескольких статей методики
nsa = np.where(is_res, blocks["S квартир, м2"] + blocks["S коммерции 1 эт., м2"] + blocks["S кладовых, м2"], 0.0)
blocks["NSA, м2"] = nsa

# ------------------------------------------------------------------
# ВКЛАДКА 2: каталог расценок + расчет себестоимости коробки по методике
# ------------------------------------------------------------------
with tab_smr:
    st.subheader("Каталог расценок по видам работ (действующая методика)")
    st.caption(
        "Ставки задаются один раз на проект (по объекту-аналогу) и применяются ко всем жилым блокам. "
        "Код и база расчета фиксированы методикой — редактируется только ставка."
    )
    rates_edited = st.data_editor(
        st.session_state.rates_df,
        use_container_width=True,
        num_rows="fixed",
        key="rates_editor",
        column_config={
            "Код": st.column_config.TextColumn(disabled=True),
            "Статья затрат": st.column_config.TextColumn(disabled=True),
            "Группа": st.column_config.TextColumn(disabled=True),
            "Единица измерения": st.column_config.TextColumn(disabled=True),
            "_basis": None,  # служебная колонка — скрыта
            "Ставка, руб/ед.": st.column_config.NumberColumn(min_value=0, format="%.0f"),
        },
    )
    st.session_state.rates_df = rates_edited

rates_df = st.session_state.rates_df.copy()
rates_df["Ставка, руб/ед."] = pd.to_numeric(rates_df["Ставка, руб/ед."], errors="coerce").fillna(0.0)

# -- Расчет себестоимости коробки по методике: матрица (блок x статья) --
item_cost_matrix = {}  # код статьи -> np.array стоимости по блокам (только жилые, иначе 0)
for _, item_row in rates_df.iterrows():
    code = item_row["Код"]
    basis_key = item_row["_basis"]
    rate = item_row["Ставка, руб/ед."]
    qty = get_basis_values(basis_key, blocks, nsa)
    cost = np.where(is_res, qty * rate, 0.0)
    item_cost_matrix[code] = cost

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
        for _, item_row in rates_df.iterrows():
            g = GROUP_LABELS[item_row["Группа"]]
            group_totals[g] = group_totals.get(g, 0.0) + item_cost_matrix[item_row["Код"]].sum()
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
        detail_df.insert(0, "Статья затрат", rates_df.set_index("Код")["Статья затрат"].reindex(detail_df.index).values)
        st.dataframe(detail_df, use_container_width=True)

# ======================================================================
# 7. РАСЧЕТ ЭКОНОМИКИ (единая логика на всю таблицу, ветвление по типу)
# ======================================================================
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
    ws1["H4"] = indirect_pool_total
    ws1["H4"].number_format = MONEY_FMT
    ws1["G4"].font = PARAM_FONT

    header_row1 = 6
    first_row1 = header_row1 + 1
    last_row1 = first_row1 + N_ROWS - 1

    for j, h in enumerate(EXCEL_HEADERS, start=1):
        ws1.cell(row=header_row1, column=j, value=h)
    style_header_row(ws1, header_row1, N_COLS_EXCEL)

    # ------------------------------------------------------------------
    # Лист 2: каталог расценок + расчет себестоимости коробки по блокам
    # (строится ПЕРЕД заполнением листа 1, т.к. лист 1 на него ссылается)
    # ------------------------------------------------------------------
    ws2 = wb.create_sheet(SHEET2_NAME)
    ws2["A1"] = "Расчет СМР по укрупненным видам работ (действующая методика)"
    ws2["A1"].font = TITLE_FONT
    ws2["A2"] = f"Проект: {project_name} | Сценарий: {scenario_name}"

    ws2["A4"] = "Каталог расценок (редактируется только ставка)"
    ws2["A4"].font = PARAM_FONT
    rates_header_row = 5
    for j, h in enumerate(["Код", "Статья затрат", "Группа", "Единица измерения", "Ставка, руб/ед."], start=1):
        ws2.cell(row=rates_header_row, column=j, value=h)
    style_header_row(ws2, rates_header_row, 5)

    rates_first_row = rates_header_row + 1
    rate_row_by_code = {}
    for i, item_row in rates_df.iterrows():
        r = rates_first_row + i
        rate_row_by_code[item_row["Код"]] = r
        vals = [item_row["Код"], item_row["Статья затрат"], item_row["Группа"], item_row["Единица измерения"], item_row["Ставка, руб/ед."]]
        for j, v in enumerate(vals, start=1):
            cell = ws2.cell(row=r, column=j, value=v)
            style_cell(cell, number_format=(MONEY_FMT if j == 5 else None))
    rates_last_row = rates_first_row + len(rates_df) - 1

    # -- Расчет по блокам: строки = блоки (в том же порядке, что на листе 1) --
    calc_title_row = rates_last_row + 3
    ws2.cell(row=calc_title_row, column=1, value="Себестоимость коробки по блокам (только «Жилой блок»)")
    ws2.cell(row=calc_title_row, column=1).font = PARAM_FONT
    calc_header_row = calc_title_row + 1
    item_codes = list(rates_df["Код"])
    calc_headers2 = ["Название блока", "Тип блока"] + item_codes + ["ИТОГО СМР коробки, руб"]
    for j, h in enumerate(calc_headers2, start=1):
        ws2.cell(row=calc_header_row, column=j, value=h)
    style_header_row(ws2, calc_header_row, len(calc_headers2))

    calc_first_row = calc_header_row + 1
    total_col_idx2 = len(calc_headers2)  # последняя колонка — ИТОГО
    for i, (_, row) in enumerate(blocks.iterrows()):
        r2 = calc_first_row + i
        r1 = first_row1 + i  # соответствующая строка на листе 1 (тот же порядок блоков)
        ws2.cell(row=r2, column=1, value=row["Название блока"])
        ws2.cell(row=r2, column=2, value=row["Тип блока"])
        style_cell(ws2.cell(row=r2, column=1))
        style_cell(ws2.cell(row=r2, column=2))

        type_ref = f"'{SHEET1_NAME}'!{COL['Тип блока']}{r1}"
        for k, item_row in rates_df.iterrows():
            code = item_row["Код"]
            basis_key = item_row["_basis"]
            rr = rate_row_by_code[code]
            rate_ref = f"$E${rr}"
            if basis_key == "NSA":
                qty_ref = f"'{SHEET1_NAME}'!{COL['NSA, м2']}{r1}"
            elif basis_key == "VOL_TOTAL":
                qty_ref = f"('{SHEET1_NAME}'!{COL['Объем здания ниже 0, м3']}{r1}+'{SHEET1_NAME}'!{COL['Объем здания выше 0, м3']}{r1})"
            elif basis_key == "STORAGE_AREA":
                qty_ref = f"'{SHEET1_NAME}'!{COL['S кладовых, м2']}{r1}"
            else:
                basis_col_name = BASIS_COLUMN[basis_key]
                qty_ref = f"'{SHEET1_NAME}'!{COL[basis_col_name]}{r1}"
            formula = f'=IF({type_ref}="{TYPE_RESIDENTIAL}",{qty_ref}*{rate_ref},0)'
            cell = ws2.cell(row=r2, column=3 + k, value=formula)
            style_cell(cell, number_format=MONEY_FMT)

        first_item_col = get_column_letter(3)
        last_item_col = get_column_letter(2 + len(item_codes))
        total_cell = ws2.cell(row=r2, column=total_col_idx2, value=f"=SUM({first_item_col}{r2}:{last_item_col}{r2})")
        style_cell(total_cell, number_format=MONEY_FMT, bold=True, fill=TOTAL_FILL)
    calc_last_row = calc_first_row + N_ROWS - 1

    autosize(ws2, len(calc_headers2), width=13)
    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 24

    # -- Группы работ (для круговой диаграммы) — суммы по прямоугольным
    #    диапазонам колонок статей внутри каждой группы (колонки статей
    #    идут подряд, сгруппированы по буквенному коду) --
    group_order, group_col_ranges = [], {}
    start_idx = None
    prev_group = None
    for k, item_row in rates_df.iterrows():
        g = item_row["Группа"]
        col_idx = 3 + k
        if g != prev_group:
            if prev_group is not None:
                group_col_ranges[prev_group] = (group_col_ranges[prev_group][0], col_idx - 1)
            group_col_ranges[g] = (col_idx, col_idx)
            group_order.append(g)
            prev_group = g
        else:
            group_col_ranges[g] = (group_col_ranges[g][0], col_idx)

    group_table_row0 = calc_last_row + 3
    ws2.cell(row=group_table_row0, column=1, value="Структура СМР коробки по группам работ")
    ws2.cell(row=group_table_row0, column=1).font = PARAM_FONT
    group_header_row = group_table_row0 + 1
    ws2.cell(row=group_header_row, column=1, value="Группа")
    ws2.cell(row=group_header_row, column=2, value="Сумма, руб")
    style_header_row(ws2, group_header_row, 2)
    for gi, g in enumerate(group_order):
        r = group_header_row + 1 + gi
        c1, c2 = group_col_ranges[g]
        col1_letter, col2_letter = get_column_letter(c1), get_column_letter(c2)
        label_cell = ws2.cell(row=r, column=1, value=GROUP_LABELS[g])
        sum_cell = ws2.cell(
            row=r, column=2,
            value=f"=SUM({col1_letter}{calc_first_row}:{col2_letter}{calc_last_row})" if N_ROWS > 0 else 0,
        )
        style_cell(label_cell)
        style_cell(sum_cell, number_format=MONEY_FMT)
    group_last_row = group_header_row + len(group_order)

    # -- Встроенные графики Листа 2 --
    if N_ROWS > 0:
        bar2 = BarChart()
        bar2.type = "col"
        bar2.title = "Себестоимость коробки по блокам (методика)"
        bar2.y_axis.title = "руб"
        bar2.style = 10
        data_ref2 = Reference(ws2, min_col=total_col_idx2, max_col=total_col_idx2, min_row=calc_header_row, max_row=calc_last_row)
        cats_ref2 = Reference(ws2, min_col=1, min_row=calc_first_row, max_row=calc_last_row)
        bar2.add_data(data_ref2, titles_from_data=True)
        bar2.set_categories(cats_ref2)
        bar2.series[0].graphicalProperties.solidFill = COLOR_DIRECT_COST
        bar2.height, bar2.width = 10, 22
        ws2.add_chart(bar2, f"{get_column_letter(total_col_idx2 + 2)}{calc_header_row}")

        pie2 = PieChart()
        pie2.title = "Структура СМР коробки по группам работ"
        data_ref_g = Reference(ws2, min_col=2, min_row=group_header_row, max_row=group_last_row)
        cats_ref_g = Reference(ws2, min_col=1, min_row=group_header_row + 1, max_row=group_last_row)
        pie2.add_data(data_ref_g, titles_from_data=True)
        pie2.set_categories(cats_ref_g)
        pie2.dataLabels = DataLabelList()
        pie2.dataLabels.showPercent = True
        pie2.series[0].data_points = [
            DataPoint(idx=i, spPr=GraphicalProperties(solidFill=GROUP_COLORS[i % len(GROUP_COLORS)]))
            for i in range(len(group_order))
        ]
        pie2.height, pie2.width = 10, 14
        ws2.add_chart(pie2, f"{get_column_letter(total_col_idx2 + 2)}{calc_header_row + 22}")

    # ------------------------------------------------------------------
    # Возвращаемся к листу 1: входные данные + формулы (ссылаются на лист 2)
    # ------------------------------------------------------------------
    for i, (_, row) in enumerate(blocks.iterrows()):
        r1 = first_row1 + i
        input_values = [row[h] for h in EXCEL_INPUT_HEADERS]
        for j, v in enumerate(input_values, start=1):
            cell = ws1.cell(row=r1, column=j, value=v)
            style_cell(cell, number_format=(MONEY_FMT if j >= 3 else None))

        r2 = calc_first_row + i
        c = COL
        f_nsa = f'=IF({c["Тип блока"]}{r1}="{TYPE_RESIDENTIAL}",{c["S квартир, м2"]}{r1}+{c["S коммерции 1 эт., м2"]}{r1}+{c["S кладовых, м2"]}{r1},0)'
        f_share = f'=IF(SUM(${c["NSA, м2"]}${first_row1}:${c["NSA, м2"]}${last_row1})=0,0,{c["NSA, м2"]}{r1}/SUM(${c["NSA, м2"]}${first_row1}:${c["NSA, м2"]}${last_row1}))'
        f_alloc = f'={c["Доля аллокации"]}{r1}*$H$4*$E$4'
        f_korobka = f"='{SHEET2_NAME}'!{get_column_letter(total_col_idx2)}{r2}"
        f_direct = (
            f'=IF({c["Тип блока"]}{r1}="{TYPE_RESIDENTIAL}",'
            f'({c["Себестоимость коробки (методика)"]}{r1}+{c["Подземный паркинг, м/м"]}{r1}*{c["Ставка СМР подземного м/м, руб"]}{r1})*$E$4,'
            f'{c["Наземный/Многоуровневый паркинг, м/м"]}{r1}*{c["Ставка СМР наземного м/м, руб"]}{r1}*$E$4)'
        )
        f_full = f'={c["Аллоцированные затраты"]}{r1}+{c["Прямые затраты"]}{r1}'
        f_revenue = (
            f'=IF({c["Тип блока"]}{r1}="{TYPE_RESIDENTIAL}",'
            f'({c["S квартир, м2"]}{r1}*{c["Цена жилья, руб/м2"]}{r1}'
            f'+{c["S коммерции 1 эт., м2"]}{r1}*{c["Цена коммерции, руб/м2"]}{r1}'
            f'+{c["S кладовых, м2"]}{r1}*{c["Цена кладовых, руб/м2"]}{r1}'
            f'+{c["Подземный паркинг, м/м"]}{r1}*{c["Цена подземного м/м, руб"]}{r1})*$B$4,'
            f'{c["Наземный/Многоуровневый паркинг, м/м"]}{r1}*{c["Цена наземного м/м, руб"]}{r1}*$B$4)'
        )
        f_profit = f'={c["Выручка"]}{r1}-{c["Полные затраты"]}{r1}'
        f_margin = f'=IF({c["Выручка"]}{r1}=0,0,{c["Валовая прибыль"]}{r1}/{c["Выручка"]}{r1})'

        calc_formulas = [f_nsa, f_share, f_alloc, f_korobka, f_direct, f_full, f_revenue, f_profit, f_margin]
        for k, formula in enumerate(calc_formulas):
            col_idx = len(EXCEL_INPUT_HEADERS) + 1 + k
            cell = ws1.cell(row=r1, column=col_idx, value=formula)
            header_name = EXCEL_CALC_HEADERS[k]
            fmt = PERCENT_FMT if header_name in ("Доля аллокации", "Рентабельность") else MONEY_FMT
            style_cell(cell, number_format=fmt)

    total_row1 = last_row1 + 1 if N_ROWS > 0 else first_row1
    ws1.cell(row=total_row1, column=1, value="ИТОГО")
    if N_ROWS > 0:
        for header_name in ["NSA, м2", "Доля аллокации", "Аллоцированные затраты", "Себестоимость коробки (методика)",
                             "Прямые затраты", "Полные затраты", "Выручка", "Валовая прибыль"]:
            col_letter = COL[header_name]
            ws1[f"{col_letter}{total_row1}"] = f"=SUM({col_letter}{first_row1}:{col_letter}{last_row1})"
        ws1[f'{COL["Рентабельность"]}{total_row1}'] = (
            f'=IF({COL["Выручка"]}{total_row1}=0,0,{COL["Валовая прибыль"]}{total_row1}/{COL["Выручка"]}{total_row1})'
        )
    else:
        for header_name in EXCEL_CALC_HEADERS:
            ws1[f"{COL[header_name]}{total_row1}"] = 0
    for j in range(1, N_COLS_EXCEL + 1):
        cell = ws1.cell(row=total_row1, column=j)
        header_name = EXCEL_HEADERS[j - 1]
        fmt = PERCENT_FMT if header_name in ("Доля аллокации", "Рентабельность") else (MONEY_FMT if j >= 3 else None)
        style_cell(cell, number_format=fmt, bold=True, fill=TOTAL_FILL)

    autosize(ws1, N_COLS_EXCEL, width=14)
    ws1.column_dimensions["A"].width = 16
    ws1.column_dimensions["B"].width = 26

    if N_ROWS > 0:
        bar = BarChart()
        bar.type = "col"
        bar.title = "Выручка vs Полные затраты по блокам"
        bar.y_axis.title = "руб"
        bar.style = 10
        full_cost_idx = EXCEL_HEADERS.index("Полные затраты") + 1
        revenue_idx = EXCEL_HEADERS.index("Выручка") + 1
        data_ref = Reference(ws1, min_col=full_cost_idx, max_col=revenue_idx, min_row=header_row1, max_row=last_row1)
        cats_ref = Reference(ws1, min_col=1, min_row=first_row1, max_row=last_row1)
        bar.add_data(data_ref, titles_from_data=True)
        bar.set_categories(cats_ref)
        bar.series[0].graphicalProperties.solidFill = COLOR_COST
        bar.series[1].graphicalProperties.solidFill = COLOR_REVENUE
        bar.height, bar.width = 10, 24
        ws1.add_chart(bar, f"{get_column_letter(N_COLS_EXCEL + 2)}{header_row1}")

    struct_row0 = total_row1 + 3
    ws1.cell(row=struct_row0, column=1, value="Структура выручки проекта")
    ws1.cell(row=struct_row0, column=1).font = TITLE_FONT
    struct_header_row = struct_row0 + 1
    ws1.cell(row=struct_header_row, column=1, value="Статья выручки")
    ws1.cell(row=struct_header_row, column=2, value="Сумма, руб")
    style_header_row(ws1, struct_header_row, 2)

    if N_ROWS > 0:
        rng_type = f'${COL["Тип блока"]}${first_row1}:${COL["Тип блока"]}${last_row1}'
        rng_apt = f'${COL["S квартир, м2"]}${first_row1}:${COL["S квартир, м2"]}${last_row1}'
        rng_p_apt = f'${COL["Цена жилья, руб/м2"]}${first_row1}:${COL["Цена жилья, руб/м2"]}${last_row1}'
        rng_c1 = f'${COL["S коммерции 1 эт., м2"]}${first_row1}:${COL["S коммерции 1 эт., м2"]}${last_row1}'
        rng_p_c1 = f'${COL["Цена коммерции, руб/м2"]}${first_row1}:${COL["Цена коммерции, руб/м2"]}${last_row1}'
        rng_storage = f'${COL["S кладовых, м2"]}${first_row1}:${COL["S кладовых, м2"]}${last_row1}'
        rng_p_storage = f'${COL["Цена кладовых, руб/м2"]}${first_row1}:${COL["Цена кладовых, руб/м2"]}${last_row1}'
        rng_underground = f'${COL["Подземный паркинг, м/м"]}${first_row1}:${COL["Подземный паркинг, м/м"]}${last_row1}'
        rng_p_underground = f'${COL["Цена подземного м/м, руб"]}${first_row1}:${COL["Цена подземного м/м, руб"]}${last_row1}'
        rng_ground = f'${COL["Наземный/Многоуровневый паркинг, м/м"]}${first_row1}:${COL["Наземный/Многоуровневый паркинг, м/м"]}${last_row1}'
        rng_p_ground = f'${COL["Цена наземного м/м, руб"]}${first_row1}:${COL["Цена наземного м/м, руб"]}${last_row1}'
        struct_formulas = [
            ("Жилье", f'=SUMPRODUCT(({rng_type}="{TYPE_RESIDENTIAL}")*{rng_apt}*{rng_p_apt})*$B$4'),
            ("Коммерция", f'=SUMPRODUCT(({rng_type}="{TYPE_RESIDENTIAL}")*{rng_c1}*{rng_p_c1})*$B$4'),
            ("Кладовые", f'=SUMPRODUCT(({rng_type}="{TYPE_RESIDENTIAL}")*{rng_storage}*{rng_p_storage})*$B$4'),
            ("Подземные м/м", f'=SUMPRODUCT(({rng_type}="{TYPE_RESIDENTIAL}")*{rng_underground}*{rng_p_underground})*$B$4'),
            ("Наземные/Многоур. паркинги", f'=SUMPRODUCT(({rng_type}="{TYPE_PARKING}")*{rng_ground}*{rng_p_ground})*$B$4'),
        ]
    else:
        struct_formulas = [(label, 0) for label in revenue_components.keys()]

    for i, (label, formula) in enumerate(struct_formulas):
        r = struct_header_row + 1 + i
        c1 = ws1.cell(row=r, column=1, value=label)
        c2 = ws1.cell(row=r, column=2, value=formula)
        style_cell(c1)
        style_cell(c2, number_format=MONEY_FMT)
    struct_last_row = struct_header_row + len(struct_formulas)

    pie = PieChart()
    pie.title = "Структура выручки проекта"
    data_ref = Reference(ws1, min_col=2, min_row=struct_header_row, max_row=struct_last_row)
    cats_ref = Reference(ws1, min_col=1, min_row=struct_header_row + 1, max_row=struct_last_row)
    pie.add_data(data_ref, titles_from_data=True)
    pie.set_categories(cats_ref)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.series[0].data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=PIE_COLORS[i % len(PIE_COLORS)]))
        for i in range(len(struct_formulas))
    ]
    pie.height, pie.width = 10, 16
    ws1.add_chart(pie, f"{get_column_letter(N_COLS_EXCEL + 2)}{struct_header_row}")

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


excel_bytes = build_excel_report()
st.divider()
st.download_button(
    label="📥 Скачать отчет в Excel (2 листа, живые формулы + графики)",
    data=excel_bytes,
    file_name=f"financial_model_{scenario_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
