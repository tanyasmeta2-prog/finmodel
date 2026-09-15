# -*- coding: utf-8 -*-
"""
Финансовая модель девелоперского проекта (Streamlit + openpyxl)
=================================================================
Универсальный расчет экономики проекта (от 1 до N блоков) на уровне
Валовой прибыли и Валовой рентабельности (без налогов и кредитов).

Ключевое допущение этой версии — ЕДИНАЯ таблица ТЭП, где каждая строка
(урбан-блок) имеет "Тип блока":
  • "Жилой блок" — жилье, коммерция 1 эт., кладовые, подземный паркинг.
    Именно на эти строки распределяется пул косвенных расходов площадки,
    пропорционально их суммарной продаваемой площади (NSA).
  • "Наземный/Многоуровневый паркинг" — отдельно стоящий паркинг. Полностью
    освобожден от аллокации косвенных расходов (доля аллокации = 0),
    считается только по своим прямым затратам и выручке.

Лоты коммерции в стилобате из модели удалены.

Экспорт в Excel строится ОДНОЙ таблицей "Экономика проекта" с живыми
формулами Excel (SUM, IF, SUMPRODUCT) — при ручном изменении входных ячеек
прямо в Excel итоги и графики пересчитаются сами, без Python.

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
    "blue": "2a78d6",
    "orange": "eb6834",
    "aqua": "1baf7a",
    "yellow": "eda100",
    "magenta": "e87ba4",
    "violet": "4a3aa7",
    "red": "e34948",
}
COLOR_REVENUE = PALETTE["blue"]
COLOR_COST = PALETTE["orange"]
COLOR_POSITIVE = PALETTE["blue"]
COLOR_NEGATIVE = PALETTE["red"]
COLOR_DIRECT_COST = PALETTE["blue"]
COLOR_ALLOC_COST = PALETTE["violet"]
PIE_COLORS = [PALETTE["blue"], PALETTE["orange"], PALETTE["aqua"], PALETTE["yellow"], PALETTE["magenta"]]

# ======================================================================
# 1. НАСТРОЙКИ СТРАНИЦЫ И СЦЕНАРИИ
# ======================================================================
st.set_page_config(page_title="Финмодель девелоперского проекта", layout="wide")

TYPE_RESIDENTIAL = "Жилой блок"
TYPE_PARKING = "Наземный/Многоуровневый паркинг"
BLOCK_TYPES = [TYPE_RESIDENTIAL, TYPE_PARKING]

# Коэффициенты применяются ко ВСЕЙ выручке и ко ВСЕМ затратам (прямым и
# аллоцированным) — единообразно для всех строк таблицы.
SCENARIOS = {
    "Базовый": {"revenue": 1.00, "cost": 1.00},
    "Стресс": {"revenue": 0.85, "cost": 1.15},         # выручка -15%, затраты +15%
    "Оптимистичный": {"revenue": 1.10, "cost": 0.95},  # выручка +10%, затраты -5%
}

# ======================================================================
# 2. КОЛОНКИ ЕДИНОЙ ТАБЛИЦЫ И ТЕСТОВЫЕ ДАННЫЕ
# ======================================================================
NUMERIC_COLS = [
    "S квартир, м2",
    "S коммерции 1 эт., м2",
    "S кладовых, м2",
    "Подземный паркинг, м/м",
    "Наземный/Многоуровневый паркинг, м/м",
    "Цена жилья, руб/м2",
    "Цена коммерции, руб/м2",
    "Цена кладовых, руб/м2",
    "Цена подземного м/м, руб",
    "Цена наземного м/м, руб",
    "Ставка СМР площадей, руб/м2",
    "Ставка СМР подземного м/м, руб",
    "Ставка СМР наземного м/м, руб",
]
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
        rows.append(
            {
                "Название блока": f"УБ {i}",
                "Тип блока": TYPE_RESIDENTIAL,
                "S квартир, м2": s_apt,
                "S коммерции 1 эт., м2": s_c1,
                "S кладовых, м2": s_storage,
                "Подземный паркинг, м/м": parking_u,
                "Наземный/Многоуровневый паркинг, м/м": 0,
                "Цена жилья, руб/м2": int(rng.integers(150, 221) * 1000),
                "Цена коммерции, руб/м2": int(rng.integers(200, 321) * 1000),
                "Цена кладовых, руб/м2": int(rng.integers(40, 91) * 1000),
                "Цена подземного м/м, руб": int(rng.integers(700, 1001) * 1000),
                "Цена наземного м/м, руб": 0,
                "Ставка СМР площадей, руб/м2": int(rng.integers(50, 69) * 1000),
                "Ставка СМР подземного м/м, руб": int(rng.integers(650, 821) * 1000),
                "Ставка СМР наземного м/м, руб": 0,
            }
        )

    # Два блока-паркинга по ТЗ — фиксированное количество машиномест
    for name, count in [("УБ-Паркинг 1", 1494), ("УБ-Паркинг 2", 1437)]:
        rows.append(
            {
                "Название блока": name,
                "Тип блока": TYPE_PARKING,
                "S квартир, м2": 0,
                "S коммерции 1 эт., м2": 0,
                "S кладовых, м2": 0,
                "Подземный паркинг, м/м": 0,
                "Наземный/Многоуровневый паркинг, м/м": count,
                "Цена жилья, руб/м2": 0,
                "Цена коммерции, руб/м2": 0,
                "Цена кладовых, руб/м2": 0,
                "Цена подземного м/м, руб": 0,
                "Цена наземного м/м, руб": 450000,
                "Ставка СМР площадей, руб/м2": 0,
                "Ставка СМР подземного м/м, руб": 0,
                "Ставка СМР наземного м/м, руб": 320000,
            }
        )

    return pd.DataFrame(rows)[ALL_COLS]


# ======================================================================
# 3. ПАСПОРТ ПРОЕКТА
# ======================================================================
st.title("Финансовая модель девелоперского проекта")

pass_col1, pass_col2 = st.columns(2)
with pass_col1:
    project_name = st.text_input("Наименование проекта", value="ЖК «Пример»")
with pass_col2:
    project_city = st.text_input("Город", value="Самара")

st.caption("Валовая прибыль и валовая рентабельность. Налоги и кредиты не учитываются.")

# ======================================================================
# 4. БОКОВАЯ ПАНЕЛЬ — СЦЕНАРИЙ И ПУЛ КОСВЕННЫХ РАСХОДОВ
# ======================================================================
with st.sidebar:
    st.header("Сценарий расчета")
    scenario_name = st.selectbox("Выберите сценарий", list(SCENARIOS.keys()))
    rev_factor = SCENARIOS[scenario_name]["revenue"]
    cost_factor = SCENARIOS[scenario_name]["cost"]

    st.header("Пул косвенных расходов проекта, руб")
    st.caption("Распределяется ТОЛЬКО на строки «Жилой блок», пропорционально их NSA")
    cost_land = st.number_input("Земля", min_value=0.0, value=500_000_000.0, step=1_000_000.0)
    cost_infra = st.number_input("Сети / Инфраструктура", min_value=0.0, value=300_000_000.0, step=1_000_000.0)
    cost_landscape = st.number_input("Благоустройство / Дороги", min_value=0.0, value=150_000_000.0, step=1_000_000.0)
    cost_social = st.number_input("Социальные объекты (школы/сады)", min_value=0.0, value=400_000_000.0, step=1_000_000.0)
    cost_soft = st.number_input("Прочие Soft Costs", min_value=0.0, value=100_000_000.0, step=1_000_000.0)
    indirect_pool_total = cost_land + cost_infra + cost_landscape + cost_social + cost_soft
    st.caption(f"Итого пул косвенных расходов: {indirect_pool_total:,.0f} руб".replace(",", " "))

# ======================================================================
# 5. ЕДИНАЯ ТАБЛИЦА ТЭП (st.data_editor)
# ======================================================================
if "blocks_df" not in st.session_state:
    st.session_state.blocks_df = generate_default_table()

st.subheader("ТЭП проекта (жилые блоки и блоки-паркинги в одной таблице)")
st.caption(
    "Тип блока определяет формулу расчета. Строки можно добавлять/удалять — "
    "модель работает для любого количества и сочетания блоков."
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

# Приведение типов и отбрасывание пустых строк
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

# ======================================================================
# 6. РАСЧЕТ ЭКОНОМИКИ (единая логика на всю таблицу, ветвление по типу)
# ======================================================================
# NSA (продаваемая площадь для аллокации) — только у жилых блоков
nsa = np.where(is_res, blocks["S квартир, м2"] + blocks["S коммерции 1 эт., м2"] + blocks["S кладовых, м2"], 0.0)
blocks["NSA, м2"] = nsa
total_nsa = nsa.sum()

# Доля аллокации — 0 у блоков-паркингов по определению (nsa=0 для них).
# Деление выполняется только при total_nsa > 0, чтобы не ловить 0/0.
share = np.zeros_like(nsa, dtype=float)
if total_nsa > 0:
    share = nsa / total_nsa
blocks["Доля аллокации"] = share

indirect_pool_scenario = indirect_pool_total * cost_factor
blocks["Аллоцированные затраты"] = blocks["Доля аллокации"] * indirect_pool_scenario

# Прямые затраты: разная формула для жилого блока и блока-паркинга
direct_res = (
    nsa * blocks["Ставка СМР площадей, руб/м2"]
    + blocks["Подземный паркинг, м/м"] * blocks["Ставка СМР подземного м/м, руб"]
) * cost_factor
direct_park = blocks["Наземный/Многоуровневый паркинг, м/м"] * blocks["Ставка СМР наземного м/м, руб"] * cost_factor
blocks["Прямые затраты"] = np.where(is_res, direct_res, direct_park)

blocks["Полные затраты"] = blocks["Прямые затраты"] + blocks["Аллоцированные затраты"]

# Выручка: разная формула для жилого блока и блока-паркинга
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

# Структура выручки проекта (для круговой диаграммы) — по типам лотов
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
# 7. КОНСОЛИДИРОВАННЫЕ ПОКАЗАТЕЛИ ПРОЕКТА
# ======================================================================
total_revenue = blocks["Выручка"].sum()
total_cost = blocks["Полные затраты"].sum()
total_profit = total_revenue - total_cost
avg_margin = (total_profit / total_revenue) if total_revenue > 0 else 0.0

no_residential_warning = (blocks.shape[0] > 0) and (total_nsa == 0) and (indirect_pool_total > 0)

# ======================================================================
# 8. ВЕБ-ИНТЕРФЕЙС — МЕТРИКИ И ИНТЕРАКТИВНАЯ ГРАФИКА (PLOTLY)
# ======================================================================
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
    blocks[tbl_cols],
    use_container_width=True,
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
        title="Выручка vs Затраты по блокам",
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(t=60, b=40),
    )
    st.plotly_chart(fig_rev_cost, use_container_width=True)

with chart_col2:
    margin_colors = [f"#{COLOR_POSITIVE}" if v >= 0 else f"#{COLOR_NEGATIVE}" for v in blocks["Рентабельность"]]
    fig_margin = go.Figure(
        go.Bar(
            x=blocks["Название блока"],
            y=blocks["Рентабельность"] * 100,
            marker_color=margin_colors,
            text=[f"{v * 100:.1f}%" for v in blocks["Рентабельность"]],
            textposition="outside",
        )
    )
    fig_margin.update_layout(title="Валовая рентабельность по блокам, %", yaxis_title="%", margin=dict(t=60, b=40))
    fig_margin.add_hline(y=0, line_color="#898781", line_width=1)
    st.plotly_chart(fig_margin, use_container_width=True)

chart_col3, chart_col4 = st.columns(2)

with chart_col3:
    fig_pie = go.Figure(
        go.Pie(
            labels=list(revenue_components.keys()),
            values=list(revenue_components.values()),
            marker=dict(colors=[f"#{c}" for c in PIE_COLORS]),
            hole=0.35,
        )
    )
    fig_pie.update_layout(title="Структура выручки проекта (Итого)", margin=dict(t=60, b=20))
    st.plotly_chart(fig_pie, use_container_width=True)

with chart_col4:
    fig_cost_structure = go.Figure()
    fig_cost_structure.add_bar(name="Прямые затраты", x=blocks["Название блока"], y=blocks["Прямые затраты"], marker_color=f"#{COLOR_DIRECT_COST}")
    fig_cost_structure.add_bar(name="Аллоцированные затраты", x=blocks["Название блока"], y=blocks["Аллоцированные затраты"], marker_color=f"#{COLOR_ALLOC_COST}")
    fig_cost_structure.update_layout(
        title="Структура затрат по блокам",
        barmode="stack",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(t=60, b=40),
    )
    st.plotly_chart(fig_cost_structure, use_container_width=True)

# ======================================================================
# 9. ЭКСПОРТ В EXCEL (openpyxl) — ЖИВЫЕ ФОРМУЛЫ + ВСТРОЕННЫЕ ГРАФИКИ
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

# Порядок колонок таблицы в Excel: A..O — входные данные, P..W — расчет
EXCEL_INPUT_HEADERS = [
    "Название блока", "Тип блока",
    "S квартир, м2", "S коммерции 1 эт., м2", "S кладовых, м2",
    "Подземный паркинг, м/м", "Наземный/Многоуровневый паркинг, м/м",
    "Цена жилья, руб/м2", "Цена коммерции, руб/м2", "Цена кладовых, руб/м2",
    "Цена подземного м/м, руб", "Цена наземного м/м, руб",
    "Ставка СМР площадей, руб/м2", "Ставка СМР подземного м/м, руб", "Ставка СМР наземного м/м, руб",
]
EXCEL_CALC_HEADERS = [
    "NSA, м2", "Доля аллокации", "Аллоцированные затраты", "Прямые затраты",
    "Полные затраты", "Выручка", "Валовая прибыль", "Рентабельность",
]
EXCEL_HEADERS = EXCEL_INPUT_HEADERS + EXCEL_CALC_HEADERS
N_COLS_EXCEL = len(EXCEL_HEADERS)  # 15 входных + 8 расчетных = 23 (A..W)

# Буквы ключевых столбцов (по фиксированному порядку EXCEL_HEADERS, с 1)
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
    """Собирает Excel-отчет с одной главной таблицей, живыми формулами и
    встроенными графиками (BarChart, PieChart). Все диапазоны формул и
    графиков строятся динамически под текущее число строк таблицы."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Экономика проекта"

    # -- Паспорт проекта и параметры сценария (живые ячейки-параметры) ----
    ws["A1"] = f"Проект: {project_name}"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Город: {project_city}"
    ws["A2"].font = TITLE_FONT
    ws["A3"] = f"Сценарий: {scenario_name}"
    ws["A3"].font = TITLE_FONT

    ws["A4"] = "Коэфф. выручки (сценарий):"
    ws["B4"] = rev_factor
    ws["A4"].font = PARAM_FONT
    ws["D4"] = "Коэфф. затрат (сценарий):"
    ws["E4"] = cost_factor
    ws["D4"].font = PARAM_FONT
    ws["G4"] = "Пул косвенных расходов, руб (база):"
    ws["H4"] = indirect_pool_total
    ws["H4"].number_format = MONEY_FMT
    ws["G4"].font = PARAM_FONT
    # Изменение B4 / E4 / H4 прямо в Excel пересчитает всю таблицу — формулы
    # ниже ссылаются на эти ячейки абсолютными ссылками.

    header_row = 6
    first_row = header_row + 1
    last_row = first_row + N_ROWS - 1  # если N_ROWS == 0, last_row < first_row

    for j, h in enumerate(EXCEL_HEADERS, start=1):
        ws.cell(row=header_row, column=j, value=h)
    style_header_row(ws, header_row, N_COLS_EXCEL)

    # -- Строки данных: входные значения + живые формулы расчета ----------
    for i, (_, row) in enumerate(blocks.iterrows()):
        r = first_row + i
        # входные колонки A..O
        input_values = [
            row["Название блока"], row["Тип блока"],
            row["S квартир, м2"], row["S коммерции 1 эт., м2"], row["S кладовых, м2"],
            row["Подземный паркинг, м/м"], row["Наземный/Многоуровневый паркинг, м/м"],
            row["Цена жилья, руб/м2"], row["Цена коммерции, руб/м2"], row["Цена кладовых, руб/м2"],
            row["Цена подземного м/м, руб"], row["Цена наземного м/м, руб"],
            row["Ставка СМР площадей, руб/м2"], row["Ставка СМР подземного м/м, руб"], row["Ставка СМР наземного м/м, руб"],
        ]
        for j, v in enumerate(input_values, start=1):
            cell = ws.cell(row=r, column=j, value=v)
            style_cell(cell, number_format=(MONEY_FMT if j >= 3 else None))

        # Расчетные формулы — тип блока проверяется прямо в формуле (IF)
        c = COL  # короткий алиас
        f_nsa = f'=IF({c["Тип блока"]}{r}="{TYPE_RESIDENTIAL}",{c["S квартир, м2"]}{r}+{c["S коммерции 1 эт., м2"]}{r}+{c["S кладовых, м2"]}{r},0)'
        f_share = f'=IF(SUM(${c["NSA, м2"]}${first_row}:${c["NSA, м2"]}${last_row})=0,0,{c["NSA, м2"]}{r}/SUM(${c["NSA, м2"]}${first_row}:${c["NSA, м2"]}${last_row}))'
        f_alloc = f'={c["Доля аллокации"]}{r}*$H$4*$E$4'
        f_direct = (
            f'=IF({c["Тип блока"]}{r}="{TYPE_RESIDENTIAL}",'
            f'{c["NSA, м2"]}{r}*{c["Ставка СМР площадей, руб/м2"]}{r}*$E$4'
            f'+{c["Подземный паркинг, м/м"]}{r}*{c["Ставка СМР подземного м/м, руб"]}{r}*$E$4,'
            f'{c["Наземный/Многоуровневый паркинг, м/м"]}{r}*{c["Ставка СМР наземного м/м, руб"]}{r}*$E$4)'
        )
        f_full = f'={c["Аллоцированные затраты"]}{r}+{c["Прямые затраты"]}{r}'
        f_revenue = (
            f'=IF({c["Тип блока"]}{r}="{TYPE_RESIDENTIAL}",'
            f'({c["S квартир, м2"]}{r}*{c["Цена жилья, руб/м2"]}{r}'
            f'+{c["S коммерции 1 эт., м2"]}{r}*{c["Цена коммерции, руб/м2"]}{r}'
            f'+{c["S кладовых, м2"]}{r}*{c["Цена кладовых, руб/м2"]}{r}'
            f'+{c["Подземный паркинг, м/м"]}{r}*{c["Цена подземного м/м, руб"]}{r})*$B$4,'
            f'{c["Наземный/Многоуровневый паркинг, м/м"]}{r}*{c["Цена наземного м/м, руб"]}{r}*$B$4)'
        )
        f_profit = f'={c["Выручка"]}{r}-{c["Полные затраты"]}{r}'
        f_margin = f'=IF({c["Выручка"]}{r}=0,0,{c["Валовая прибыль"]}{r}/{c["Выручка"]}{r})'

        calc_formulas = [f_nsa, f_share, f_alloc, f_direct, f_full, f_revenue, f_profit, f_margin]
        for k, formula in enumerate(calc_formulas):
            col_idx = len(EXCEL_INPUT_HEADERS) + 1 + k
            cell = ws.cell(row=r, column=col_idx, value=formula)
            header_name = EXCEL_CALC_HEADERS[k]
            fmt = PERCENT_FMT if header_name in ("Доля аллокации", "Рентабельность") else MONEY_FMT
            style_cell(cell, number_format=fmt)

    # -- Итоговая строка (живые SUM-формулы) -------------------------------
    total_row = last_row + 1 if N_ROWS > 0 else first_row
    ws.cell(row=total_row, column=1, value="ИТОГО")
    if N_ROWS > 0:
        for header_name in ["NSA, м2", "Доля аллокации", "Аллоцированные затраты", "Прямые затраты",
                             "Полные затраты", "Выручка", "Валовая прибыль"]:
            col_letter = COL[header_name]
            ws[f"{col_letter}{total_row}"] = f"=SUM({col_letter}{first_row}:{col_letter}{last_row})"
        ws[f'{COL["Рентабельность"]}{total_row}'] = (
            f'=IF({COL["Выручка"]}{total_row}=0,0,{COL["Валовая прибыль"]}{total_row}/{COL["Выручка"]}{total_row})'
        )
    else:
        for header_name in EXCEL_CALC_HEADERS:
            ws[f"{COL[header_name]}{total_row}"] = 0
    for j in range(1, N_COLS_EXCEL + 1):
        cell = ws.cell(row=total_row, column=j)
        header_name = EXCEL_HEADERS[j - 1] if j <= len(EXCEL_HEADERS) else None
        fmt = None
        if header_name in ("Доля аллокации", "Рентабельность"):
            fmt = PERCENT_FMT
        elif header_name in EXCEL_CALC_HEADERS or (header_name and j >= 3):
            fmt = MONEY_FMT
        style_cell(cell, number_format=fmt, bold=True, fill=TOTAL_FILL)

    autosize(ws, N_COLS_EXCEL, width=15)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 26

    # -- Встроенный BarChart: Выручка vs Полные затраты по блокам ----------
    if N_ROWS > 0:
        bar = BarChart()
        bar.type = "col"
        bar.title = "Выручка vs Полные затраты по блокам"
        bar.y_axis.title = "руб"
        bar.style = 10
        full_cost_idx = EXCEL_HEADERS.index("Полные затраты") + 1
        revenue_idx = EXCEL_HEADERS.index("Выручка") + 1
        data_ref = Reference(ws, min_col=full_cost_idx, max_col=revenue_idx, min_row=header_row, max_row=last_row)
        cats_ref = Reference(ws, min_col=1, min_row=first_row, max_row=last_row)
        bar.add_data(data_ref, titles_from_data=True)
        bar.set_categories(cats_ref)
        bar.series[0].graphicalProperties.solidFill = COLOR_COST
        bar.series[1].graphicalProperties.solidFill = COLOR_REVENUE
        bar.height = 10
        bar.width = 24
        anchor_col = get_column_letter(N_COLS_EXCEL + 2)
        ws.add_chart(bar, f"{anchor_col}{header_row}")

    # -- Таблица структуры выручки (живые формулы SUMPRODUCT) + PieChart --
    struct_row0 = total_row + 3
    ws.cell(row=struct_row0, column=1, value="Структура выручки проекта")
    ws.cell(row=struct_row0, column=1).font = TITLE_FONT
    struct_header_row = struct_row0 + 1
    ws.cell(row=struct_header_row, column=1, value="Статья выручки")
    ws.cell(row=struct_header_row, column=2, value="Сумма, руб")
    style_header_row(ws, struct_header_row, 2)

    if N_ROWS > 0:
        rng_type = f'${COL["Тип блока"]}${first_row}:${COL["Тип блока"]}${last_row}'
        rng_apt = f'${COL["S квартир, м2"]}${first_row}:${COL["S квартир, м2"]}${last_row}'
        rng_p_apt = f'${COL["Цена жилья, руб/м2"]}${first_row}:${COL["Цена жилья, руб/м2"]}${last_row}'
        rng_c1 = f'${COL["S коммерции 1 эт., м2"]}${first_row}:${COL["S коммерции 1 эт., м2"]}${last_row}'
        rng_p_c1 = f'${COL["Цена коммерции, руб/м2"]}${first_row}:${COL["Цена коммерции, руб/м2"]}${last_row}'
        rng_storage = f'${COL["S кладовых, м2"]}${first_row}:${COL["S кладовых, м2"]}${last_row}'
        rng_p_storage = f'${COL["Цена кладовых, руб/м2"]}${first_row}:${COL["Цена кладовых, руб/м2"]}${last_row}'
        rng_underground = f'${COL["Подземный паркинг, м/м"]}${first_row}:${COL["Подземный паркинг, м/м"]}${last_row}'
        rng_p_underground = f'${COL["Цена подземного м/м, руб"]}${first_row}:${COL["Цена подземного м/м, руб"]}${last_row}'
        rng_ground = f'${COL["Наземный/Многоуровневый паркинг, м/м"]}${first_row}:${COL["Наземный/Многоуровневый паркинг, м/м"]}${last_row}'
        rng_p_ground = f'${COL["Цена наземного м/м, руб"]}${first_row}:${COL["Цена наземного м/м, руб"]}${last_row}'

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
        c1 = ws.cell(row=r, column=1, value=label)
        c2 = ws.cell(row=r, column=2, value=formula)
        style_cell(c1)
        style_cell(c2, number_format=MONEY_FMT)
    struct_last_row = struct_header_row + len(struct_formulas)

    pie = PieChart()
    pie.title = "Структура выручки проекта"
    data_ref = Reference(ws, min_col=2, min_row=struct_header_row, max_row=struct_last_row)
    cats_ref = Reference(ws, min_col=1, min_row=struct_header_row + 1, max_row=struct_last_row)
    pie.add_data(data_ref, titles_from_data=True)
    pie.set_categories(cats_ref)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.series[0].data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=PIE_COLORS[i % len(PIE_COLORS)]))
        for i in range(len(struct_formulas))
    ]
    pie.height = 10
    pie.width = 16
    anchor_col2 = get_column_letter(N_COLS_EXCEL + 2)
    ws.add_chart(pie, f"{anchor_col2}{struct_header_row}")

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


excel_bytes = build_excel_report()
st.divider()
st.download_button(
    label="📥 Скачать отчет в Excel (живые формулы + графики)",
    data=excel_bytes,
    file_name=f"financial_model_{scenario_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
