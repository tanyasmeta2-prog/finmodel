# -*- coding: utf-8 -*-
"""
Финансовая модель девелоперского проекта (Streamlit + openpyxl)
=================================================================
Полностью универсальный расчет экономики проекта (от 1 до N урбан-блоков)
на уровне Валовой прибыли и Валовой рентабельности (без налогов и кредитов).

Ключевая логика:
  1. Себестоимость и распределение косвенных затрат ведутся от суммарной
     продаваемой площади блока (жилье + вся коммерция + кладовые).
  2. Подземный паркинг жестко закреплен за своим урбан-блоком — его выручка
     и прямой СМР считаются строго внутри блока.
  3. Наземные паркинги — отдельный изолированный пул, не распределяется
     на блоки, влияет только на консолидированные показатели проекта.
  4. Цены продажи и ставки СМР задаются ИНДИВИДУАЛЬНО для каждого блока
     (в отличие от глобальных цен — каждый УБ может иметь свою цену/ставку).

Выгрузка в Excel (openpyxl) собирает отформатированный файл с 3 вкладками
и встроенными графиками Excel (BarChart, PieChart), которые ссылаются на
динамические диапазоны (подстраиваются под число урбан-блоков).

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
# 0. ЦВЕТОВАЯ ПАЛИТРА (валидированная категориальная палитра, фиксированный
#    порядок слотов — используется и в веб-графиках Plotly, и в Excel)
# ======================================================================
PALETTE = {
    "blue": "2a78d6",
    "orange": "eb6834",
    "aqua": "1baf7a",
    "yellow": "eda100",
    "magenta": "e87ba4",
    "violet": "4a3aa7",
    "red": "e34948",
    "green": "008300",
}
COLOR_REVENUE = PALETTE["blue"]
COLOR_COST = PALETTE["orange"]
COLOR_POSITIVE = PALETTE["blue"]
COLOR_NEGATIVE = PALETTE["red"]
COLOR_DIRECT_COST = PALETTE["blue"]
COLOR_ALLOC_COST = PALETTE["violet"]
PIE_COLORS = [
    PALETTE["blue"],     # Жилье
    PALETTE["orange"],   # Коммерция
    PALETTE["aqua"],     # Кладовые
    PALETTE["yellow"],   # Подземные м/м
    PALETTE["magenta"],  # Наземные паркинги
]

# ======================================================================
# 1. НАСТРОЙКИ СТРАНИЦЫ И СЦЕНАРИИ
# ======================================================================
st.set_page_config(page_title="Финмодель девелоперского проекта", layout="wide")

# Коэффициенты применяются ко ВСЕЙ выручке (revenue) и ко ВСЕМ затратам (cost):
# прямым, аллоцированным и затратам наземных паркингов.
SCENARIOS = {
    "Базовый": {"revenue": 1.00, "cost": 1.00},
    "Стресс": {"revenue": 0.85, "cost": 1.15},         # выручка -15%, затраты +15%
    "Оптимистичный": {"revenue": 1.10, "cost": 0.95},  # выручка +10%, затраты -5%
}

# ======================================================================
# 2. КОЛОНКИ ТАБЛИЦЫ УРБАН-БЛОКОВ И ТЕСТОВЫЕ ДАННЫЕ (13 БЛОКОВ)
# ======================================================================
NUMERIC_COLS = [
    "S квартир, м2",
    "Цена квартир, руб/м2",
    "S коммерции 1 эт., м2",
    "Цена коммерции 1 эт., руб/м2",
    "S коммерции стилобата, м2",
    "Цена коммерции стилобата, руб/м2",
    "S кладовых, м2",
    "Цена кладовых, руб/м2",
    "Подземный паркинг, м/м",
    "Цена подз. паркинга, руб/м-место",
    "Ставка СМР площади, руб/м2",
    "Ставка СМР подз. паркинга, руб/м-место",
]
ALL_BLOCK_COLS = ["Название блока"] + NUMERIC_COLS


def generate_default_blocks(n: int = 13) -> pd.DataFrame:
    """Генерирует детерминированный набор тестовых ТЭП на n урбан-блоков.

    Используется как стартовое наполнение таблицы — пользователь может
    свободно менять, добавлять и удалять строки прямо в интерфейсе.
    """
    rng = np.random.default_rng(7)  # фиксированный seed — данные воспроизводимы
    rows = []
    for i in range(1, n + 1):
        s_apt = int(rng.integers(6000, 16000) // 100 * 100)
        s_c1 = int(rng.integers(150, 700) // 10 * 10)
        # часть блоков без коммерции в стилобате — проверяет расчет с нулями
        s_stylobate = 0 if rng.random() < 0.3 else int(rng.integers(150, 600) // 10 * 10)
        s_storage = int(rng.integers(50, 400) // 10 * 10)
        parking = int(rng.integers(60, 170) // 5 * 5)

        rows.append(
            {
                "Название блока": f"УБ {i}",
                "S квартир, м2": s_apt,
                "Цена квартир, руб/м2": int(rng.integers(150, 221) * 1000),
                "S коммерции 1 эт., м2": s_c1,
                "Цена коммерции 1 эт., руб/м2": int(rng.integers(200, 321) * 1000),
                "S коммерции стилобата, м2": s_stylobate,
                "Цена коммерции стилобата, руб/м2": int(rng.integers(120, 201) * 1000),
                "S кладовых, м2": s_storage,
                "Цена кладовых, руб/м2": int(rng.integers(40, 91) * 1000),
                "Подземный паркинг, м/м": parking,
                "Цена подз. паркинга, руб/м-место": int(rng.integers(700, 1001) * 1000),
                "Ставка СМР площади, руб/м2": int(rng.integers(50, 69) * 1000),
                "Ставка СМР подз. паркинга, руб/м-место": int(rng.integers(650, 821) * 1000),
            }
        )
    return pd.DataFrame(rows)[ALL_BLOCK_COLS]


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
# 4. БОКОВАЯ ПАНЕЛЬ — СЦЕНАРИЙ И ГЛОБАЛЬНЫЕ ИНПУТЫ
# ======================================================================
with st.sidebar:
    st.header("Сценарий расчета")
    scenario_name = st.selectbox("Выберите сценарий", list(SCENARIOS.keys()))
    rev_factor = SCENARIOS[scenario_name]["revenue"]
    cost_factor = SCENARIOS[scenario_name]["cost"]

    st.header("Пул косвенных расходов, руб")
    cost_land = st.number_input("Земля", min_value=0.0, value=500_000_000.0, step=1_000_000.0)
    cost_infra = st.number_input("Сети / Инфраструктура", min_value=0.0, value=300_000_000.0, step=1_000_000.0)
    cost_landscape = st.number_input("Благоустройство / Дороги", min_value=0.0, value=150_000_000.0, step=1_000_000.0)
    cost_social = st.number_input("Социальные объекты (школы/сады)", min_value=0.0, value=400_000_000.0, step=1_000_000.0)
    cost_soft = st.number_input("Прочие Soft Costs", min_value=0.0, value=100_000_000.0, step=1_000_000.0)
    indirect_pool_total = cost_land + cost_infra + cost_landscape + cost_social + cost_soft
    st.caption(f"Итого пул косвенных расходов: {indirect_pool_total:,.0f} руб".replace(",", " "))

    st.header("Наземные паркинги (изолированный пул)")
    ground_count = st.number_input("Кол-во м/м наземных", min_value=0, value=300, step=10)
    ground_price = st.number_input("Цена продажи 1 м/м, руб", min_value=0.0, value=500000.0, step=10000.0)
    ground_unit_cost = st.number_input("Себестоимость 1 м/м, руб", min_value=0.0, value=350000.0, step=10000.0)

# ======================================================================
# 5. ТАБЛИЦА ТЭП УРБАН-БЛОКОВ (st.data_editor)
# ======================================================================
if "blocks_df" not in st.session_state:
    st.session_state.blocks_df = generate_default_blocks(13)

st.subheader("ТЭП урбан-блоков")
st.caption(
    "Цены продажи и ставки СМР задаются индивидуально для каждого блока. "
    "Строки можно добавлять/удалять — расчет полностью универсален (1..N блоков)."
)

column_config = {"Название блока": st.column_config.TextColumn(required=True)}
for col in NUMERIC_COLS:
    column_config[col] = st.column_config.NumberColumn(min_value=0, format="%.0f")

edited_blocks = st.data_editor(
    st.session_state.blocks_df,
    num_rows="dynamic",
    use_container_width=True,
    key="blocks_editor",
    column_config=column_config,
)
st.session_state.blocks_df = edited_blocks

# Приведение типов и отбрасывание пустых строк
blocks = edited_blocks.copy()
for col in NUMERIC_COLS:
    blocks[col] = pd.to_numeric(blocks[col], errors="coerce").fillna(0.0)
blocks["Название блока"] = blocks["Название блока"].fillna("").astype(str)
blocks = blocks[~((blocks["Название блока"] == "") & (blocks[NUMERIC_COLS].sum(axis=1) == 0))]
blocks = blocks.reset_index(drop=True)
N_BLOCKS = len(blocks)

# ======================================================================
# 6. РАСЧЕТ ЭКОНОМИКИ УРБАН-БЛОКОВ
# ======================================================================
# Суммарная продаваемая площадь блока = жилье + вся коммерция + кладовые
blocks["Площадь всего, м2"] = (
    blocks["S квартир, м2"]
    + blocks["S коммерции 1 эт., м2"]
    + blocks["S коммерции стилобата, м2"]
    + blocks["S кладовых, м2"]
)
total_area = blocks["Площадь всего, м2"].sum()

# Доля аллокации косвенных расходов пропорционально суммарной площади блока
blocks["Доля аллокации"] = blocks["Площадь всего, м2"] / total_area if total_area > 0 else 0.0

indirect_pool_scenario = indirect_pool_total * cost_factor
blocks["Аллоцированные затраты"] = blocks["Доля аллокации"] * indirect_pool_scenario

# Прямые затраты блока: СМР продаваемой площади (по ИНДИВИДУАЛЬНОЙ ставке блока)
# + СМР подземного паркинга блока (по ИНДИВИДУАЛЬНОЙ ставке блока)
blocks["Прямые затраты"] = (
    blocks["Площадь всего, м2"] * blocks["Ставка СМР площади, руб/м2"]
    + blocks["Подземный паркинг, м/м"] * blocks["Ставка СМР подз. паркинга, руб/м-место"]
) * cost_factor

blocks["Полные затраты"] = blocks["Прямые затраты"] + blocks["Аллоцированные затраты"]

# Выручка блока по индивидуальным ценам блока (жилье + коммерция 1эт +
# коммерция стилобата + кладовые + подземный паркинг — все внутри блока)
rev_apt = blocks["S квартир, м2"] * blocks["Цена квартир, руб/м2"]
rev_c1 = blocks["S коммерции 1 эт., м2"] * blocks["Цена коммерции 1 эт., руб/м2"]
rev_stylobate = blocks["S коммерции стилобата, м2"] * blocks["Цена коммерции стилобата, руб/м2"]
rev_storage = blocks["S кладовых, м2"] * blocks["Цена кладовых, руб/м2"]
rev_underground = blocks["Подземный паркинг, м/м"] * blocks["Цена подз. паркинга, руб/м-место"]

blocks["Выручка"] = (rev_apt + rev_c1 + rev_stylobate + rev_storage + rev_underground) * rev_factor
blocks["Валовая прибыль"] = blocks["Выручка"] - blocks["Полные затраты"]
blocks["Рентабельность"] = np.where(
    blocks["Выручка"] > 0, blocks["Валовая прибыль"] / blocks["Выручка"], 0.0
)  # хранится долей (0..1) — на web/Excel выводится в процентах

# Компоненты выручки по типам (со сценарным коэффициентом) — для структуры выручки
revenue_components = {
    "Жилье": (rev_apt * rev_factor).sum(),
    "Коммерция": ((rev_c1 + rev_stylobate) * rev_factor).sum(),
    "Кладовые": (rev_storage * rev_factor).sum(),
    "Подземные м/м": (rev_underground * rev_factor).sum(),
}

# ======================================================================
# 7. РАСЧЕТ НАЗЕМНЫХ ПАРКИНГОВ (изолированный пул)
# ======================================================================
ground_revenue = ground_count * ground_price * rev_factor
ground_cost = ground_count * ground_unit_cost * cost_factor
ground_profit = ground_revenue - ground_cost
ground_margin = (ground_profit / ground_revenue) if ground_revenue > 0 else 0.0

revenue_components["Наземные паркинги"] = ground_revenue

# ======================================================================
# 8. КОНСОЛИДИРОВАННЫЕ ПОКАЗАТЕЛИ ПРОЕКТА
# ======================================================================
total_revenue = blocks["Выручка"].sum() + ground_revenue
total_cost = blocks["Полные затраты"].sum() + ground_cost
total_profit = total_revenue - total_cost
avg_margin = (total_profit / total_revenue) if total_revenue > 0 else 0.0

# ======================================================================
# 9. ВЕБ-ИНТЕРФЕЙС — МЕТРИКИ И ИНТЕРАКТИВНАЯ ГРАФИКА (PLOTLY)
# ======================================================================
st.divider()
st.subheader(f"«{project_name}», {project_city} — сценарий «{scenario_name}»")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Выручка проекта", f"{total_revenue:,.0f} руб".replace(",", " "))
m2.metric("Затраты проекта", f"{total_cost:,.0f} руб".replace(",", " "))
m3.metric("Валовая прибыль", f"{total_profit:,.0f} руб".replace(",", " "))
m4.metric("Средняя рентабельность", f"{avg_margin * 100:.1f} %")

st.markdown("**Экономика урбан-блоков**")
tbl_cols = [
    "Название блока",
    "Площадь всего, м2",
    "Доля аллокации",
    "Прямые затраты",
    "Аллоцированные затраты",
    "Полные затраты",
    "Выручка",
    "Валовая прибыль",
    "Рентабельность",
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
    # График 1: Выручка vs Полные затраты по блокам (группированные столбцы)
    fig_rev_cost = go.Figure()
    fig_rev_cost.add_bar(
        name="Выручка", x=blocks["Название блока"], y=blocks["Выручка"],
        marker_color=f"#{COLOR_REVENUE}",
    )
    fig_rev_cost.add_bar(
        name="Полные затраты", x=blocks["Название блока"], y=blocks["Полные затраты"],
        marker_color=f"#{COLOR_COST}",
    )
    fig_rev_cost.update_layout(
        title="Выручка vs Затраты по урбан-блокам",
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(t=60, b=40),
    )
    st.plotly_chart(fig_rev_cost, use_container_width=True)

with chart_col2:
    # График 2: Рентабельность по блокам — диверг. цвет (полярность выше/ниже нуля)
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
    fig_margin.update_layout(
        title="Валовая рентабельность по блокам, %",
        yaxis_title="%",
        margin=dict(t=60, b=40),
    )
    fig_margin.add_hline(y=0, line_color="#898781", line_width=1)
    st.plotly_chart(fig_margin, use_container_width=True)

chart_col3, chart_col4 = st.columns(2)

with chart_col3:
    # График 3: Структура выручки проекта (Итого)
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
    # График 4: Структура затрат блока — прямые vs аллоцированные (stacked)
    fig_cost_structure = go.Figure()
    fig_cost_structure.add_bar(
        name="Прямые затраты", x=blocks["Название блока"], y=blocks["Прямые затраты"],
        marker_color=f"#{COLOR_DIRECT_COST}",
    )
    fig_cost_structure.add_bar(
        name="Аллоцированные затраты", x=blocks["Название блока"], y=blocks["Аллоцированные затраты"],
        marker_color=f"#{COLOR_ALLOC_COST}",
    )
    fig_cost_structure.update_layout(
        title="Структура затрат по блокам",
        barmode="stack",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(t=60, b=40),
    )
    st.plotly_chart(fig_cost_structure, use_container_width=True)

st.markdown("**Экономика наземных паркингов (изолированный пул)**")
ground_df = pd.DataFrame(
    {
        "Показатель": [
            "Кол-во м/м",
            "Цена продажи, руб/м-место",
            "Выручка, руб",
            "Себестоимость 1 м/м, руб",
            "Затраты, руб",
            "Валовая прибыль, руб",
            "Рентабельность, %",
        ],
        "Значение": [
            ground_count,
            ground_price,
            ground_revenue,
            ground_unit_cost,
            ground_cost,
            ground_profit,
            ground_margin * 100,
        ],
    }
)
st.dataframe(ground_df, use_container_width=True, hide_index=True)

# ======================================================================
# 10. ЭКСПОРТ В EXCEL (openpyxl) — форматирование + встроенные графики
# ======================================================================
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=13)
TOTAL_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
TOTAL_FONT = Font(bold=True)
MONEY_FMT = "#,##0"
PERCENT_FMT = "0.0%"


def write_passport(ws, scenario: str, extra_rows=0):
    """Пишет паспорт проекта (наименование, город, сценарий) в верхние строки листа."""
    ws["A1"] = f"Проект: {project_name}"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Город: {project_city}"
    ws["A2"].font = TITLE_FONT
    ws["A3"] = f"Сценарий: {scenario}"
    ws["A3"].font = TITLE_FONT
    return 5 + extra_rows  # номер строки, с которой начинается таблица


def style_header_row(ws, row: int, n_cols: int):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_data_cell(cell, number_format=None, bold=False, fill=None):
    cell.border = BORDER
    if number_format:
        cell.number_format = number_format
    if bold:
        cell.font = TOTAL_FONT
    if fill:
        cell.fill = fill


def autosize_columns(ws, n_cols: int, width: int = 16):
    for c in range(1, n_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def build_excel_report() -> bytes:
    """Собирает многостраничный Excel-отчет со встроенными графиками и возвращает байты."""
    wb = Workbook()

    # ------------------------------------------------------------------
    # Лист 1: Экономика Урбан-Блоков
    # ------------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Экономика Урбан-Блоков"
    header_row = write_passport(ws1, scenario_name)

    headers = [
        "Название блока", "Площадь всего, м2", "Доля аллокации",
        "Прямые затраты", "Аллоцированные затраты", "Полные затраты",
        "Выручка", "Валовая прибыль", "Рентабельность",
    ]
    for j, h in enumerate(headers, start=1):
        ws1.cell(row=header_row, column=j, value=h)
    style_header_row(ws1, header_row, len(headers))

    first_data_row = header_row + 1
    for i, row in blocks.iterrows():
        r = first_data_row + i
        values = [
            row["Название блока"], row["Площадь всего, м2"], row["Доля аллокации"],
            row["Прямые затраты"], row["Аллоцированные затраты"], row["Полные затраты"],
            row["Выручка"], row["Валовая прибыль"], row["Рентабельность"],
        ]
        for j, v in enumerate(values, start=1):
            cell = ws1.cell(row=r, column=j, value=v)
            if j == 1:
                style_data_cell(cell)
            elif j == 3 or j == 9:  # Доля аллокации, Рентабельность — проценты
                style_data_cell(cell, number_format=PERCENT_FMT)
            else:
                style_data_cell(cell, number_format=MONEY_FMT)

    last_data_row = first_data_row + N_BLOCKS - 1
    total_row = last_data_row + 1 if N_BLOCKS > 0 else first_data_row

    # Итоговая строка
    ws1.cell(row=total_row, column=1, value="ИТОГО")
    if N_BLOCKS > 0:
        col_letters = {j: get_column_letter(j) for j in range(1, 10)}
        ws1.cell(row=total_row, column=2, value=f"=SUM({col_letters[2]}{first_data_row}:{col_letters[2]}{last_data_row})")
        ws1.cell(row=total_row, column=3, value=f"=SUM({col_letters[3]}{first_data_row}:{col_letters[3]}{last_data_row})")
        ws1.cell(row=total_row, column=4, value=f"=SUM({col_letters[4]}{first_data_row}:{col_letters[4]}{last_data_row})")
        ws1.cell(row=total_row, column=5, value=f"=SUM({col_letters[5]}{first_data_row}:{col_letters[5]}{last_data_row})")
        ws1.cell(row=total_row, column=6, value=f"=SUM({col_letters[6]}{first_data_row}:{col_letters[6]}{last_data_row})")
        ws1.cell(row=total_row, column=7, value=f"=SUM({col_letters[7]}{first_data_row}:{col_letters[7]}{last_data_row})")
        ws1.cell(row=total_row, column=8, value=f"=SUM({col_letters[8]}{first_data_row}:{col_letters[8]}{last_data_row})")
        ws1.cell(row=total_row, column=9, value=f"=IF(G{total_row}=0,0,H{total_row}/G{total_row})")
    else:
        # Блоков нет — заполняем итоговую строку нулями, а не пустыми ячейками
        for j in range(2, 10):
            ws1.cell(row=total_row, column=j, value=0)
    for j in range(1, 10):
        cell = ws1.cell(row=total_row, column=j)
        fmt = PERCENT_FMT if j in (3, 9) else (MONEY_FMT if j >= 2 else None)
        style_data_cell(cell, number_format=fmt, bold=True, fill=TOTAL_FILL)

    autosize_columns(ws1, len(headers))
    ws1.column_dimensions["A"].width = 16

    # Встроенный график: Выручка vs Полные затраты по блокам (BarChart)
    if N_BLOCKS > 0:
        bar = BarChart()
        bar.type = "col"
        bar.title = "Выручка vs Полные затраты по урбан-блокам"
        bar.y_axis.title = "руб"
        bar.style = 10

        # Данные: колонка F (Полные затраты) и G (Выручка), с заголовками в header_row
        data_ref = Reference(ws1, min_col=6, max_col=7, min_row=header_row, max_row=last_data_row)
        cats_ref = Reference(ws1, min_col=1, min_row=first_data_row, max_row=last_data_row)
        bar.add_data(data_ref, titles_from_data=True)
        bar.set_categories(cats_ref)

        # Цвета серий — из той же категориальной палитры, что и веб-графики
        bar.series[0].graphicalProperties.solidFill = COLOR_COST      # Полные затраты
        bar.series[1].graphicalProperties.solidFill = COLOR_REVENUE   # Выручка
        bar.height = 10
        bar.width = 22
        ws1.add_chart(bar, "N5")

    # ------------------------------------------------------------------
    # Лист 2: Наземные паркинги
    # ------------------------------------------------------------------
    ws2 = wb.create_sheet("Наземные паркинги")
    header_row2 = write_passport(ws2, scenario_name)
    ws2.cell(row=header_row2 - 1, column=1, value="Изолированный пул — не распределяется на урбан-блоки")

    ws2.cell(row=header_row2, column=1, value="Показатель")
    ws2.cell(row=header_row2, column=2, value="Значение")
    style_header_row(ws2, header_row2, 2)

    ground_rows = [
        ("Кол-во м/м", ground_count, MONEY_FMT),
        ("Цена продажи, руб/м-место", ground_price, MONEY_FMT),
        ("Выручка, руб", ground_revenue, MONEY_FMT),
        ("Себестоимость 1 м/м, руб", ground_unit_cost, MONEY_FMT),
        ("Затраты, руб", ground_cost, MONEY_FMT),
        ("Валовая прибыль, руб", ground_profit, MONEY_FMT),
        ("Рентабельность", ground_margin, PERCENT_FMT),
    ]
    for i, (label, value, fmt) in enumerate(ground_rows):
        r = header_row2 + 1 + i
        c1 = ws2.cell(row=r, column=1, value=label)
        c2 = ws2.cell(row=r, column=2, value=value)
        style_data_cell(c1)
        style_data_cell(c2, number_format=fmt)
    autosize_columns(ws2, 2, width=28)

    # ------------------------------------------------------------------
    # Лист 3: Консолидированный Дашборд
    # ------------------------------------------------------------------
    ws3 = wb.create_sheet("Консолидированный Дашборд")
    header_row3 = write_passport(ws3, scenario_name)

    ws3.cell(row=header_row3, column=1, value="Показатель")
    ws3.cell(row=header_row3, column=2, value="Значение")
    style_header_row(ws3, header_row3, 2)

    summary_rows = [
        ("Выручка проекта, руб", total_revenue, MONEY_FMT),
        ("Затраты проекта, руб", total_cost, MONEY_FMT),
        ("Валовая прибыль проекта, руб", total_profit, MONEY_FMT),
        ("Средняя рентабельность", avg_margin, PERCENT_FMT),
    ]
    for i, (label, value, fmt) in enumerate(summary_rows):
        r = header_row3 + 1 + i
        c1 = ws3.cell(row=r, column=1, value=label)
        c2 = ws3.cell(row=r, column=2, value=value)
        style_data_cell(c1, bold=True)
        style_data_cell(c2, number_format=fmt, bold=True)

    # Таблица структуры выручки (источник данных для круговой диаграммы)
    struct_header_row = header_row3 + len(summary_rows) + 2
    ws3.cell(row=struct_header_row, column=1, value="Статья выручки")
    ws3.cell(row=struct_header_row, column=2, value="Сумма, руб")
    style_header_row(ws3, struct_header_row, 2)

    struct_items = list(revenue_components.items())
    for i, (label, value) in enumerate(struct_items):
        r = struct_header_row + 1 + i
        c1 = ws3.cell(row=r, column=1, value=label)
        c2 = ws3.cell(row=r, column=2, value=value)
        style_data_cell(c1)
        style_data_cell(c2, number_format=MONEY_FMT)
    struct_last_row = struct_header_row + len(struct_items)

    autosize_columns(ws3, 2, width=30)

    # Встроенный график: круговая диаграмма структуры выручки
    pie = PieChart()
    pie.title = "Структура выручки проекта"
    data_ref = Reference(ws3, min_col=2, min_row=struct_header_row, max_row=struct_last_row)
    cats_ref = Reference(ws3, min_col=1, min_row=struct_header_row + 1, max_row=struct_last_row)
    pie.add_data(data_ref, titles_from_data=True)
    pie.set_categories(cats_ref)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    # Раскраска долей той же категориальной палитрой, что и веб-график
    pie.series[0].data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=PIE_COLORS[i % len(PIE_COLORS)]))
        for i in range(len(struct_items))
    ]
    pie.height = 10
    pie.width = 16
    ws3.add_chart(pie, "E5")

    # ------------------------------------------------------------------
    # Сохранение в байты
    # ------------------------------------------------------------------
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


excel_bytes = build_excel_report()
st.divider()
st.download_button(
    label="📥 Скачать отчет в Excel (с графиками)",
    data=excel_bytes,
    file_name=f"financial_model_{scenario_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
