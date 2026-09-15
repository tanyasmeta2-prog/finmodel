# -*- coding: utf-8 -*-
"""
Финансовая модель девелоперского проекта (Streamlit)
======================================================
Универсальный расчет экономики проекта на уровне Валовой прибыли и
Валовой рентабельности (без налогов и кредитов).

Логика:
  1. Себестоимость и распределение затрат ведутся от продаваемой площади
     (жилье + коммерция). Площадь ГНС не используется.
  2. Подземный паркинг жестко закреплен за своим урбан-блоком.
  3. Наземные/многоуровневые паркинги — отдельный пул, не распределяется
     на блоки, влияет только на консолидированные показатели.

Запуск:
    streamlit run app.py
"""

import io
import pandas as pd
import streamlit as st

# ======================================================================
# 1. НАСТРОЙКИ СТРАНИЦЫ
# ======================================================================
st.set_page_config(
    page_title="Финмодель девелоперского проекта",
    layout="wide",
)

st.title("Финансовая модель девелоперского проекта")
st.caption(
    "Валовая прибыль и валовая рентабельность. Налоги и кредиты не учитываются."
)

# ======================================================================
# 2. СЦЕНАРИИ
# ======================================================================
# Коэффициенты применяются ко ВСЕЙ выручке (revenue_factor) и
# ко ВСЕМ затратам (cost_factor) — прямым, аллоцированным и паркингам.
SCENARIOS = {
    "Базовый": {"revenue": 1.00, "cost": 1.00},
    "Стресс": {"revenue": 0.85, "cost": 1.15},          # выручка -15%, затраты +15%
    "Оптимистичный": {"revenue": 1.10, "cost": 0.95},   # выручка +10%, затраты -5%
}

# ======================================================================
# 3. ТЕСТОВЫЕ ДАННЫЕ ПО УМОЛЧАНИЮ (13 урбан-блоков)
# ======================================================================
DEFAULT_BLOCKS = pd.DataFrame(
    {
        "Название блока": [f"УБ {i}" for i in range(1, 14)],
        "S квартир, м2": [
            12000, 9500, 15000, 8000, 11000, 13500, 7000,
            10000, 14000, 9000, 12500, 6500, 16000,
        ],
        "S коммерции 1 эт., м2": [
            450, 380, 600, 250, 400, 500, 200,
            350, 550, 300, 470, 180, 650,
        ],
        "S коммерции стилобата, м2": [
            300, 0, 500, 200, 0, 350, 0,
            250, 400, 0, 320, 0, 550,
        ],
        "Подземный паркинг, м/м": [
            120, 95, 150, 80, 110, 135, 70,
            100, 140, 90, 125, 65, 160,
        ],
        "Ставка СМР продаваемой пл., руб/м2": [
            58000, 56000, 62000, 54000, 57000, 60000, 53000,
            55000, 61000, 55500, 59000, 52000, 63000,
        ],
        "Ставка СМР подз. паркинга, руб/м-место": [
            750000, 720000, 780000, 700000, 740000, 760000, 690000,
            710000, 770000, 715000, 755000, 680000, 790000,
        ],
    }
)

# Инициализация состояния таблицы блоков (чтобы правки/добавление/
# удаление строк сохранялись между перерисовками страницы)
if "blocks_df" not in st.session_state:
    st.session_state.blocks_df = DEFAULT_BLOCKS.copy()

# ======================================================================
# 4. БОКОВАЯ ПАНЕЛЬ — ГЛОБАЛЬНЫЕ ИНПУТЫ
# ======================================================================
with st.sidebar:
    st.header("Сценарий расчета")
    scenario_name = st.selectbox("Выберите сценарий", list(SCENARIOS.keys()))
    rev_factor = SCENARIOS[scenario_name]["revenue"]
    cost_factor = SCENARIOS[scenario_name]["cost"]

    st.header("Цены продажи, руб/м2")
    price_apt = st.number_input("Жилье", min_value=0.0, value=180000.0, step=1000.0)
    price_c1 = st.number_input("Коммерция 1 эт.", min_value=0.0, value=250000.0, step=1000.0)
    price_cs = st.number_input("Коммерция стилобат", min_value=0.0, value=150000.0, step=1000.0)

    st.header("Цены продажи паркингов, руб/м-место")
    price_park_underground = st.number_input(
        "Подземный паркинг", min_value=0.0, value=900000.0, step=10000.0
    )
    price_park_ground = st.number_input(
        "Наземный паркинг", min_value=0.0, value=500000.0, step=10000.0
    )
    price_park_multilevel = st.number_input(
        "Многоуровневый паркинг", min_value=0.0, value=650000.0, step=10000.0
    )

    st.header("Пул косвенных расходов, руб")
    cost_land = st.number_input("Земля", min_value=0.0, value=500_000_000.0, step=1_000_000.0)
    cost_infra = st.number_input(
        "Сети / Инфраструктура", min_value=0.0, value=300_000_000.0, step=1_000_000.0
    )
    cost_landscape = st.number_input(
        "Благоустройство / Дороги", min_value=0.0, value=150_000_000.0, step=1_000_000.0
    )
    cost_social = st.number_input(
        "Социальные объекты (школы/сады)", min_value=0.0, value=400_000_000.0, step=1_000_000.0
    )
    cost_other_soft = st.number_input(
        "Прочие Soft Costs", min_value=0.0, value=100_000_000.0, step=1_000_000.0
    )
    indirect_pool_total = (
        cost_land + cost_infra + cost_landscape + cost_social + cost_other_soft
    )
    st.caption(f"Итого пул косвенных расходов: {indirect_pool_total:,.0f} руб".replace(",", " "))

    st.header("Наземные / многоуровневые паркинги")
    ground_count = st.number_input("Кол-во м/м наземных", min_value=0, value=300, step=10)
    ground_unit_cost = st.number_input(
        "Себестоимость 1 м/м наземного, руб", min_value=0.0, value=350000.0, step=10000.0
    )
    multilevel_count = st.number_input(
        "Кол-во м/м многоуровневых", min_value=0, value=150, step=10
    )
    multilevel_unit_cost = st.number_input(
        "Себестоимость 1 м/м многоуровневого, руб", min_value=0.0, value=550000.0, step=10000.0
    )

# ======================================================================
# 5. ДИНАМИЧЕСКАЯ ТАБЛИЦА ТЭП УРБАН-БЛОКОВ
# ======================================================================
st.subheader("ТЭП урбан-блоков")
st.caption("Добавляйте / удаляйте строки прямо в таблице. Расчет обновляется на лету.")

edited_blocks = st.data_editor(
    st.session_state.blocks_df,
    num_rows="dynamic",          # позволяет добавлять и удалять строки (этапы)
    use_container_width=True,
    key="blocks_editor",
    column_config={
        "Название блока": st.column_config.TextColumn(required=True),
        "S квартир, м2": st.column_config.NumberColumn(min_value=0, format="%.0f"),
        "S коммерции 1 эт., м2": st.column_config.NumberColumn(min_value=0, format="%.0f"),
        "S коммерции стилобата, м2": st.column_config.NumberColumn(min_value=0, format="%.0f"),
        "Подземный паркинг, м/м": st.column_config.NumberColumn(min_value=0, format="%.0f"),
        "Ставка СМР продаваемой пл., руб/м2": st.column_config.NumberColumn(
            min_value=0, format="%.0f"
        ),
        "Ставка СМР подз. паркинга, руб/м-место": st.column_config.NumberColumn(
            min_value=0, format="%.0f"
        ),
    },
)
# Сохраняем правки в состояние, чтобы они не терялись при перерисовке
st.session_state.blocks_df = edited_blocks

# Приводим числовые колонки к float и заполняем пустые ячейки нулями
# (пользователь мог добавить пустую строку и не успеть заполнить её)
NUMERIC_COLS = [
    "S квартир, м2",
    "S коммерции 1 эт., м2",
    "S коммерции стилобата, м2",
    "Подземный паркинг, м/м",
    "Ставка СМР продаваемой пл., руб/м2",
    "Ставка СМР подз. паркинга, руб/м-место",
]
blocks = edited_blocks.copy()
for col in NUMERIC_COLS:
    blocks[col] = pd.to_numeric(blocks[col], errors="coerce").fillna(0.0)
blocks["Название блока"] = blocks["Название блока"].fillna("").astype(str)
# Отбрасываем полностью пустые строки (например, только что добавленную)
blocks = blocks[~((blocks["Название блока"] == "") & (blocks[NUMERIC_COLS].sum(axis=1) == 0))]
blocks = blocks.reset_index(drop=True)

# ======================================================================
# 6. РАСЧЕТ ЭКОНОМИКИ УРБАН-БЛОКОВ
# ======================================================================
# Модель полностью универсальна: работает для 1 блока и для N блоков,
# так как все формулы построены на суммах/долях по строкам таблицы blocks.

# 6.1. Суммарная продаваемая площадь блока (жилье + коммерция 1 эт. + стилобат)
blocks["Суммарная продаваемая площадь, м2"] = (
    blocks["S квартир, м2"] + blocks["S коммерции 1 эт., м2"] + blocks["S коммерции стилобата, м2"]
)

total_sellable_area = blocks["Суммарная продаваемая площадь, м2"].sum()

# 6.2. Доля аллокации блока в общем пуле косвенных расходов
if total_sellable_area > 0:
    blocks["Доля аллокации"] = blocks["Суммарная продаваемая площадь, м2"] / total_sellable_area
else:
    blocks["Доля аллокации"] = 0.0
# Та же доля в процентах — только для удобного отображения в таблицах
blocks["Доля аллокации, %"] = blocks["Доля аллокации"] * 100

# 6.3. Аллоцированные косвенные затраты блока (с учетом сценария по затратам)
indirect_pool_scenario = indirect_pool_total * cost_factor
blocks["Аллоцированные общие затраты, руб"] = blocks["Доля аллокации"] * indirect_pool_scenario

# 6.4. Прямые затраты блока: СМР площади + СМР подземного паркинга блока
blocks["Прямые затраты, руб"] = (
    blocks["Суммарная продаваемая площадь, м2"] * blocks["Ставка СМР продаваемой пл., руб/м2"]
    + blocks["Подземный паркинг, м/м"] * blocks["Ставка СМР подз. паркинга, руб/м-место"]
) * cost_factor

# 6.5. Полные затраты блока
blocks["Полные затраты, руб"] = blocks["Прямые затраты, руб"] + blocks["Аллоцированные общие затраты, руб"]

# 6.6. Выручка блока: квартиры + коммерция 1 эт. + коммерция стилобат + подземный паркинг блока
blocks["Выручка, руб"] = (
    blocks["S квартир, м2"] * price_apt
    + blocks["S коммерции 1 эт., м2"] * price_c1
    + blocks["S коммерции стилобата, м2"] * price_cs
    + blocks["Подземный паркинг, м/м"] * price_park_underground
) * rev_factor

# 6.7. Валовая прибыль и рентабельность блока
blocks["Валовая прибыль, руб"] = blocks["Выручка, руб"] - blocks["Полные затраты, руб"]
blocks["Валовая рентабельность, %"] = blocks.apply(
    lambda r: (r["Валовая прибыль, руб"] / r["Выручка, руб"] * 100) if r["Выручка, руб"] > 0 else 0.0,
    axis=1,
)

# ======================================================================
# 7. РАСЧЕТ НАЗЕМНЫХ / МНОГОУРОВНЕВЫХ ПАРКИНГОВ (отдельный изолированный пул)
# ======================================================================
parking_rows = []

ground_revenue = ground_count * price_park_ground * rev_factor
ground_cost = ground_count * ground_unit_cost * cost_factor
parking_rows.append(
    {
        "Объект": "Наземный паркинг",
        "Кол-во м/м": ground_count,
        "Выручка, руб": ground_revenue,
        "Затраты, руб": ground_cost,
        "Валовая прибыль, руб": ground_revenue - ground_cost,
        "Валовая рентабельность, %": (
            (ground_revenue - ground_cost) / ground_revenue * 100 if ground_revenue > 0 else 0.0
        ),
    }
)

multilevel_revenue = multilevel_count * price_park_multilevel * rev_factor
multilevel_cost = multilevel_count * multilevel_unit_cost * cost_factor
parking_rows.append(
    {
        "Объект": "Многоуровневый паркинг",
        "Кол-во м/м": multilevel_count,
        "Выручка, руб": multilevel_revenue,
        "Затраты, руб": multilevel_cost,
        "Валовая прибыль, руб": multilevel_revenue - multilevel_cost,
        "Валовая рентабельность, %": (
            (multilevel_revenue - multilevel_cost) / multilevel_revenue * 100
            if multilevel_revenue > 0
            else 0.0
        ),
    }
)

parking_df = pd.DataFrame(parking_rows)

# ======================================================================
# 8. КОНСОЛИДИРОВАННЫЕ ПОКАЗАТЕЛИ ПРОЕКТА (ИТОГО)
# ======================================================================
# Затраты блоков уже включают аллоцированную долю всего пула косвенных
# расходов (сумма долей аллокации = 1), поэтому пул не добавляется повторно.
total_revenue = blocks["Выручка, руб"].sum() + parking_df["Выручка, руб"].sum()
total_costs = blocks["Полные затраты, руб"].sum() + parking_df["Затраты, руб"].sum()
total_profit = total_revenue - total_costs
avg_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0.0

# ======================================================================
# 9. ВЫВОД РЕЗУЛЬТАТОВ
# ======================================================================
st.divider()
st.subheader(f"Результаты — сценарий «{scenario_name}»")

st.markdown("**1. Экономика урбан-блоков**")
blocks_display_cols = [
    "Название блока",
    "Суммарная продаваемая площадь, м2",
    "Доля аллокации, %",
    "Прямые затраты, руб",
    "Аллоцированные общие затраты, руб",
    "Полные затраты, руб",
    "Выручка, руб",
    "Валовая прибыль, руб",
    "Валовая рентабельность, %",
]
st.dataframe(
    blocks[blocks_display_cols],
    use_container_width=True,
    column_config={
        "Доля аллокации, %": st.column_config.NumberColumn(format="%.2f"),
        "Прямые затраты, руб": st.column_config.NumberColumn(format="%.0f"),
        "Аллоцированные общие затраты, руб": st.column_config.NumberColumn(format="%.0f"),
        "Полные затраты, руб": st.column_config.NumberColumn(format="%.0f"),
        "Выручка, руб": st.column_config.NumberColumn(format="%.0f"),
        "Валовая прибыль, руб": st.column_config.NumberColumn(format="%.0f"),
        "Валовая рентабельность, %": st.column_config.NumberColumn(format="%.1f"),
    },
)

st.markdown("**2. Экономика наземных / многоуровневых паркингов**")
st.dataframe(
    parking_df,
    use_container_width=True,
    column_config={
        "Выручка, руб": st.column_config.NumberColumn(format="%.0f"),
        "Затраты, руб": st.column_config.NumberColumn(format="%.0f"),
        "Валовая прибыль, руб": st.column_config.NumberColumn(format="%.0f"),
        "Валовая рентабельность, %": st.column_config.NumberColumn(format="%.1f"),
    },
)

st.divider()
st.subheader("Консолидированные показатели проекта — Итого")
col1, col2, col3, col4 = st.columns(4)
col1.metric("Выручка проекта", f"{total_revenue:,.0f} руб".replace(",", " "))
col2.metric("Затраты проекта", f"{total_costs:,.0f} руб".replace(",", " "))
col3.metric("Валовая прибыль", f"{total_profit:,.0f} руб".replace(",", " "))
col4.metric("Средняя рентабельность", f"{avg_margin:.1f} %")

with st.expander("Проверка расчета"):
    st.write(
        f"Сумма долей аллокации по блокам: {blocks['Доля аллокации'].sum():.4f} "
        "(должна быть равна 1.0, если сумма площадей блоков > 0)"
    )
    st.write(f"Пул косвенных расходов (со сценарием): {indirect_pool_scenario:,.0f} руб".replace(",", " "))
    st.write(
        f"Сумма аллоцированных затрат по блокам: "
        f"{blocks['Аллоцированные общие затраты, руб'].sum():,.0f} руб".replace(",", " ")
    )

# ======================================================================
# 10. ВЫГРУЗКА ОТЧЕТА В EXCEL (с разделением по вкладкам)
# ======================================================================
def build_excel_report() -> bytes:
    """Собирает многостраничный Excel-отчет и возвращает его как байты."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Вкладка 1: исходные ТЭП урбан-блоков
        blocks[
            [
                "Название блока",
                "S квартир, м2",
                "S коммерции 1 эт., м2",
                "S коммерции стилобата, м2",
                "Подземный паркинг, м/м",
                "Ставка СМР продаваемой пл., руб/м2",
                "Ставка СМР подз. паркинга, руб/м-место",
            ]
        ].to_excel(writer, sheet_name="Вводные ТЭП", index=False)

        # Вкладка 2: экономика урбан-блоков
        blocks[blocks_display_cols].to_excel(writer, sheet_name="Урбан-блоки", index=False)

        # Вкладка 3: экономика наземных/многоуровневых паркингов
        parking_df.to_excel(writer, sheet_name="Паркинги (пул)", index=False)

        # Вкладка 4: консолидированные показатели
        summary_df = pd.DataFrame(
            {
                "Показатель": [
                    "Сценарий",
                    "Выручка проекта, руб",
                    "Затраты проекта, руб",
                    "Валовая прибыль проекта, руб",
                    "Средняя рентабельность, %",
                ],
                "Значение": [
                    scenario_name,
                    total_revenue,
                    total_costs,
                    total_profit,
                    avg_margin,
                ],
            }
        )
        summary_df.to_excel(writer, sheet_name="Итого", index=False)

    return buffer.getvalue()


excel_bytes = build_excel_report()
st.download_button(
    label="Скачать отчет в Excel",
    data=excel_bytes,
    file_name=f"finmodel_report_{scenario_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
