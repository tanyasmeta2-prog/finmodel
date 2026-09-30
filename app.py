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

import base64
import io
import json
import re
import threading
from datetime import datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import requests
import streamlit as st
from PIL import Image as PILImage

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

# ======================================================================
# 0. ЦВЕТОВАЯ ПАЛИТРА (фиксированный порядок слотов — единая для веб- и
#    Excel-графиков)
# ======================================================================
# Фирменная палитра Талан для веб-графиков Plotly (те же цвета, что в Excel).
PALETTE = {
    "green": "41AA37", "green_dark": "2E7D27", "green2": "84D26D", "green3": "B2DAAE",
    "bordo": "AE4B67", "gray": "A6A6A6",
}
COLOR_REVENUE = PALETTE["green"]
COLOR_COST = PALETTE["gray"]
COLOR_POSITIVE = PALETTE["green"]
COLOR_NEGATIVE = PALETTE["bordo"]
COLOR_DIRECT_COST = PALETTE["green"]
COLOR_ALLOC_COST = PALETTE["green2"]
PIE_COLORS = [PALETTE["green"], PALETTE["green_dark"], PALETTE["green2"], PALETTE["bordo"], PALETTE["gray"]]
GROUP_COLORS = [PALETTE["green"], PALETTE["green_dark"], PALETTE["green2"], PALETTE["green3"], PALETTE["bordo"]]

# ----------------------------------------------------------------------
# Фирменные цвета Талан для Excel-отчета (веб-графики Plotly — PALETTE выше,
# те же фирменные цвета).
# ----------------------------------------------------------------------
XL_GREEN = "41AA37"
XL_GREEN2 = "84D26D"
XL_GREEN3 = "B2DAAE"
XL_GREEN4 = "D8E8D5"
XL_BORDO = "AE4B67"
XL_GRAY = "CECECE"
XL_TEXT = "000000"
XL_TEXT2 = "3F3F3F"
XL_SURFACE = "F2F2F2"
XL_WHITE = "FFFFFF"
XL_FONT = "Arial"
XL_GREEN_RAMP5 = [XL_GREEN, XL_GREEN2, XL_GREEN3, XL_GREEN4, XL_GRAY]

# Логотип Талан (PNG, base64) — встраивается в Excel-отчет без внешних файлов.
TALAN_LOGO_PNG_B64 = (
    "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"
    "YNTjqUSRUwAAAABJRU5ErkJggg=="
)

# ======================================================================
# 1. НАСТРОЙКИ СТРАНИЦЫ И СЦЕНАРИИ
# ======================================================================
# Фирменный стиль Талан для веб-интерфейса (тот же принцип, что в макетах
# экранов генератора посадок, раздел 6.10 ТЗ): белый фон, Талан-зелёный
# #41AA37 — акценты и рамки, тёмно-зелёный #2E7D27 — кнопки и шапки
# (контраст белого текста). Использует те же цвета, что уже зашиты в
# Excel-экспорт (XL_GREEN/XL_GREEN_DARK и логотип TALAN_LOGO_PNG_B64) —
# единый фирменный стиль в вебе и в выгрузке.
TALAN_UI_GREEN = f"#{XL_GREEN}"       # 41AA37
TALAN_UI_GREEN_DARK = "#2E7D27"       # тёмно-зелёный (кнопки, шапки)
_talan_logo_bytes = base64.b64decode(TALAN_LOGO_PNG_B64)
_talan_logo_icon = PILImage.open(io.BytesIO(_talan_logo_bytes))

st.set_page_config(
    page_title="Финмодель девелоперского проекта",
    page_icon=_talan_logo_icon,
    layout="wide",
)

st.markdown(
    f"""
    <style>
    :root {{
        --talan-green: {TALAN_UI_GREEN};
        --talan-green-dark: {TALAN_UI_GREEN_DARK};
    }}
    /* Кнопки (в т.ч. скачивание) — тёмно-зелёный фон, белый текст */
    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {{
        background-color: var(--talan-green-dark);
        color: #FFFFFF;
        border: 1px solid var(--talan-green-dark);
        border-radius: 6px;
    }}
    .stButton > button:hover, .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {{
        background-color: var(--talan-green);
        border-color: var(--talan-green);
        color: #FFFFFF;
    }}
    .stButton > button:disabled, .stDownloadButton > button:disabled {{
        background-color: #CECECE;
        border-color: #CECECE;
        color: #FFFFFF;
    }}
    /* Активная вкладка — акцент Талан-зелёным */
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {{
        border-bottom-color: var(--talan-green) !important;
    }}
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] p {{
        color: var(--talan-green-dark) !important;
        font-weight: 600;
    }}
    /* Значения метрик — акцент тёмно-зелёным */
    [data-testid="stMetricValue"] {{
        color: var(--talan-green-dark);
    }}
    /* Сайдбар — светлый фон с зелёной рамкой (по мотивам макета «Боковая панель») */
    [data-testid="stSidebar"] {{
        background-color: #FAFAFA;
        border-right: 2px solid var(--talan-green);
    }}
    /* Шапка плотнее */
    .block-container {{ padding-top: 3rem; padding-bottom: 2rem; }}
    /* Полоса KPI и карточки — плитки с фирменным светло-зелёным фоном */
    [data-testid="stMetric"] {{
        background-color: #EEF7EC;
        border-radius: 8px;
        padding: 8px 14px;
    }}
    [data-testid="stMetricLabel"] p {{ font-size: 13px; color: #4A4A4A; }}
    [data-testid="stMetricValue"] {{ font-size: 1.45rem; font-weight: 700; }}
    /* Выбор блока / стадии / группы — активная кнопка тёмно-зелёная */
    button[data-testid="stBaseButton-segmented_controlActive"] {{
        background-color: var(--talan-green-dark) !important;
        border-color: var(--talan-green-dark) !important;
        color: #FFFFFF !important;
    }}
    button[data-testid="stBaseButton-segmented_controlActive"] p {{ color: #FFFFFF !important; }}
    /* Прогресс-бар и слайдеры — фирменный зелёный */
    .stProgress > div > div > div > div {{
        background-color: var(--talan-green);
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# Шаблон графиков Plotly в фирменном стиле (шрифт Source Sans 3, как в интерфейсе)
pio.templates["talan"] = go.layout.Template(layout=dict(
    font=dict(family="'Source Sans', 'Source Sans 3', 'Source Sans Pro', sans-serif", size=13, color="#1F1F1F"),
    title=dict(font=dict(size=15, color="#1F1F1F", weight=700), x=0, xanchor="left"),
    colorway=["#41AA37", "#A6A6A6", "#2E7D27", "#84D26D", "#AE4B67", "#B2DAAE"],
    paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
    xaxis=dict(gridcolor="#EFEFEF", linecolor="#BDBDBD", zerolinecolor="#BDBDBD"),
    yaxis=dict(gridcolor="#EFEFEF", linecolor="#BDBDBD", zerolinecolor="#BDBDBD"),
    separators=", ",
))
pio.templates.default = "plotly_white+talan"

TYPE_RESIDENTIAL = "Жилой блок"
TYPE_PARKING = "Наземный/Многоуровневый паркинг"
BLOCK_TYPES = [TYPE_RESIDENTIAL, TYPE_PARKING]

# Стадия проектирования для расчета себестоимости коробки (A-E) — один переключатель
# на весь проект. МП: детальные площади/объемы блока еще не известны, коробка считается
# укрупненной ставкой руб/м2 NSA. ЭП: детальный расчет по 32 статьям методики.
STAGE_MP = "МП (мастер-план)"
STAGE_EP = "ЭП (эскизный проект)"
DESIGN_STAGES = [STAGE_MP, STAGE_EP]
MP_RATE_DEFAULT = 0.0

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
    "PLOT_AREA": "Площадь участка блока, га",
}
PARAM_COLS = list(BASIS_COLUMN.values())  # 10 доп. параметров — новые колонки таблицы блоков

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
    ("D.10.10", "Лифты", "D", "руб/остановку", "ELEVATOR_STOPS", 3800000),
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
    [{"Код": c, "Статья затрат": n, "Группа": g, "Единица измерения": u, "_basis": b, "Ставка, руб/ед.": 0.0}
     for c, n, g, u, b, r in ITEMS]
)

# ----------------------------------------------------------------------
# Наружные работы (код G) и прочие затраты, связанные с СМР (код Z) — из
# той же методики. Считаются ИНДИВИДУАЛЬНО ПО КАЖДОМУ ЖИЛОМУ БЛОКУ: общий
# участок делится на локальные участки под каждым УБ, и у каждого блока
# свои затраты на благоустройство, сети и т.п. Статьи со статусом "на
# данный момент не используем" не включены. G.20.20 «Автостоянки» тоже не
# включена — паркинг уже считается отдельно по блокам (ставки СМР
# подземного/наземного м/м), включение этой статьи задвоило бы затраты.
# Списки ниже — ШАБЛОНЫ ставок по умолчанию для нового блока; фактические
# ставки/кол-во хранятся индивидуально по каждому блоку (session_state).
# ----------------------------------------------------------------------
# Статьи G на площадь участка — количество берется из площади участка
# САМОГО БЛОКА (колонка «Площадь участка блока, га», вкладка «Исходные
# данные», ТЭП по ЭП), редактируется только ставка.
ITEMS_G_AREA = [
    # (код, статья затрат, единица измерения, ставка по умолчанию)
    ("G.10.10", "Очистка площадки", "площадь участка, га", 250_000),
    ("G.10.20", "Разборка и вывоз сооружений", "площадь участка, га", 180_000),
    ("G.10.30", "Земляные работы по подготовке площадки", "площадь участка, га", 650_000),
    ("G.10.50", "Инженерная подготовка территории", "площадь участка, га", 400_000),
    ("G.20.10", "Дороги", "площадь участка, га", 3_500_000),
    ("G.20.30", "Пешеходные дороги", "площадь участка, га", 0),
    ("G.20.40", "МАФ, детские площадки, ограждение", "площадь участка, га", 1_200_000),
    ("G.20.50", "Озеленение", "площадь участка, га", 900_000),
    ("G.40.20", "Наружное освещение", "площадь участка, га", 350_000),
]
DEFAULT_G_AREA_DF = pd.DataFrame(
    [{"Код": c, "Статья затрат": n, "Единица измерения": u, "Ставка, руб/ед.": 0.0} for c, n, u, r in ITEMS_G_AREA]
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
    [{"Код": c, "Статья затрат": n, "Единица измерения": u, "Кол-во": 0.0, "Ставка, руб/ед.": 0.0}
     for c, n, u, q, r in ITEMS_G_LENGTH]
)

ITEMS_Z_PCT = [
    # (код, статья затрат, ставка по умолчанию — доля от (СМР коробки + G))
    ("Z.10.10", "Услуги генподрядчика", 0.06),
    ("Z.20.10", "Непредвиденные расходы по объекту", 0.07),
]
DEFAULT_Z_PCT_DF = pd.DataFrame(
    [{"Код": c, "Статья затрат": n, "Ставка, доля от СМР+G": 0.0} for c, n, r in ITEMS_Z_PCT]
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
    ("Z.50.50", "Оформление объекта (постановление, кадастр, регистрация, межевание)", 0),
]
DEFAULT_Z_FIXED_DF = pd.DataFrame(
    [{"Код": c, "Статья затрат": n, "Сумма, руб": 0.0} for c, n, s in ITEMS_Z_FIXED]
)

# Вкладка 4 «Инфраструктура» — код Z.50.10 «Покупка объекта недвижимости
# (Инфраструктура)». В отличие от остальных статей Z, здесь нет ни каталога
# заранее известных статей, ни ограничения по типу блока: пользователь сам
# вписывает произвольные статьи (благоустройство, плейхаб, парк, дороги и
# т.п.) по КАЖДОМУ урбан-блоку (жилому и паркингу) — строки добавляются
# свободно, без предзаполненных значений. Сумма статей блока = его Z.50.10.
INFRA_ITEM_COLS = ["Статья затрат", "Сумма, руб"]
EMPTY_INFRA_DF = pd.DataFrame(columns=INFRA_ITEM_COLS)
MAX_INFRA_ITEMS_PER_BLOCK = 10

# ----------------------------------------------------------------------
# Полная иерархия статей методики (группы A-Z) — для формы выгрузки Excel
# "Экономика проекта" по образцу заказчика (3-колоночный код: буква/№1/№2).
# Статьи, которых нет в модели (вся группа F, часть D/G/Z), помечены
# source=None — в выгрузке это пустая строка (0), каталог модели
# (ITEMS/ITEMS_G_*/ITEMS_Z_*) при этом НЕ меняется.
# ----------------------------------------------------------------------
REPORT_ROWS = [
    # (уровень, буква, №1, №2, наименование, источник)
    # источник: None (сумма дочерних строк / нет данных в модели) |
    # ("ITEMS", код) | ("G_AREA", код) | ("G_LENGTH", код) | ("Z_PCT", код) | ("Z_FIXED", код)
    ("group", "A", None, None, "ПОДЗЕМНАЯ ЧАСТЬ", None),
    ("subgroup", "A", 10, None, "Фундаменты", None),
    ("leaf", "A", 10, 10, "Типовые фундаменты (ростверк)", ("ITEMS", "A.10.10")),
    ("leaf", "A", 10, 20, "Специализированные работы (сваи)", ("ITEMS", "A.10.20")),
    ("leaf", "A", 10, 30, "Фундаментная плита", ("ITEMS", "A.10.30")),
    ("subgroup", "A", 20, None, "Строительство подземной части", None),
    ("leaf", "A", 20, 10, "Земляные работы", ("ITEMS", "A.20.10")),
    ("leaf", "A", 20, 20, "Конструкции подземной части", ("ITEMS", "A.20.20")),
    ("leaf", "A", None, None, "Кладовые (перегородки, отделка)", ("ITEMS", "A.—")),
    ("group", "B", None, None, "КОНСТРУКЦИИ", None),
    ("subgroup", "B", 10, None, "Несущие конструкции", None),
    ("leaf", "B", 10, 10, "Несущий каркас и плиты перекрытий, крыльцо, терраса", ("ITEMS", "B.10.10")),
    ("leaf", "B", 10, 20, "Несущие конструкции кровли", None),
    ("subgroup", "B", 20, None, "Ограждающие конструкции", None),
    ("leaf", "B", 20, 10, "Наружные стены и фасады, Ограждение балконов", ("ITEMS", "B.20.10")),
    ("leaf", "B", 20, 20, "Заполнение оконных проемов", ("ITEMS", "B.20.20")),
    ("leaf", "B", 20, 30, "Заполнение дверных проемов", ("ITEMS", "B.20.30")),
    ("leaf", "B", 20, 40, "Остекление лоджий и балконов", ("ITEMS", "B.20.40")),
    ("subgroup", "B", 30, None, "Кровля", None),
    ("leaf", "B", 30, 10, "Кровельные покрытия", ("ITEMS", "B.30.10")),
    ("leaf", "B", 30, 20, "Кровельные проемы", None),
    ("group", "C", None, None, "ВНУТРЕННИЕ РАБОТЫ", None),
    ("subgroup", "C", 10, None, "Общестроительные работы", None),
    ("leaf", "C", 10, 10, "Перегородки", ("ITEMS", "C.10.10")),
    ("leaf", "C", 10, 20, "Внутренние двери", None),
    ("leaf", "C", 10, 30, "Фурнитура", None),
    ("subgroup", "C", 20, None, "Лестницы", None),
    ("leaf", "C", 20, 10, "Конструкции лестниц", ("ITEMS", "C.20.10")),
    ("leaf", "C", 20, 20, "Отделка лестниц", ("ITEMS", "C.20.20")),
    ("subgroup", "C", 30, None, "Отделочные работы предчистовые квартир, чистовая МОП", None),
    ("leaf", "C", 30, 10, "Отделка стен", ("ITEMS", "C.30.10")),
    ("leaf", "C", 30, 20, "Отделка полов", ("ITEMS", "C.30.20")),
    ("leaf", "C", 30, 30, "Отделка потолков", ("ITEMS", "C.30.30")),
    ("subgroup", "C", 40, None, "Отделочные работы чистовые (по квартирам, по домам)", None),
    ("leaf", "C", 40, 10, "Отделка стен", None),
    ("leaf", "C", 40, 20, "Отделка полов", None),
    ("leaf", "C", 40, 30, "Отделка потолков", None),
    ("group", "D", None, None, "ВНУТРЕННИЕ ИНЖЕНЕРНЫЕ СИСТЕМЫ", None),
    ("subgroup", "D", 10, None, "Подъемно-транспортное оборудование", None),
    ("leaf", "D", 10, 10, "Лифты", ("ITEMS", "D.10.10")),
    ("leaf", "D", 10, 20, "Эскалаторы", None),
    ("leaf", "D", 10, 90, "Специализированное транспортное оборудование, Мусоропровод", None),
    ("subgroup", "D", 20, None, "Системы ВК, газоснабжения и технологические трубопроводы", None),
    ("leaf", "D", 20, 10, "Сантехническое оборудование", ("ITEMS", "D.20.10")),
    ("leaf", "D", 20, 20, "Водоснабжение", ("ITEMS", "D.20.20")),
    ("leaf", "D", 20, 30, "Хоз.быт. канализация", ("ITEMS", "D.20.30")),
    ("leaf", "D", 20, 40, "Ливневая канализация", ("ITEMS", "D.20.40")),
    ("leaf", "D", 20, 90, "Технологические трубопроводы и газоснабжение", ("ITEMS", "D.20.90")),
    ("subgroup", "D", 30, None, "Отопление, вентиляция и кондиционирования", None),
    ("leaf", "D", 30, 10, "Автономные источники энергоснабжения", None),
    ("leaf", "D", 30, 20, "Котельные, ИТП и оборудование", None),
    ("leaf", "D", 30, 30, "Система отопления", ("ITEMS", "D.30.30")),
    ("leaf", "D", 30, 40, "Системы вентиляции и кондиционирования", ("ITEMS", "D.30.40")),
    ("leaf", "D", 30, 50, "Оборудование", None),
    ("leaf", "D", 30, 60, "КИП, Автоматизация и Диспетчеризация", None),
    ("leaf", "D", 30, 70, "Пуско-наладочные работы систем ОВиК", None),
    ("leaf", "D", 30, 80, "Специализированные системы ОВиК", None),
    ("subgroup", "D", 40, None, "Противопожарные системы", None),
    ("leaf", "D", 40, 10, "Система спринклерного пожаротушения", None),
    ("leaf", "D", 40, 20, "Система пожарного водоснабжения", None),
    ("leaf", "D", 40, 30, "Пожарные шкафы и огнетушители", None),
    ("leaf", "D", 40, 40, "Система АПС (Автоматическая пожарная сигнализация)", ("ITEMS", "D.40.40")),
    ("leaf", "D", 40, 90, "Специализированные системы пожаротушения", None),
    ("subgroup", "D", 50, None, "Электрические сети и оборудование", None),
    ("leaf", "D", 50, 10, "Силовое оборудование и распределительные сети", ("ITEMS", "D.50.10")),
    ("leaf", "D", 50, 20, "Сети освещения", ("ITEMS", "D.50.20")),
    ("leaf", "D", 50, 30, "Слаботочные сети, системы телекоммуникации, интернет, домофон, АСКУЭ", ("ITEMS", "D.50.30")),
    ("leaf", "D", 50, 40, "Установка конечных элементов в МОП", None),
    ("leaf", "D", 50, 50, "Установка конечных элементов в квартирах", None),
    ("leaf", "D", 50, 90, "Другие слаботочные сети и оборудование", None),
    ("group", "E", None, None, "ОБОРУДОВАНИЕ", None),
    ("subgroup", "E", 10, None, "Технологическое оборудование", None),
    ("leaf", "E", 10, 10, "Коммерческое оборудование", None),
    ("leaf", "E", 10, 20, "Оборудование промышленного назначения", None),
    ("leaf", "E", 10, 30, "Транспортное оборудование", None),
    ("leaf", "E", 10, 90, "Другое технологическое оборудование", None),
    ("subgroup", "E", 20, None, "Мебель и аксессуары", None),
    ("leaf", "E", 20, 10, "Встроенная мебель и аксессуары", ("ITEMS", "E.20.10")),
    ("leaf", "E", 20, 20, "Мебель, оборудование и аксессуары", ("ITEMS", "E.20.20")),
    ("group", "F", None, None, "СПЕЦИАЛИЗИРОВАННЫЕ РАБОТЫ", None),
    ("subgroup", "F", 10, None, "Специализированные конструкции и системы", None),
    ("leaf", "F", 10, 10, "Специализированные несущие конструкции", None),
    ("leaf", "F", 10, 20, "Сборные здания и сооружения", None),
    ("leaf", "F", 10, 30, "Системы специального назначения", None),
    ("leaf", "F", 10, 40, "Специализированные технологические системы", None),
    ("leaf", "F", 10, 50, "КИП и автоматика для специализированных систем", None),
    ("subgroup", "F", 20, None, "Специализированные - разборка, удаление и вывоз конструкций", None),
    ("leaf", "F", 20, 10, "Разборка отдельных конструкций существующих зданий", None),
    ("leaf", "F", 20, 20, "Удаление опасных материалов", None),
    ("group", "G", None, None, "НАРУЖНЫЕ РАБОТЫ", None),
    ("subgroup", "G", 10, None, "Подготовка площадки", None),
    ("leaf", "G", 10, 10, "Очистка площадки", ("G_AREA", "G.10.10")),
    ("leaf", "G", 10, 20, "Разборка и вывоз сооружений", ("G_AREA", "G.10.20")),
    ("leaf", "G", 10, 30, "Земляные работы по подготовке площадки", ("G_AREA", "G.10.30")),
    ("leaf", "G", 10, 40, "Удаление, вывоз и утилизация материалов и мусора", None),
    ("leaf", "G", 10, 50, "Инженерная подготовка территории", ("G_AREA", "G.10.50")),
    ("subgroup", "G", 20, None, "Благоустройство и озеленение", None),
    ("leaf", "G", 20, 10, "Дороги", ("G_AREA", "G.20.10")),
    ("leaf", "G", 20, 20, "Автостоянки", None),
    ("leaf", "G", 20, 30, "Пешеходные дороги", ("G_AREA", "G.20.30")),
    ("leaf", "G", 20, 40, "МАФ, оборудование детских площадок, постоянный забор", ("G_AREA", "G.20.40")),
    ("leaf", "G", 20, 50, "Озеленение", ("G_AREA", "G.20.50")),
    ("subgroup", "G", 30, None, "Наружные трубопроводы", None),
    ("leaf", "G", 30, 10, "Водоснабжение", ("G_LENGTH", "G.30.10")),
    ("leaf", "G", 30, 20, "Хоз.быт. канализация", ("G_LENGTH", "G.30.20")),
    ("leaf", "G", 30, 30, "Ливневая канализация", ("G_LENGTH", "G.30.30")),
    ("leaf", "G", 30, 40, "Сети теплоснабжения", ("G_LENGTH", "G.30.40")),
    ("leaf", "G", 30, 50, "Сети холодоснабжения", ("G_LENGTH", "G.30.50")),
    ("leaf", "G", 30, 60, "Топливоснабжение", ("G_LENGTH", "G.30.60")),
    ("leaf", "G", 30, 90, "Технологические трубопроводы", None),
    ("subgroup", "G", 40, None, "Наружные электротехнические сети", None),
    ("leaf", "G", 40, 10, "Сети электроснабжения, трансформаторная подстанция", ("G_LENGTH", "G.40.10")),
    ("leaf", "G", 40, 20, "Наружное освещение", ("G_AREA", "G.40.20")),
    ("leaf", "G", 40, 30, "Наружные слаботочные сети", None),
    ("leaf", "G", 40, 90, "Системы аварийного электроснабжения", None),
    ("subgroup", "G", 90, None, "Специальные наружные работы", None),
    ("leaf", "G", 90, 10, "Технологические каналы и пешеходные тоннели", None),
    ("leaf", "G", 90, 20, "Системы снеготаяния", None),
    ("group", "Z", None, None, "ПРОЧИЕ ЗАТРАТЫ, связанные с СМР", None),
    ("subgroup", "Z", 10, None, "Услуги, временные здания и работы", None),
    ("leaf", "Z", 10, 10, "Услуги ген.подрядчика", ("Z_PCT", "Z.10.10")),
    ("leaf", "Z", 10, 20, "Контроль качества выполнения работ", ("Z_FIXED", "Z.10.20")),
    ("leaf", "Z", 10, 30, "Строительство временных зданий и сооружений", ("Z_FIXED", "Z.10.30")),
    ("leaf", "Z", 10, 40, "Коммунальные услуги по объекту в период строительства", ("Z_FIXED", "Z.10.40")),
    ("leaf", "Z", 10, 50, "Услуги сторонних организаций, связанные с СМР, разрешения, согласования, сдача объекта", ("Z_FIXED", "Z.10.50")),
    ("leaf", "Z", 10, 60, "Проектирование, изыскания, авторский надзор", ("Z_FIXED", "Z.10.60")),
    ("subgroup", "Z", 20, None, "Непредвиденные расходы по объекту", None),
    ("leaf", "Z", 20, 10, "Непредвиденные расходы по объекту", ("Z_PCT", "Z.20.10")),
    ("subgroup", "Z", 50, None, "Расходы по предпроектной стадии", None),
    ("leaf", "Z", 50, 10, "Покупка объекта недвижимости (Инфраструктура)", ("INFRA_Z5010", None)),
    ("leaf", "Z", 50, 20, "Покупка объекта недвижимости (земля, недострой)", None),
    ("leaf", "Z", 50, 30, "Арендные платежи и расходы по регистрации договора аренды зем.участка", None),
    ("leaf", "Z", 50, 40, "Отселение/снос строений", None),
    ("leaf", "Z", 50, 50, "Оформление объекта (постановление, кадастр, регистрация, межевание)", ("Z_FIXED", "Z.50.50")),
    ("subgroup", "Z", 150, None, "Технологическое присоединение", ("Z_FIXED", "Z.150")),
    ("leaf", "Z", 150, 10, "Технологическое присоединение водопровод", None),
    ("leaf", "Z", 150, 20, "Технологическое присоединение хозбытовая канализация", None),
    ("leaf", "Z", 150, 30, "Технологическое присоединение теплоснабжение", None),
    ("leaf", "Z", 150, 40, "Технологическое присоединение электроснабжение", None),
    ("leaf", "Z", 150, 50, "Технологическое присоединение газоснабжение", None),
]


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
    "Площадь участка блока, га": "Площадь участка блока, га",  # ЭП.9 — локальный участок под блоком
}
TEP_COLS = list(TEP_TO_MAIN_COL.keys())

# Из 9 полей ТЭП четыре нужны ВСЕГДА (площади квартир/коммерции/кладовых —
# NSA и выручка, площадь участка блока — расчет G): показываются в блоке ЭП
# на Вкладке 1 независимо от стадии. Остальные 5 — только база каталога A-E
# (кол-во подъездов/квартир/лифтов, площадь застройки, остекление окон) —
# видны только на этапе ЭП, как и EXTRA_EP_PARAM_COLS ниже.
ALWAYS_VISIBLE_TEP_COLS = [
    "Общая площадь квартир (с летними, с коэф.), м2", "Площадь коммерции, м2",
    "Площадь кладовых в доме, м2", "Площадь участка блока, га",
]

# Доп. параметры блока, нужные ТОЛЬКО для ЭП-каталога A-E (базы VOL_BELOW,
# VOL_ABOVE, AREA_FACADE, AREA_GLAZING_BALCONY) — не входят в методику ТЭП
# (TEP_TO_MAIN_COL), поэтому редактируются напрямую в таблице блоков, без
# промежуточного tep_store. Показываются вместе с блоком ТЭП на Вкладке 1.
EXTRA_EP_PARAM_COLS = [
    "Объем здания ниже 0, м3", "Объем здания выше 0, м3",
    "Площадь фасада, м2", "Площадь остекления лоджий, м2",
]

# Цены продаж — задаются глобально (Блок А) с точечным переопределением по
# блоку (Блок Б) на Вкладке 2.
PRICE_COLS = [
    "Цена жилья, руб/м2", "Цена коммерции, руб/м2", "Цена кладовых, руб/м2",
    "Цена подземного м/м, руб", "Цена наземного м/м, руб",
]
# Ставки СМР паркинга — тоже глобально с переопределением, но на Вкладке 3
# (нужны в обоих режимах МП/ЭП — паркинг считается вне каталога A-E).
SMR_PARKING_RATE_COLS = [
    "Ставка СМР подземного м/м, руб", "Ставка СМР наземного м/м, руб",
    "Ставка СМР коммерции паркинга, руб/м2", "Ставка СМР кладовых паркинга, руб/м2",
]


# ======================================================================
# 3. КОЛОНКИ ЕДИНОЙ ТАБЛИЦЫ ТЭП И ТЕСТОВЫЕ ДАННЫЕ
# ======================================================================
NUMERIC_COLS = [
    "S квартир, м2", "S коммерции 1 эт., м2", "S кладовых, м2",
    "Подземный паркинг, м/м", "Наземный/Многоуровневый паркинг, м/м",
    "Цена жилья, руб/м2", "Цена коммерции, руб/м2", "Цена кладовых, руб/м2",
    "Цена подземного м/м, руб", "Цена наземного м/м, руб",
    "Ставка СМР подземного м/м, руб", "Ставка СМР наземного м/м, руб",
    "Ставка СМР коммерции паркинга, руб/м2", "Ставка СМР кладовых паркинга, руб/м2",
] + PARAM_COLS  # + 9 доп. параметров для расчета СМР коробки по методике
ALL_COLS = ["Название блока", "Тип блока"] + NUMERIC_COLS

# Вкладка 1 «Объемы и ТЭП этапов» — единственная таблица с динамическими
# строками (добавление/удаление урбан-блоков). Только объемные показатели —
# цены (PRICE_COLS), ставки СМР паркинга (SMR_PARKING_RATE_COLS) и доп.
# параметры ЭП (PARAM_COLS) вводятся на Вкладках 2/3 и хранятся отдельно
# (cascade-таблицы / tep_store), а не как колонки этой таблицы.
MAIN_TABLE_NUMERIC_COLS = [
    "S квартир, м2", "S коммерции 1 эт., м2", "S кладовых, м2",
    "Подземный паркинг, м/м", "Наземный/Многоуровневый паркинг, м/м",
]
MAIN_TABLE_COLS = ["Название блока", "Тип блока"] + MAIN_TABLE_NUMERIC_COLS


def generate_default_table() -> pd.DataFrame:
    """Пустой шаблон: один жилой урбан-блок («УБ 1»), готовый для заполнения.
    Новые блоки добавляются через таблицу (+)."""
    row = {col: 0.0 for col in MAIN_TABLE_NUMERIC_COLS}
    row.update({"Название блока": "УБ 1", "Тип блока": TYPE_RESIDENTIAL})
    return pd.DataFrame([row])[MAIN_TABLE_COLS]


# ======================================================================
# 3.5. СОХРАНЕНИЕ ПО ПОЛЬЗОВАТЕЛЮ И ПРОЕКТУ (переживает перезапуск страницы)
# ======================================================================
# Каждое сохранение — отдельный файл, ключ = имя пользователя + название
# проекта. Так несколько человек могут пользоваться моделью, не перезаписывая
# данные друг друга, и держать несколько проектов одновременно.
#
# Хранилище:
#   • Яндекс Диск — если в Streamlit Secrets задан YANDEX_DISK_TOKEN (OAuth-токен
#     приложения с правом «Доступ к папке приложения на Диске»). Файлы лежат в
#     «Приложения / <имя приложения> / projects». Запись идет в фоне (не тормозит
#     интерфейс), только при изменении данных.
#   • Локальная папка рядом с app.py — если токена нет (запуск на своем ПК).
#     На Streamlit Cloud локальный диск временный — данные теряются при
#     перезапуске, поэтому там нужен Яндекс Диск.
SAVES_DIR = Path(__file__).resolve().parent / "talan_model_saves"
POINTER_PATH = Path(__file__).resolve().parent / "talan_model_last_opened.json"
YA_API = "https://cloud-api.yandex.net/v1/disk/resources"
YA_FOLDER = "app:/projects"
YA_TIMEOUT = 20  # сек на один HTTP-запрос
YA_LIST_TTL = 20  # сек — кэш списка проектов в сессии

# Одиночные DataFrame
_SAVE_DF_KEYS = ["blocks_df"]
# Словари {имя блока -> DataFrame}
_SAVE_DICT_OF_DF_KEYS = [
    "block_rates", "block_g_area", "block_g_length", "block_z_pct", "block_z_fixed",
    "block_infra",
]
# Простые JSON-совместимые значения (строки/словари чисел/словарь словарей)
_SAVE_PLAIN_KEYS = [
    "tep_store", "block_mp_rate", "block_elevator_stops", "design_stage",
    "project_name_input", "project_city_input", "scenario_select",
]
# Динамически именуемые ключи (cascade-хранилища и значения "Блок А" по умолчанию)
_SAVE_PREFIXES = ("_cascade_", "_price_default_", "_smr_parking_default_", "_mp_korobka_default")


def _slugify(s: str) -> str:
    """Имя пользователя/проекта -> безопасное имя файла."""
    s = (s or "").strip()
    s = re.sub(r"[^\w\-]+", "_", s, flags=re.UNICODE)
    return s.strip("_") or "без_имени"


def _json_default(o):
    """numpy-числа и прочее -> JSON."""
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating,)):
        return float(o)
    if isinstance(o, np.ndarray):
        return o.tolist()
    return str(o)


def _ya_token():
    try:
        tok = st.secrets.get("YANDEX_DISK_TOKEN")
    except Exception:
        return None
    return str(tok).strip() if tok else None


YA_TOKEN = _ya_token()
STORAGE_MODE = "yadisk" if YA_TOKEN else "local"


# ---------------------------- Яндекс Диск ------------------------------
def _ya_headers():
    return {"Authorization": f"OAuth {YA_TOKEN}"}


def _ya_check(resp, ok=(200, 201, 202, 204)):
    if resp.status_code not in ok:
        try:
            msg = resp.json().get("message") or resp.text
        except Exception:
            msg = resp.text
        raise RuntimeError(f"Яндекс Диск: HTTP {resp.status_code} — {msg}")
    return resp


def _ya_ensure_folder():
    r = requests.put(YA_API, params={"path": YA_FOLDER}, headers=_ya_headers(), timeout=YA_TIMEOUT)
    _ya_check(r, ok=(201, 409))  # 409 — папка уже есть


def _ya_upload(name: str, text: str, meta: dict):
    path = f"{YA_FOLDER}/{name}"
    r = requests.get(
        f"{YA_API}/upload", params={"path": path, "overwrite": "true"},
        headers=_ya_headers(), timeout=YA_TIMEOUT,
    )
    if r.status_code == 409:  # нет папки — создаем и повторяем
        _ya_ensure_folder()
        r = requests.get(
            f"{YA_API}/upload", params={"path": path, "overwrite": "true"},
            headers=_ya_headers(), timeout=YA_TIMEOUT,
        )
    href = _ya_check(r).json()["href"]
    _ya_check(requests.put(href, data=text.encode("utf-8"), timeout=YA_TIMEOUT * 3))
    # имя пользователя/проекта кладем в свойства файла — список проектов
    # строится одним запросом, без скачивания каждого файла
    _ya_check(requests.patch(
        YA_API, params={"path": path}, headers=_ya_headers(),
        json={"custom_properties": {k: str(v) for k, v in meta.items()}}, timeout=YA_TIMEOUT,
    ))


def _ya_download(name: str) -> str:
    r = requests.get(
        f"{YA_API}/download", params={"path": f"{YA_FOLDER}/{name}"},
        headers=_ya_headers(), timeout=YA_TIMEOUT,
    )
    href = _ya_check(r).json()["href"]
    r2 = _ya_check(requests.get(href, timeout=YA_TIMEOUT * 3))
    return r2.content.decode("utf-8")


def _ya_list() -> list:
    r = requests.get(
        YA_API,
        params={
            "path": YA_FOLDER, "limit": 1000,
            "fields": "_embedded.items.name,_embedded.items.type,_embedded.items.modified,"
                      "_embedded.items.custom_properties",
        },
        headers=_ya_headers(), timeout=YA_TIMEOUT,
    )
    if r.status_code == 404:  # папки еще нет — сохранений нет
        return []
    items = []
    for it in _ya_check(r).json().get("_embedded", {}).get("items", []):
        if it.get("type") != "file" or not it.get("name", "").endswith(".json"):
            continue
        props = it.get("custom_properties") or {}
        if props.get("user_name") and props.get("project_name"):
            items.append({
                "user_name": props["user_name"],
                "project_name": props["project_name"],
                "saved_at": props.get("saved_at") or it.get("modified", ""),
                "ref": it["name"],
            })
    return items


def _ya_delete(name: str):
    r = requests.delete(
        YA_API, params={"path": f"{YA_FOLDER}/{name}", "permanently": "true"},
        headers=_ya_headers(), timeout=YA_TIMEOUT,
    )
    _ya_check(r, ok=(202, 204, 404))


class _YaWriter:
    """Фоновая запись на Яндекс Диск: интерфейс не ждет сети. По каждому файлу
    хранится только последняя версия (промежуточные правки не отправляются)."""

    def __init__(self):
        self.cond = threading.Condition()
        self.pending = {}  # имя файла -> (текст, meta)
        self.in_flight = {}  # имя файла -> текст, который пишется прямо сейчас
        self.errors = {}  # имя файла -> текст последней ошибки
        self.last_ok = {}  # имя файла -> время последней успешной записи
        threading.Thread(target=self._run, daemon=True).start()

    def submit(self, name, text, meta):
        with self.cond:
            self.pending[name] = (text, meta)
            self.cond.notify()

    def latest_text(self, name):
        """Самая свежая версия, еще не дошедшая до Диска (для чтения своих правок)."""
        with self.cond:
            if name in self.pending:
                return self.pending[name][0]
            return self.in_flight.get(name)

    def forget(self, name):
        with self.cond:
            self.pending.pop(name, None)
            self.errors.pop(name, None)

    def busy(self):
        with self.cond:
            return bool(self.pending or self.in_flight)

    def _run(self):
        while True:
            with self.cond:
                while not self.pending:
                    self.cond.wait()
                name, (text, meta) = next(iter(self.pending.items()))
                del self.pending[name]
                self.in_flight[name] = text
            try:
                _ya_upload(name, text, meta)
                with self.cond:
                    self.errors.pop(name, None)
                    self.last_ok[name] = datetime.now().strftime("%H:%M:%S")
            except Exception as e:  # сеть/токен — сохраняем ошибку, повтор при следующей правке
                with self.cond:
                    self.errors[name] = str(e)[:300]
            finally:
                with self.cond:
                    self.in_flight.pop(name, None)


@st.cache_resource
def _get_ya_writer():
    return _YaWriter()


# ------------------------ общий интерфейс хранилища ---------------------
def save_path_for(user_name: str, project_name: str) -> str:
    """Имя файла сохранения (одинаково для Яндекс Диска и локальной папки)."""
    return f"{_slugify(user_name)}__{_slugify(project_name)}.json"


def current_user_and_project() -> tuple:
    return (
        st.session_state.get("user_name_input", "").strip(),
        st.session_state.get("project_name_input", "").strip(),
    )


def current_save_path():
    user_name, project_name = current_user_and_project()
    if not user_name or not project_name:
        return None
    return save_path_for(user_name, project_name)


def _invalidate_list_cache():
    st.session_state.pop("_saves_cache", None)


def list_saved_projects() -> list:
    """Список всех сохранений (метаданные), новые сверху."""
    if STORAGE_MODE == "yadisk":
        cache = st.session_state.get("_saves_cache")
        if cache and (datetime.now().timestamp() - cache[0]) < YA_LIST_TTL:
            return cache[1]
        try:
            items = _ya_list()
            st.session_state.pop("_storage_error", None)
        except Exception as e:
            st.session_state["_storage_error"] = str(e)[:300]
            items = cache[1] if cache else []
        items.sort(key=lambda x: x["saved_at"], reverse=True)
        st.session_state["_saves_cache"] = (datetime.now().timestamp(), items)
        return items

    if not SAVES_DIR.exists():
        return []
    items = []
    for f in SAVES_DIR.glob("*.json"):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            meta = data.get("_meta", {})
            if meta.get("user_name") and meta.get("project_name"):
                items.append({
                    "user_name": meta["user_name"],
                    "project_name": meta["project_name"],
                    "saved_at": meta.get("saved_at", ""),
                    "ref": f.name,
                })
        except Exception:
            continue
    items.sort(key=lambda x: x["saved_at"], reverse=True)
    return items


def storage_delete(ref: str) -> None:
    if STORAGE_MODE == "yadisk":
        _get_ya_writer().forget(ref)
        try:
            _ya_delete(ref)
        except Exception as e:
            st.session_state["_storage_error"] = str(e)[:300]
        _invalidate_list_cache()
        return
    try:
        (SAVES_DIR / ref).unlink(missing_ok=True)
    except Exception:
        pass


def _autosave_collect() -> dict:
    """Собирает весь пользовательский ввод из session_state в JSON-совместимый словарь."""
    data = {"df": {}, "dict_of_df": {}, "plain": {}, "prefixed": {}, "indirect_items": {}}
    for key in _SAVE_DF_KEYS:
        if key in st.session_state:
            data["df"][key] = st.session_state[key].to_dict(orient="records")
    for key in _SAVE_DICT_OF_DF_KEYS:
        if key in st.session_state:
            data["dict_of_df"][key] = {
                name: df.to_dict(orient="records") for name, df in st.session_state[key].items()
            }
    for key in _SAVE_PLAIN_KEYS:
        if key in st.session_state:
            data["plain"][key] = st.session_state[key]
    if "indirect_items" in st.session_state:
        data["indirect_items"] = {
            name: df.to_dict(orient="records") for name, df in st.session_state.indirect_items.items()
        }
    for key in list(st.session_state.keys()):
        if isinstance(key, str) and key.startswith(_SAVE_PREFIXES):
            val = st.session_state[key]
            data["prefixed"][key] = val
    return data


def autosave_write(ref: str, user_name: str, project_name: str) -> None:
    """Сохраняет текущее состояние, только если оно изменилось с прошлой записи.
    Никогда не роняет приложение при ошибке."""
    try:
        data = _autosave_collect()
        body = json.dumps(data, ensure_ascii=False, default=_json_default, sort_keys=True)
        sig = (ref, user_name, project_name, hash(body))
        if st.session_state.get("_last_saved_sig") == sig:
            return
        meta = {
            "user_name": user_name,
            "project_name": project_name,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        }
        data["_meta"] = meta
        text = json.dumps(data, ensure_ascii=False, default=_json_default)
        if STORAGE_MODE == "yadisk":
            _get_ya_writer().submit(ref, text, meta)
            cache = st.session_state.get("_saves_cache")
            if cache and not any(s["ref"] == ref for s in cache[1]):
                _invalidate_list_cache()  # новый проект — обновим список
        else:
            SAVES_DIR.mkdir(parents=True, exist_ok=True)
            (SAVES_DIR / ref).write_text(text, encoding="utf-8")
            POINTER_PATH.write_text(
                json.dumps({"user_name": user_name, "project_name": project_name}, ensure_ascii=False),
                encoding="utf-8",
            )
        st.session_state["_last_saved_sig"] = sig
    except Exception:
        pass


def _storage_read(ref: str):
    if STORAGE_MODE == "yadisk":
        text = _get_ya_writer().latest_text(ref)
        if text is not None:
            return text
        try:
            return _ya_download(ref)
        except Exception as e:
            st.session_state["_storage_error"] = str(e)[:300]
            return None
    p = SAVES_DIR / ref
    return p.read_text(encoding="utf-8") if p.exists() else None


def autosave_load(ref: str) -> bool:
    """Загружает сохраненное состояние в session_state ДО отрисовки виджетов.
    Возвращает True, если что-то было восстановлено."""
    if ref is None:
        return False
    try:
        text = _storage_read(ref)
        if text is None:
            return False
        data = json.loads(text)
    except Exception:
        return False
    try:
        for key, records in data.get("df", {}).items():
            if records:
                st.session_state[key] = pd.DataFrame(records)
        for key, blocks in data.get("dict_of_df", {}).items():
            st.session_state[key] = {
                name: pd.DataFrame(records) if records else pd.DataFrame()
                for name, records in blocks.items()
            }
        for key, val in data.get("plain", {}).items():
            st.session_state[key] = val
        indirect = data.get("indirect_items", {})
        if indirect:
            st.session_state.indirect_items = {
                name: pd.DataFrame(records) if records else pd.DataFrame()
                for name, records in indirect.items()
            }
        for key, val in data.get("prefixed", {}).items():
            st.session_state[key] = val
        meta = data.get("_meta", {})
        if meta.get("user_name"):
            st.session_state["user_name_input"] = meta["user_name"]
        if meta.get("project_name"):
            st.session_state["project_name_input"] = meta["project_name"]
        return True
    except Exception:
        return False


def _read_pointer():
    try:
        return json.loads(POINTER_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None


# При первом запуске сессии (только локальный режим, один человек за ПК) —
# подхватываем последний открытый проект. На общем Яндекс Диске так не делаем:
# коллега не должен открыть чужой проект — выбор только в боковой панели.
if "_bootstrapped" not in st.session_state:
    st.session_state["_bootstrapped"] = True
    st.session_state["_autosave_loaded"] = False
    if STORAGE_MODE == "local":
        pointer = _read_pointer()
        if pointer and pointer.get("user_name") and pointer.get("project_name"):
            p = save_path_for(pointer["user_name"], pointer["project_name"])
            if autosave_load(p):
                st.session_state["_autosave_loaded"] = True

if "user_name_input" not in st.session_state:
    st.session_state["user_name_input"] = ""
if "design_stage" not in st.session_state:
    st.session_state.design_stage = STAGE_EP  # общий на весь проект переключатель МП/ЭП для расчета коробки (A-E)
if "project_name_input" not in st.session_state:
    st.session_state["project_name_input"] = "ЖК «Пример»"
if "project_city_input" not in st.session_state:
    st.session_state["project_city_input"] = "Самара"
if "scenario_select" not in st.session_state or st.session_state["scenario_select"] not in SCENARIOS:
    st.session_state["scenario_select"] = list(SCENARIOS.keys())[0]

MAX_INDIRECT_ITEMS = 10


def cascade_sync(store_key: str, block_names: list, col_defaults: dict) -> dict:
    """Глобальное значение по умолчанию (Блок А) -> точечно редактируемая таблица
    по блокам (Блок Б). Ячейка обновляется новым дефолтом, только если она еще
    равна ПРЕДЫДУЩЕМУ дефолту (т.е. пользователь ее вручную не менял). Возвращает
    dict {имя блока -> {колонка: значение}} — актуальный на блоки block_names."""
    store = st.session_state.setdefault(f"_cascade_{store_key}", {})
    prev_defaults = st.session_state.setdefault(f"_cascade_{store_key}_prev", {})
    for name in block_names:
        row = store.setdefault(name, {})
        for col, default_val in col_defaults.items():
            old_default = prev_defaults.get(col)
            if col not in row or row[col] == old_default:
                row[col] = default_val
    for name in list(store.keys()):
        if name not in block_names:
            del store[name]
    prev_defaults.update(col_defaults)
    return store


BASE_ROW_LABEL = "База — все блоки"
RESET_COL = "↺ к базе"
CHANGED_COL = "Изменено вручную"


def _num(x) -> float:
    """Число из ячейки редактора (None/NaN/пусто -> 0)."""
    v = pd.to_numeric(x, errors="coerce")
    return 0.0 if pd.isna(v) else float(v)


def fmt_money(v: float) -> str:
    return f"{v:,.0f}".replace(",", " ")


def fmt_mln(v: float, unit: bool = True) -> str:
    s = f"{v / 1e6:,.1f}".replace(",", " ").replace(".", ",")
    return f"{s} млн ₽" if unit else s


def _bump(key: str) -> None:
    st.session_state[f"_ver_{key}"] = st.session_state.get(f"_ver_{key}", 0) + 1


def fixed_selector(label: str, options: list, key: str, **kw):
    """Выбор блока кнопками. Значение всегда из списка (при удалении/переименовании
    блока или снятии выбора — первый блок)."""
    if st.session_state.get(key) not in options:
        st.session_state[key] = options[0]
    val = st.segmented_control(label, options, key=key, **kw)
    return val if val in options else options[0]


def base_row_table(store_key: str, cols: list, default_keys: dict, block_names: list,
                   editor_key: str, labels: dict, formats: dict = None) -> dict:
    """Таблица «База + блоки»: первая строка задает значение всем блокам, ячейка
    блока, измененная вручную, сохраняет свое значение при смене базы, «↺» —
    вернуть строку блока к базе. Хранилище — прежнее cascade-хранилище
    (cascade_sync), базовые значения — прежние ключи session_state."""
    formats = formats or {}
    defaults = {c: _num(st.session_state.get(default_keys[c], 0.0)) for c in cols}
    for c in cols:
        st.session_state[default_keys[c]] = defaults[c]
    store = cascade_sync(store_key, block_names, defaults)
    for n in block_names:
        for c in cols:
            store[n][c] = _num(store[n].get(c, 0.0))
    rows = [{"Блок": BASE_ROW_LABEL, **defaults, RESET_COL: False, CHANGED_COL: ""}]
    for n in block_names:
        changed = [labels[c] for c in cols if store[n][c] != defaults[c]]
        rows.append({"Блок": n, **{c: store[n][c] for c in cols}, RESET_COL: False, CHANGED_COL: ", ".join(changed)})
    df = pd.DataFrame(rows)
    cfg = {
        "Блок": st.column_config.TextColumn(disabled=True),
        RESET_COL: st.column_config.CheckboxColumn(RESET_COL, help="Вернуть строку блока к базе", width="small"),
        CHANGED_COL: st.column_config.TextColumn(disabled=True),
    }
    for c in cols:
        cfg[c] = st.column_config.NumberColumn(labels[c], min_value=0, format=formats.get(c, "localized"))
    ver = st.session_state.get(f"_ver_{editor_key}", 0)
    ed = st.data_editor(
        df, key=f"{editor_key}_{ver}", hide_index=True, num_rows="fixed", width="stretch",
        column_order=["Блок"] + cols + [RESET_COL, CHANGED_COL], column_config=cfg,
    )
    new_defaults = {c: _num(ed.iloc[0][c]) for c in cols}
    changed_any = False
    for i in range(1, len(ed)):
        n = ed.iloc[i]["Блок"]
        if bool(ed.iloc[i][RESET_COL]):
            for c in cols:
                store[n][c] = new_defaults[c]
            changed_any = True
            continue
        for c in cols:
            v = _num(ed.iloc[i][c])
            if v != store[n][c]:
                store[n][c] = v
                changed_any = True
    if new_defaults != defaults:
        for c in cols:
            st.session_state[default_keys[c]] = new_defaults[c]
        store = cascade_sync(store_key, block_names, new_defaults)
        changed_any = True
    if changed_any:
        _bump(editor_key)
        st.rerun()
    return store


def transposed_editor(params: list, names: list, get, set_, key: str, allowed) -> None:
    """Развернутая таблица: строки — параметры, столбцы — блоки."""
    rows = []
    for k, label, unit in params:
        r = {"Параметр": label, "Ед.": unit}
        for n in names:
            r[n] = get(n, k) if allowed(n, k) else None
        rows.append(r)
    df = pd.DataFrame(rows)
    cfg = {
        "Параметр": st.column_config.TextColumn(disabled=True, width="medium"),
        "Ед.": st.column_config.TextColumn(disabled=True, width="small"),
    }
    for n in names:
        cfg[n] = st.column_config.NumberColumn(n, min_value=0, format="localized")
    ver = st.session_state.get(f"_ver_{key}", 0)
    ed = st.data_editor(
        df, key=f"{key}_{ver}_{abs(hash(tuple(names)))}", hide_index=True, num_rows="fixed",
        width="stretch", column_config=cfg,
    )
    changed = False
    for i, (k, _label, _unit) in enumerate(params):
        for n in names:
            if not allowed(n, k):
                if not pd.isna(pd.to_numeric(ed.iloc[i][n], errors="coerce")):
                    changed = True  # в закрытую ячейку что-то ввели — сбрасываем
                continue
            v = _num(ed.iloc[i][n])
            if v != get(n, k):
                set_(n, k, v)
                changed = True
    if changed:
        _bump(key)
        st.rerun()


def render_indirect_category(state_key: str, label: str, default_amount: float) -> float:
    """Раздел пула косвенных расходов: заголовок с суммой + список статей
    (Наименование, Сумма) с добавлением строк «+» — до MAX_INDIRECT_ITEMS."""
    if state_key not in st.session_state.indirect_items or "Сумма, руб" not in st.session_state.indirect_items[state_key].columns:
        st.session_state.indirect_items[state_key] = pd.DataFrame(
            [{"Наименование": label, "Сумма, руб": default_amount}]
        )
    with st.container(border=True):
        hdr = st.empty()
        edited = st.data_editor(
            st.session_state.indirect_items[state_key],
            width="stretch",
            num_rows="dynamic",
            hide_index=True,
            key=f"indirect_editor_{state_key}",
            column_config={
                "Наименование": st.column_config.TextColumn("Статья"),
                "Сумма, руб": st.column_config.NumberColumn(min_value=0, format="localized"),
            },
        )
        if len(edited) > MAX_INDIRECT_ITEMS:
            st.warning(f"Максимум {MAX_INDIRECT_ITEMS} статей в разделе — лишние строки не учитываются.")
            edited = edited.iloc[:MAX_INDIRECT_ITEMS].reset_index(drop=True)
        st.session_state.indirect_items[state_key] = edited
        total = float(pd.to_numeric(edited["Сумма, руб"], errors="coerce").fillna(0.0).sum())
        hdr.markdown(f"**{label}** — {fmt_money(total)} ₽ · строка «+» внизу — новая статья")
    return total


# ======================================================================
# 3.7. ШАБЛОН EXCEL: ЗАГРУЗКА ДАННЫХ ПРОЕКТА И ВЫГРУЗКА В ШАБЛОН
# ======================================================================
# Один файл .xlsx заполняет все вкладки модели. Листы:
#   «Проект» — наименование, город, стадия, сценарий;
#   «Блоки» — строка на урбан-блок: тип, площади, м/м, участок, цены, ставки
#             МП и паркинга, параметры ЭП;
#   «Статьи СМР» — ставки/суммы по кодам методики (A–E, G, Z) по блокам;
#   «Инфраструктура» — статьи Z.50.10 по блокам;
#   «Косвенные» — пул косвенных расходов проекта по разделам;
#   «Справочник кодов» — какие коды принимаются и что в них вводить.
TPL_SHEET_PROJECT = "Проект"
TPL_SHEET_BLOCKS = "Блоки"
TPL_SHEET_ITEMS = "Статьи СМР"
TPL_SHEET_INFRA = "Инфраструктура"
TPL_SHEET_INDIRECT = "Косвенные"
TPL_SHEET_CODES = "Справочник кодов"

TPL_TYPE_ALIASES = {
    "жилой": TYPE_RESIDENTIAL, "жилой блок": TYPE_RESIDENTIAL,
    "паркинг": TYPE_PARKING, "наземный паркинг": TYPE_PARKING, "многоуровневый паркинг": TYPE_PARKING,
    TYPE_PARKING.lower(): TYPE_PARKING,
}
# (заголовок колонки листа «Блоки», куда пишется, ключ)
TPL_BLOCK_COLS = [
    ("Название блока", "name", None),
    ("Тип блока", "type", None),
    ("Площадь квартир, м2", "tep", "Общая площадь квартир (с летними, с коэф.), м2"),
    ("Площадь коммерции, м2", "tep", "Площадь коммерции, м2"),
    ("Площадь кладовых, м2", "tep", "Площадь кладовых в доме, м2"),
    ("Подземный паркинг, м/м", "main", "Подземный паркинг, м/м"),
    ("Наземный/многоуровневый паркинг, м/м", "main", "Наземный/Многоуровневый паркинг, м/м"),
    ("Площадь участка блока, га", "tep", "Площадь участка блока, га"),
    ("Цена жилья, руб/м2", "price", "Цена жилья, руб/м2"),
    ("Цена коммерции, руб/м2", "price", "Цена коммерции, руб/м2"),
    ("Цена кладовых, руб/м2", "price", "Цена кладовых, руб/м2"),
    ("Цена подземного м/м, руб", "price", "Цена подземного м/м, руб"),
    ("Цена наземного м/м, руб", "price", "Цена наземного м/м, руб"),
    ("Ставка коробки МП, руб/м2 NSA", "mp", "Ставка, руб/м2 NSA"),
    ("Ставка СМР подземного м/м, руб", "park", "Ставка СМР подземного м/м, руб"),
    ("Ставка СМР наземного м/м, руб", "park", "Ставка СМР наземного м/м, руб"),
    ("Ставка СМР коммерции паркинга, руб/м2", "park", "Ставка СМР коммерции паркинга, руб/м2"),
    ("Ставка СМР кладовых паркинга, руб/м2", "park", "Ставка СМР кладовых паркинга, руб/м2"),
    ("ЭП: кол-во секций, шт", "tep", "Кол-во секций, шт"),
    ("ЭП: кол-во квартир, шт", "tep", "Кол-во квартир, шт"),
    ("ЭП: кол-во лифтов, шт", "tep", "Кол-во лифтов, шт"),
    ("ЭП: остановок лифтов, шт", "stops", None),
    ("ЭП: площадь застройки, м2", "tep", "Площадь застройки, м2"),
    ("ЭП: площадь фасада, м2", "tep", "Площадь фасада, м2"),
    ("ЭП: остекление окон, м2", "tep", "Площадь остекления окон, м2"),
    ("ЭП: остекление лоджий, м2", "tep", "Площадь остекления лоджий, м2"),
    ("ЭП: объем ниже 0, м3", "tep", "Объем здания ниже 0, м3"),
    ("ЭП: объем выше 0, м3", "tep", "Объем здания выше 0, м3"),
]
TPL_ITEM_COLS = ["Блок", "Код", "Статья (справочно)", "Кол-во", "Ставка / доля / сумма"]
TPL_INDIRECT_SECTIONS = {
    "land": "Земля",
    "infra": "Магистральные сети (на весь участок)",
    "landscape": "Благоустройство мест общего пользования",
    "social": "Социальные объекты (школы/сады)",
    "soft": "Прочие Soft Costs",
}
# Применимость цен/ставок по типу блока (в неприменимых ячейках — база, чтобы
# таблица «База + блоки» не показывала ложное «изменено вручную»)
_PRICE_APPLIES = {
    "Цена жилья, руб/м2": {TYPE_RESIDENTIAL}, "Цена подземного м/м, руб": {TYPE_RESIDENTIAL},
    "Цена наземного м/м, руб": {TYPE_PARKING},
    "Цена коммерции, руб/м2": {TYPE_RESIDENTIAL, TYPE_PARKING}, "Цена кладовых, руб/м2": {TYPE_RESIDENTIAL, TYPE_PARKING},
}
_PARK_APPLIES = {
    "Ставка СМР подземного м/м, руб": {TYPE_RESIDENTIAL},
    "Ставка СМР наземного м/м, руб": {TYPE_PARKING},
    "Ставка СМР коммерции паркинга, руб/м2": {TYPE_PARKING},
    "Ставка СМР кладовых паркинга, руб/м2": {TYPE_PARKING},
}
_CODE_KIND = {}  # код -> (вид, наименование, что вводить)
for _c, _n, _g, _u, _b, _r in ITEMS:
    _CODE_KIND[_c] = ("ITEMS", _n, f"ставка, {_u} (стадия ЭП)")
for _c, _n, _u, _r in ITEMS_G_AREA:
    _CODE_KIND[_c] = ("G_AREA", _n, "ставка, руб/га площади участка блока")
for _c, _n, _u, _q, _r in ITEMS_G_LENGTH:
    _CODE_KIND[_c] = ("G_LENGTH", _n, f"кол-во ({_u}) и ставка, руб/ед.")
for _c, _n, _r in ITEMS_Z_PCT:
    _CODE_KIND[_c] = ("Z_PCT", _n, "доля от базы Z (0,08 = 8 %)")
for _c, _n, _s in ITEMS_Z_FIXED:
    _CODE_KIND[_c] = ("Z_FIXED", _n, "сумма, руб")


def _norm_code(code) -> str:
    """«А.10.10» (кириллица), «a 10 10», «Z.150.10» -> код каталога модели."""
    s = str(code or "").strip().upper()
    s = s.translate(str.maketrans({"А": "A", "В": "B", "С": "C", "Е": "E", "Д": "D"}))
    s = re.sub(r"[\s,;/]+", ".", s).strip(".")
    if s in _CODE_KIND:
        return s
    if s.startswith("Z.150"):
        return "Z.150"
    return s


def _xl_num(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(" ", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return 0.0


def _tpl_style_header(ws, n_cols, row=1):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill(start_color="41AA37", end_color="41AA37", fill_type="solid")
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 42


def collect_template_data() -> dict:
    """Текущий проект из session_state -> структура шаблона (для выгрузки)."""
    ss = st.session_state
    bdf = ss.get("blocks_df")
    data = {
        "project": {
            "Наименование проекта": ss.get("project_name_input", ""),
            "Город": ss.get("project_city_input", ""),
            "Стадия (МП/ЭП)": "МП" if ss.get("design_stage") == STAGE_MP else "ЭП",
            "Сценарий": ss.get("scenario_select", "Базовый"),
        },
        "blocks": [], "items": [], "infra": [], "indirect": [],
    }
    if bdf is None:
        return data
    prices = ss.get("_cascade_prices", {})
    parks = ss.get("_cascade_smr_parking", {})
    for _, row in bdf.iterrows():
        name, btype = str(row["Название блока"]), row["Тип блока"]
        tep = ss.get("tep_store", {}).get(name, {})
        rec = {}
        for hdr, kind, key in TPL_BLOCK_COLS:
            if kind == "name":
                v = name
            elif kind == "type":
                v = btype
            elif kind == "tep":
                v = tep.get(key, 0.0)
            elif kind == "main":
                v = row.get(key, 0.0)
            elif kind == "price":
                v = prices.get(name, {}).get(key, ss.get(f"_price_default_{key}", 0.0))
            elif kind == "park":
                v = parks.get(name, {}).get(key, ss.get(f"_smr_parking_default_{key}", 0.0))
            elif kind == "mp":
                v = ss.get("block_mp_rate", {}).get(name, 0.0) if btype == TYPE_RESIDENTIAL else 0.0
            elif kind == "stops":
                v = ss.get("block_elevator_stops", {}).get(name, 0.0)
            rec[hdr] = v
        data["blocks"].append(rec)
        # Статьи СМР — только ненулевые
        if btype == TYPE_RESIDENTIAL and name in ss.get("block_rates", {}):
            for _, r in ss["block_rates"][name].iterrows():
                v = _xl_num(r["Ставка, руб/ед."])
                if v:
                    data["items"].append([name, r["Код"], r["Статья затрат"], None, v])
        for store, qcol, vcol in [("block_g_area", None, "Ставка, руб/ед."), ("block_g_length", "Кол-во", "Ставка, руб/ед."),
                                  ("block_z_pct", None, "Ставка, доля от СМР+G"), ("block_z_fixed", None, "Сумма, руб")]:
            df = ss.get(store, {}).get(name)
            if df is None:
                continue
            for _, r in df.iterrows():
                v = _xl_num(r[vcol])
                q = _xl_num(r[qcol]) if qcol else None
                if v or q:
                    data["items"].append([name, r["Код"], r["Статья затрат"], q, v])
        idf = ss.get("block_infra", {}).get(name)
        if idf is not None and "Сумма, руб" in idf.columns:
            for _, r in idf.iterrows():
                if _xl_num(r["Сумма, руб"]):
                    data["infra"].append([name, r.get("Статья затрат", ""), _xl_num(r["Сумма, руб"])])
    for key, label in TPL_INDIRECT_SECTIONS.items():
        df = ss.get("indirect_items", {}).get(key)
        if df is not None and "Сумма, руб" in df.columns:
            for _, r in df.iterrows():
                if _xl_num(r["Сумма, руб"]):
                    data["indirect"].append([label, r.get("Наименование", ""), _xl_num(r["Сумма, руб"])])
    return data


TPL_SHEET_GUIDE = "Инструкция"
TPL_SHEET_MP = "МП — коробка"
TPL_SHEET_EP_PARAMS = "ЭП — параметры"
TPL_SHEET_EP_RATES = "ЭП — ставки A–E"
TPL_SHEET_GZ = "G и Z"
TPL_SHEET_LISTS = "Списки"
TPL_MIN_BLOCKS = 20        # строк/колонок под блоки в пустом шаблоне
TPL_FIRST_BLOCK_COL = 5    # в матрицах блоки начинаются с колонки E
TPL_INFRA_SUGGEST = [
    "Благоустройство общественных пространств", "Плейхаб", "Парк / сквер", "Спортхаб", "Догхаб",
    "Пешеходный бульвар", "Дороги и проезды общего пользования", "Инженерные сети общего пользования",
]
TPL_INDIRECT_SUGGEST = [
    "Покупка земельного участка", "Аренда земельного участка", "Земля — очередь …",
    "ТП / электроснабжение (магистральное)", "Водовод", "Канализация (КНС, коллектор)", "Теплосеть магистральная",
    "Благоустройство МОП", "Школа", "Детский сад", "Маркетинг", "Управление проектом", "Прочее",
]
_COMMON_KINDS = ("name", "type", "main", "price", "park")
_COMMON_TEP = {"Общая площадь квартир (с летними, с коэф.), м2", "Площадь коммерции, м2",
               "Площадь кладовых в доме, м2", "Площадь участка блока, га"}


def _tpl_common_cols():
    return [(h, k, key) for h, k, key in TPL_BLOCK_COLS if k in _COMMON_KINDS or (k == "tep" and key in _COMMON_TEP)]


def _tpl_ep_cols():
    return [(h, k, key) for h, k, key in TPL_BLOCK_COLS if k == "stops" or (k == "tep" and key not in _COMMON_TEP)]


def _tpl_gz_rows():
    """Строки листа «G и Z»: (код, статья, что вводить, вид, поле)."""
    rows = []
    for c, n, u, r in ITEMS_G_AREA:
        rows.append((c, n, "ставка, руб/га участка", "G_AREA", "rate"))
    for c, n, u, q, r in ITEMS_G_LENGTH:
        rows.append((c, n, f"кол-во, {u}", "G_LENGTH", "qty"))
        rows.append((c, n, f"ставка, руб/{u.split()[0]}", "G_LENGTH", "rate"))
    for c, n, r in ITEMS_Z_PCT:
        rows.append((c, n, "доля от базы Z (8 % = 0,08)", "Z_PCT", "rate"))
    for c, n, s in ITEMS_Z_FIXED:
        rows.append((c, n, "сумма, руб", "Z_FIXED", "sum"))
    return rows


def build_template_bytes(data: dict = None) -> bytes:
    """Шаблон загрузки: МП и ЭП на отдельных листах, коды статей преднастроены,
    блоки подтягиваются формулами с листа «Блоки», выпадающие списки."""
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.comments import Comment
    data = data or {}
    blocks = data.get("blocks") or []
    n_bl = max(TPL_MIN_BLOCKS, len(blocks))
    last_bl_row = n_bl + 1
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    input_fill = PatternFill("solid", fgColor="FFFFFF")
    ro_fill = PatternFill("solid", fgColor="F2F2F2")
    sec_fill = PatternFill("solid", fgColor="EEF7EC")
    fnt = Font(name="Arial", size=10)

    def grid(ws, r1, r2, c1, c2, fill=None, fmt=None):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = ws.cell(r, c)
                cell.border = border
                cell.font = fnt
                if fill is not None:
                    cell.fill = fill
                if fmt:
                    cell.number_format = fmt

    def dv_list(ws, formula, rng, strict=True, prompt=None):
        dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                            showErrorMessage=strict, showInputMessage=bool(prompt))
        if prompt:
            dv.promptTitle, dv.prompt = "Подсказка", prompt
        if not strict:
            dv.errorStyle = "information"
        ws.add_data_validation(dv)
        dv.add(rng)

    wb = Workbook()
    # --- Инструкция
    wsg = wb.active
    wsg.title = TPL_SHEET_GUIDE
    guide = [
        ("Шаблон загрузки финмодели", None),
        ("1. «Проект»", "название, город, стадия (МП или ЭП — из списка), сценарий."),
        ("2. «Блоки»", "строка на урбан-блок: название, тип (из списка), площади, м/м, участок, цены, ставки СМР паркинга. "
                       "Названия блоков дальше подставляются на всех листах сами."),
        ("3. «МП — коробка»", "стадия МП: ставка коробки, руб/м² NSA по каждому жилому блоку."),
        ("4. «ЭП — параметры» и «ЭП — ставки A–E»", "стадия ЭП: параметры домов и ставки по 32 статьям A–E. Коды преднастроены, "
                                                   "вводится только ставка в колонке блока."),
        ("5. «G и Z»", "обе стадии: наружные работы и прочие затраты по кодам — ставка, кол-во, доля или сумма в колонке блока."),
        ("6. «Инфраструктура»", "Z.50.10 по блокам: блок и статья — из списка (статью можно вписать свою), сумма."),
        ("7. «Косвенные»", "затраты на весь участок: раздел — из списка, статья — из списка или своя, сумма."),
        ("Загрузка", "в приложении: боковая панель → «Данные из Excel» → выбрать файл → «Загрузить данные»."),
        ("Правила", "белые ячейки — ввод, серые — подставляются сами. Не вставляйте и не удаляйте строки на листе «Блоки» "
                    "между блоками: колонки на других листах привязаны к порядку строк. Пустые ячейки = 0."),
    ]
    for i, (a, b) in enumerate(guide, start=1):
        wsg.cell(i, 1, a).font = Font(name="Arial", bold=True, size=14 if i == 1 else 10,
                                      color="41AA37" if i == 1 else "1F1F1F")
        if b:
            wsg.cell(i, 2, b).font = fnt
            wsg.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
    wsg.column_dimensions["A"].width = 34
    wsg.column_dimensions["B"].width = 110

    # --- Списки (служебный)
    wsl = wb.create_sheet(TPL_SHEET_LISTS)
    lists = {
        "A": ["МП", "ЭП"], "B": list(SCENARIOS.keys()), "C": BLOCK_TYPES,
        "D": list(TPL_INDIRECT_SECTIONS.values()), "E": TPL_INFRA_SUGGEST, "F": TPL_INDIRECT_SUGGEST,
    }
    for col, vals in lists.items():
        for i, v in enumerate(vals, start=1):
            wsl[f"{col}{i}"] = v
    rng = {col: f"{TPL_SHEET_LISTS}!${col}$1:${col}${len(v)}" for col, v in lists.items()}
    wsl.sheet_state = "hidden"

    # --- Проект
    ws = wb.create_sheet(TPL_SHEET_PROJECT)
    ws.append(["Параметр", "Значение"])
    _tpl_style_header(ws, 2)
    proj = data.get("project") or {"Наименование проекта": "ЖК «Пример»", "Город": "",
                                   "Стадия (МП/ЭП)": "МП", "Сценарий": "Базовый"}
    for k in ["Наименование проекта", "Город", "Стадия (МП/ЭП)", "Сценарий"]:
        ws.append([k, proj.get(k, "")])
    grid(ws, 2, 5, 1, 1, ro_fill)
    grid(ws, 2, 5, 2, 2, input_fill)
    dv_list(ws, f"={rng['A']}", "B4")
    dv_list(ws, f"={rng['B']}", "B5")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 44

    # --- Блоки (общие для МП и ЭП)
    common = _tpl_common_cols()
    wsb = wb.create_sheet(TPL_SHEET_BLOCKS)
    wsb.append([h for h, _, _ in common])
    _tpl_style_header(wsb, len(common))
    for i in range(n_bl):
        rec = blocks[i] if i < len(blocks) else {}
        wsb.append([rec.get(h) for h, _, _ in common])
    grid(wsb, 2, last_bl_row, 1, 2, input_fill)
    grid(wsb, 2, last_bl_row, 3, len(common), input_fill, "#,##0.##")
    dv_list(wsb, f"={rng['C']}", f"B2:B{last_bl_row}", prompt="Жилой блок или наземный/многоуровневый паркинг")
    for i in range(1, len(common) + 1):
        wsb.column_dimensions[get_column_letter(i)].width = 15
    wsb.column_dimensions["A"].width = 28
    wsb.column_dimensions["B"].width = 32
    wsb.freeze_panes = "C2"
    names_rng = f"{TPL_SHEET_BLOCKS}!$A$2:$A${last_bl_row}"

    def name_formula(i):
        return f'=IF({TPL_SHEET_BLOCKS}!A{i + 2}="","",{TPL_SHEET_BLOCKS}!A{i + 2})'

    def type_formula(i):
        return f'=IF({TPL_SHEET_BLOCKS}!A{i + 2}="","",{TPL_SHEET_BLOCKS}!B{i + 2})'

    # --- МП — коробка
    mp_hdr = next(h for h, k, _ in TPL_BLOCK_COLS if k == "mp")
    wsm = wb.create_sheet(TPL_SHEET_MP)
    wsm.append(["Блок", "Тип блока", mp_hdr])
    _tpl_style_header(wsm, 3)
    for i in range(n_bl):
        rec = blocks[i] if i < len(blocks) else {}
        wsm.append([name_formula(i), type_formula(i), rec.get(mp_hdr) or None])
    grid(wsm, 2, last_bl_row, 1, 2, ro_fill)
    grid(wsm, 2, last_bl_row, 3, 3, input_fill, "#,##0")
    wsm.cell(1, 3).comment = Comment("Только для жилых блоков. Коробка паркинга — ставкой за м/м на листе «Блоки».", "Шаблон")
    wsm.column_dimensions["A"].width = 28
    wsm.column_dimensions["B"].width = 32
    wsm.column_dimensions["C"].width = 22

    # --- ЭП — параметры
    epc = _tpl_ep_cols()
    wse = wb.create_sheet(TPL_SHEET_EP_PARAMS)
    wse.append(["Блок", "Тип блока"] + [h.replace("ЭП: ", "") for h, _, _ in epc])
    _tpl_style_header(wse, 2 + len(epc))
    for i in range(n_bl):
        rec = blocks[i] if i < len(blocks) else {}
        wse.append([name_formula(i), type_formula(i)] + [rec.get(h) or None for h, _, _ in epc])
    grid(wse, 2, last_bl_row, 1, 2, ro_fill)
    grid(wse, 2, last_bl_row, 3, 2 + len(epc), input_fill, "#,##0.##")
    wse.column_dimensions["A"].width = 28
    wse.column_dimensions["B"].width = 32
    for i in range(3, 3 + len(epc)):
        wse.column_dimensions[get_column_letter(i)].width = 14
    wse.freeze_panes = "C2"

    # --- матрицы: код × блок
    items = data.get("items") or []
    vals = {}  # (блок, код, поле) -> значение
    for n, code, _lbl, q, v in items:
        code = _norm_code(code)
        kind = _CODE_KIND.get(code, (None,))[0]
        if kind == "G_LENGTH":
            if q:
                vals[(n, code, "qty")] = q
            if v:
                vals[(n, code, "rate")] = v
        elif kind:
            vals[(n, code, "rate" if kind != "Z_FIXED" else "sum")] = v
    bl_names = [b.get("Название блока") for b in blocks]

    def matrix(ws, rows, fmt_of):
        hdr = ["Код", "Статья", "Что вводить", "Группа / ед."]
        ws.append(hdr + [None] * n_bl)
        for j in range(n_bl):
            ws.cell(1, TPL_FIRST_BLOCK_COL + j, name_formula(j))
        _tpl_style_header(ws, TPL_FIRST_BLOCK_COL - 1 + n_bl)
        prev_group = None
        r = 2
        for code, name, what, grp, field in rows:
            ws.cell(r, 1, code)
            ws.cell(r, 2, name)
            ws.cell(r, 3, what)
            ws.cell(r, 4, grp)
            for j in range(n_bl):
                bn = bl_names[j] if j < len(bl_names) else None
                v = vals.get((bn, code, field)) if bn else None
                if v:
                    ws.cell(r, TPL_FIRST_BLOCK_COL + j, v)
            grid(ws, r, r, 1, 4, sec_fill if grp != prev_group else ro_fill)
            grid(ws, r, r, TPL_FIRST_BLOCK_COL, TPL_FIRST_BLOCK_COL - 1 + n_bl, input_fill, fmt_of(field, code))
            prev_group = grp
            r += 1
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 44
        ws.column_dimensions["C"].width = 24
        ws.column_dimensions["D"].width = 16
        for j in range(n_bl):
            ws.column_dimensions[get_column_letter(TPL_FIRST_BLOCK_COL + j)].width = 15
        ws.freeze_panes = ws.cell(2, TPL_FIRST_BLOCK_COL)

    wsr = wb.create_sheet(TPL_SHEET_EP_RATES)
    matrix(wsr, [(c, n, f"ставка, {u}", f"{g} · {GROUP_LABELS[g]}", "rate") for c, n, g, u, b, r in ITEMS],
           lambda f, c: "#,##0.##")
    wsz = wb.create_sheet(TPL_SHEET_GZ)
    gz_rows = [(c, n, what, {"G_AREA": "G · на площадь", "G_LENGTH": "G · сети", "Z_PCT": "Z · доля",
                             "Z_FIXED": "Z · сумма"}[kind], field) for c, n, what, kind, field in _tpl_gz_rows()]
    matrix(wsz, gz_rows, lambda f, c: "0.0%" if _CODE_KIND.get(c, ("",))[0] == "Z_PCT" else "#,##0.##")

    # --- Инфраструктура
    wsf = wb.create_sheet(TPL_SHEET_INFRA)
    wsf.append(["Блок", "Статья", "Сумма, руб"])
    _tpl_style_header(wsf, 3)
    infra_rows = data.get("infra") or []
    n_inf = max(60, len(infra_rows) + 20)
    for r in infra_rows:
        wsf.append(r)
    grid(wsf, 2, n_inf + 1, 1, 2, input_fill)
    grid(wsf, 2, n_inf + 1, 3, 3, input_fill, "#,##0")
    dv_list(wsf, f"={names_rng}", f"A2:A{n_inf + 1}", prompt="Блок с листа «Блоки»")
    dv_list(wsf, f"={rng['E']}", f"B2:B{n_inf + 1}", strict=False, prompt="Выберите из списка или впишите свою статью")
    for col, w in zip("ABC", [28, 50, 18]):
        wsf.column_dimensions[col].width = w

    # --- Косвенные
    wsk = wb.create_sheet(TPL_SHEET_INDIRECT)
    wsk.append(["Раздел", "Статья", "Сумма, руб"])
    _tpl_style_header(wsk, 3)
    ind_rows = data.get("indirect") or []
    n_ind = max(50, len(ind_rows) + 20)
    for r in ind_rows:
        wsk.append(r)
    grid(wsk, 2, n_ind + 1, 1, 2, input_fill)
    grid(wsk, 2, n_ind + 1, 3, 3, input_fill, "#,##0")
    dv_list(wsk, f"={rng['D']}", f"A2:A{n_ind + 1}", prompt="Раздел пула косвенных расходов")
    dv_list(wsk, f"={rng['F']}", f"B2:B{n_ind + 1}", strict=False, prompt="Выберите из списка или впишите свою статью")
    for col, w in zip("ABC", [40, 50, 18]):
        wsk.column_dimensions[col].width = w

    # --- Справочник кодов
    wsc = wb.create_sheet(TPL_SHEET_CODES)
    wsc.append(["Код", "Статья", "Где вводить / что вводить"])
    _tpl_style_header(wsc, 3)
    for c, (kind, n, what) in _CODE_KIND.items():
        where = TPL_SHEET_EP_RATES if kind == "ITEMS" else TPL_SHEET_GZ
        wsc.append([c, n, f"«{where}»: {what}"])
    for col, w in zip("ABC", [12, 60, 70]):
        wsc.column_dimensions[col].width = w

    wb._sheets = [wb[s] for s in [TPL_SHEET_GUIDE, TPL_SHEET_PROJECT, TPL_SHEET_BLOCKS, TPL_SHEET_MP,
                                  TPL_SHEET_EP_PARAMS, TPL_SHEET_EP_RATES, TPL_SHEET_GZ, TPL_SHEET_INFRA,
                                  TPL_SHEET_INDIRECT, TPL_SHEET_CODES, TPL_SHEET_LISTS]]
    wb.active = 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def parse_template(file_bytes: bytes):
    """Шаблон -> (новые значения session_state, предупреждения). Ошибку формата
    возвращает строкой в предупреждениях и state=None."""
    from openpyxl import load_workbook
    warns = []
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, [f"Файл не читается как Excel: {e}"]
    if TPL_SHEET_BLOCKS not in wb.sheetnames:
        return None, [f"Нет листа «{TPL_SHEET_BLOCKS}» — это не шаблон загрузки."]

    proj = {}
    if TPL_SHEET_PROJECT in wb.sheetnames:
        for r in wb[TPL_SHEET_PROJECT].iter_rows(min_row=2, values_only=True):
            if r and r[0]:
                proj[str(r[0]).strip()] = r[1] if len(r) > 1 else None

    wsb = wb[TPL_SHEET_BLOCKS]
    hdr = [str(c.value).strip() if c.value is not None else "" for c in wsb[1]]
    col_idx = {h: i for i, h in enumerate(hdr)}
    missing = [h for h, _, _ in TPL_BLOCK_COLS[:2] if h not in col_idx]
    if missing:
        return None, [f"На листе «{TPL_SHEET_BLOCKS}» нет колонок: {', '.join(missing)}."]
    blocks = []
    seen = set()
    for pos, r in enumerate(wsb.iter_rows(min_row=2, values_only=True)):
        name = r[col_idx["Название блока"]] if col_idx["Название блока"] < len(r) else None
        if name is None or str(name).strip() == "":
            continue
        name = str(name).strip()
        if name in seen:
            warns.append(f"Блок «{name}» повторяется — вторая строка пропущена.")
            continue
        seen.add(name)
        t_raw = str(r[col_idx["Тип блока"]] or "").strip()
        btype = TPL_TYPE_ALIASES.get(t_raw.lower(), t_raw if t_raw in BLOCK_TYPES else None)
        if btype is None:
            warns.append(f"Блок «{name}»: тип «{t_raw}» не распознан — принят «{TYPE_RESIDENTIAL}».")
            btype = TYPE_RESIDENTIAL
        rec = {"name": name, "type": btype, "pos": pos}
        for h, kind, key in TPL_BLOCK_COLS[2:]:
            rec[h] = _xl_num(r[col_idx[h]]) if h in col_idx and col_idx[h] < len(r) else 0.0
        blocks.append(rec)
    # Стадии на отдельных листах — строка листа = строка блока на «Блоки»
    _mp_hdr0 = next(h for h, k, _ in TPL_BLOCK_COLS if k == "mp")
    if TPL_SHEET_MP in wb.sheetnames:
        _rows = list(wb[TPL_SHEET_MP].iter_rows(min_row=2, values_only=True))
        for b in blocks:
            if b["pos"] < len(_rows) and len(_rows[b["pos"]]) > 2 and _rows[b["pos"]][2] not in (None, ""):
                b[_mp_hdr0] = _xl_num(_rows[b["pos"]][2])
    if TPL_SHEET_EP_PARAMS in wb.sheetnames:
        _wse = wb[TPL_SHEET_EP_PARAMS]
        _eh = [str(c.value).strip() if c.value is not None else "" for c in _wse[1]]
        _rows = list(_wse.iter_rows(min_row=2, values_only=True))
        for h, k, key in _tpl_ep_cols():
            short = h.replace("ЭП: ", "")
            if short not in _eh:
                continue
            ci = _eh.index(short)
            for b in blocks:
                if b["pos"] < len(_rows) and ci < len(_rows[b["pos"]]) and _rows[b["pos"]][ci] not in (None, ""):
                    b[h] = _xl_num(_rows[b["pos"]][ci])
    if not blocks:
        return None, [f"На листе «{TPL_SHEET_BLOCKS}» нет ни одного блока."]
    names = [b["name"] for b in blocks]
    btype_of = {b["name"]: b["type"] for b in blocks}

    ss = {}
    ss["project_name_input"] = str(proj.get("Наименование проекта") or "Проект из файла")
    ss["project_city_input"] = str(proj.get("Город") or "")
    stage = str(proj.get("Стадия (МП/ЭП)") or "МП").strip().upper()
    ss["design_stage"] = STAGE_EP if stage.startswith("ЭП") else STAGE_MP
    scen = str(proj.get("Сценарий") or "Базовый").strip()
    ss["scenario_select"] = scen if scen in SCENARIOS else "Базовый"

    # Блоки, ТЭП, лифты
    main_rows, tep_store, stops = [], {}, {}
    for b in blocks:
        tep = {k: 0.0 for k in TEP_TO_MAIN_COL}
        tep.update({c: 0.0 for c in EXTRA_EP_PARAM_COLS})
        mrow = {c: 0.0 for c in MAIN_TABLE_NUMERIC_COLS}
        for h, kind, key in TPL_BLOCK_COLS[2:]:
            if kind == "tep":
                tep[key] = b[h]
            elif kind == "main":
                mrow[key] = b[h]
            elif kind == "stops":
                stops[b["name"]] = b[h]
        if b["type"] == TYPE_PARKING:
            tep["Общая площадь квартир (с летними, с коэф.), м2"] = 0.0
        else:
            mrow["Наземный/Многоуровневый паркинг, м/м"] = 0.0
        for tep_col, main_col in TEP_TO_MAIN_COL.items():
            if main_col in mrow:
                mrow[main_col] = tep[tep_col]
        main_rows.append({"Название блока": b["name"], "Тип блока": b["type"], **mrow})
        tep_store[b["name"]] = tep
    ss["blocks_df"] = pd.DataFrame(main_rows)[MAIN_TABLE_COLS]
    ss["tep_store"] = tep_store
    ss["block_elevator_stops"] = {n: stops.get(n, 0.0) for n in names if btype_of[n] == TYPE_RESIDENTIAL}

    def _cascade(store_key, keymap, applies, default_key_fn):
        """Значения по блокам -> база (самое частое значение среди применимых) +
        cascade-хранилище; в неприменимых ячейках — база."""
        store = {n: {} for n in names}
        prev = {}
        for hdr_name, col in keymap.items():
            vals = [b[hdr_name] for b in blocks if b["type"] in applies.get(col, set(BLOCK_TYPES))]
            nz = [v for v in vals if v]
            base = max(set(nz), key=nz.count) if nz else 0.0
            ss[default_key_fn(col)] = float(base)
            prev[col] = float(base)
            for b in blocks:
                ok = b["type"] in applies.get(col, set(BLOCK_TYPES))
                store[b["name"]][col] = float(b[hdr_name]) if ok else float(base)
        ss[f"_cascade_{store_key}"] = store
        ss[f"_cascade_{store_key}_prev"] = prev

    _cascade("prices", {h: k for h, kind, k in TPL_BLOCK_COLS if kind == "price"}, _PRICE_APPLIES,
             lambda c: f"_price_default_{c}")
    _cascade("smr_parking", {h: k for h, kind, k in TPL_BLOCK_COLS if kind == "park"}, _PARK_APPLIES,
             lambda c: f"_smr_parking_default_{c}")
    _mp_hdr = next(h for h, kind, _ in TPL_BLOCK_COLS if kind == "mp")
    _res_blocks = [b for b in blocks if b["type"] == TYPE_RESIDENTIAL]
    _nz = [b[_mp_hdr] for b in _res_blocks if b[_mp_hdr]]
    _mp_base = max(set(_nz), key=_nz.count) if _nz else 0.0
    ss["_mp_korobka_default"] = float(_mp_base)
    ss["_cascade_mp_korobka"] = {b["name"]: {"Ставка, руб/м2 NSA": float(b[_mp_hdr])} for b in _res_blocks}
    ss["_cascade_mp_korobka_prev"] = {"Ставка, руб/м2 NSA": float(_mp_base)}
    ss["block_mp_rate"] = {b["name"]: float(b[_mp_hdr]) for b in _res_blocks}

    # Статьи СМР
    rates = {n: DEFAULT_RATES_DF.copy() for n in names if btype_of[n] == TYPE_RESIDENTIAL}
    g_area = {n: DEFAULT_G_AREA_DF.copy() for n in names}
    g_len = {n: DEFAULT_G_LENGTH_DF.copy() for n in names}
    z_pct = {n: DEFAULT_Z_PCT_DF.copy() for n in names}
    z_fix = {n: DEFAULT_Z_FIXED_DF.copy() for n in names}
    raw_items = []  # (источник, блок, код, кол-во, значение)
    if TPL_SHEET_ITEMS in wb.sheetnames:  # прежний формат: длинный список
        for i, r in enumerate(wb[TPL_SHEET_ITEMS].iter_rows(min_row=2, values_only=True), start=2):
            if not r or r[0] is None or r[1] is None:
                continue
            raw_items.append((f"«{TPL_SHEET_ITEMS}», строка {i}", str(r[0]).strip(), r[1],
                              _xl_num(r[3]) if len(r) > 3 else 0.0, _xl_num(r[4]) if len(r) > 4 else 0.0))
    pos_name = {b["pos"]: b["name"] for b in blocks}
    for sheet in (TPL_SHEET_EP_RATES, TPL_SHEET_GZ):
        if sheet not in wb.sheetnames:
            continue
        acc = {}  # (блок, код) -> [кол-во, значение]
        for i, r in enumerate(wb[sheet].iter_rows(min_row=2, values_only=True), start=2):
            if not r or r[0] is None:
                continue
            code = _norm_code(r[0])
            is_qty = "кол-во" in str(r[2] or "").lower()
            for j in range(TPL_FIRST_BLOCK_COL - 1, len(r)):
                v = r[j]
                if v in (None, "") or not _xl_num(v):
                    continue
                n = pos_name.get(j - (TPL_FIRST_BLOCK_COL - 1))
                if n is None:
                    warns.append(f"«{sheet}», {get_column_letter(j + 1)}{i}: значение в колонке без блока — пропущено.")
                    continue
                a = acc.setdefault((n, code), [0.0, 0.0, i])
                if is_qty:
                    a[0] = _xl_num(v)
                else:
                    a[1] = _xl_num(v)
        for (n, code), (q, v, i) in acc.items():
            raw_items.append((f"«{sheet}», строка {i}", n, code, q, v))
    for src_label, n, code_raw, qty, val in raw_items:
        if True:
            code = _norm_code(code_raw)
            r = (n, code_raw)
            i = src_label
            if n not in btype_of:
                warns.append(f"{i}: блока «{n}» нет на листе «Блоки» — пропущено.")
                continue
            kind = _CODE_KIND.get(code, (None,))[0]
            if kind is None:
                warns.append(f"{i}: код «{r[1]}» не найден в справочнике — пропущено.")
                continue
            if kind == "ITEMS":
                if btype_of[n] != TYPE_RESIDENTIAL:
                    warns.append(f"{i}: статьи A–E — только для жилых блоков, «{n}» — паркинг.")
                    continue
                df = rates[n]
                df.loc[df["Код"] == code, "Ставка, руб/ед."] = df.loc[df["Код"] == code, "Ставка, руб/ед."] + val
            elif kind == "G_AREA":
                df = g_area[n]
                df.loc[df["Код"] == code, "Ставка, руб/ед."] = df.loc[df["Код"] == code, "Ставка, руб/ед."] + val
            elif kind == "G_LENGTH":
                df = g_len[n]
                m = df["Код"] == code
                if float(df.loc[m, "Кол-во"].iloc[0]) and qty and float(df.loc[m, "Кол-во"].iloc[0]) != qty:
                    warns.append(f"{i}: код {code} у «{n}» повторяется с другим кол-вом — взята последняя строка.")
                df.loc[m, "Кол-во"] = qty if qty else 1.0
                df.loc[m, "Ставка, руб/ед."] = val
            elif kind == "Z_PCT":
                df = z_pct[n]
                df.loc[df["Код"] == code, "Ставка, доля от СМР+G"] = val
            elif kind == "Z_FIXED":
                df = z_fix[n]
                df.loc[df["Код"] == code, "Сумма, руб"] = df.loc[df["Код"] == code, "Сумма, руб"] + val
    ss["block_rates"], ss["block_g_area"], ss["block_g_length"] = rates, g_area, g_len
    ss["block_z_pct"], ss["block_z_fixed"] = z_pct, z_fix

    infra = {n: EMPTY_INFRA_DF.copy() for n in names}
    if TPL_SHEET_INFRA in wb.sheetnames:
        for i, r in enumerate(wb[TPL_SHEET_INFRA].iter_rows(min_row=2, values_only=True), start=2):
            if not r or r[0] is None:
                continue
            n = str(r[0]).strip()
            if not (len(r) > 2 and _xl_num(r[2])) and not (len(r) > 1 and r[1]):
                continue
            if n not in infra:
                warns.append(f"«{TPL_SHEET_INFRA}», строка {i}: блока «{n}» нет — пропущено.")
                continue
            infra[n] = pd.concat([infra[n], pd.DataFrame([{"Статья затрат": r[1] or "", "Сумма, руб": _xl_num(r[2])}])],
                                 ignore_index=True)
    ss["block_infra"] = infra

    ind = {k: [] for k in TPL_INDIRECT_SECTIONS}
    label_to_key = {v.lower(): k for k, v in TPL_INDIRECT_SECTIONS.items()}
    label_to_key.update({"магистральные сети": "infra", "благоустройство моп": "landscape",
                         "социальные объекты": "social", "прочие soft costs": "soft", "soft costs": "soft"})
    if TPL_SHEET_INDIRECT in wb.sheetnames:
        for i, r in enumerate(wb[TPL_SHEET_INDIRECT].iter_rows(min_row=2, values_only=True), start=2):
            if not r or r[0] is None:
                continue
            if not (len(r) > 2 and _xl_num(r[2])) and not (len(r) > 1 and r[1]):
                continue
            key = label_to_key.get(str(r[0]).strip().lower())
            if key is None:
                warns.append(f"«{TPL_SHEET_INDIRECT}», строка {i}: раздел «{r[0]}» не распознан — записан в «Прочие Soft Costs».")
                key = "soft"
            ind[key].append({"Наименование": r[1] or TPL_INDIRECT_SECTIONS[key], "Сумма, руб": _xl_num(r[2])})
    for key, label in TPL_INDIRECT_SECTIONS.items():
        rows = ind[key][:MAX_INDIRECT_ITEMS]
        if len(ind[key]) > MAX_INDIRECT_ITEMS:
            warns.append(f"«{label}»: больше {MAX_INDIRECT_ITEMS} статей — лишние свернуты в последнюю строку.")
            extra = sum(x["Сумма, руб"] for x in ind[key][MAX_INDIRECT_ITEMS - 1:])
            rows = ind[key][:MAX_INDIRECT_ITEMS - 1] + [{"Наименование": "Прочее (свернуто при загрузке)", "Сумма, руб": extra}]
        ind[key] = pd.DataFrame(rows or [{"Наименование": label, "Сумма, руб": 0.0}])
    ss["indirect_items"] = ind
    return ss, warns


def apply_imported_state(new_state: dict) -> None:
    """Заменяет ввод на экране загруженным проектом (сохранения на диске не трогает)."""
    keep_user = st.session_state.get("user_name_input", "")
    st.session_state.clear()
    st.session_state["user_name_input"] = keep_user
    st.session_state["_bootstrapped"] = True
    st.session_state["_autosave_loaded"] = False
    for k, v in new_state.items():
        st.session_state[k] = v


def _cb_import_template():
    f = st.session_state.get("_tpl_upload")
    if f is None:
        st.session_state["_tpl_msgs"] = ["Сначала выберите файл."]
        return
    new_state, warns = parse_template(f.getvalue())
    if new_state is None:
        st.session_state["_tpl_msgs"] = warns
        return
    apply_imported_state(new_state)
    st.session_state["_tpl_msgs"] = [f"Загружено: «{new_state['project_name_input']}», "
                                     f"блоков — {len(new_state['blocks_df'])}."] + warns


EMPTY_TEMPLATE_BYTES = None  # собирается при первом обращении


# ======================================================================
# 4. БОКОВАЯ ПАНЕЛЬ — ПОЛЬЗОВАТЕЛЬ, ПРОЕКТЫ, ХРАНИЛИЩЕ
# ======================================================================
with st.sidebar:
    st.image(_talan_logo_bytes, width=150)
    st.divider()
    st.text_input(
        "Ваше имя",
        key="user_name_input",
        help="Вместе с названием проекта (в шапке) определяет, куда сохраняются данные.",
    )
    _saves = list_saved_projects()
    _options = [f"{s['user_name']} — {s['project_name']}" for s in _saves]
    if st.session_state.get("_project_picker") not in _options:
        st.session_state.pop("_project_picker", None)
    _picked = st.selectbox(
        "Сохранённые проекты", _options, index=None, placeholder="— выберите проект —", key="_project_picker",
    )
    _picked_save = _saves[_options.index(_picked)] if _picked in _options else None
    sb1, sb2, sb3 = st.columns([1.25, 1.1, 0.75])

    def _cb_open(ref):
        # в callback: запускается до отрисовки виджетов — можно менять их значения
        if autosave_load(ref):
            st.session_state["_autosave_loaded"] = True
            st.session_state.pop("_last_saved_sig", None)
            st.session_state.pop("_stage_seg", None)

    def _cb_new():
        _keep_user = st.session_state.get("user_name_input", "")
        st.session_state.clear()
        st.session_state["user_name_input"] = _keep_user
        st.session_state["_bootstrapped"] = True
        st.session_state["_autosave_loaded"] = False
        st.session_state["project_name_input"] = f"Новый проект {datetime.now():%d.%m %H-%M}"

    with sb1:
        st.button(
            "Открыть", disabled=_picked_save is None, width="stretch",
            on_click=_cb_open, args=(_picked_save["ref"] if _picked_save else None,),
        )
    with sb2:
        st.button(
            "Новый", width="stretch", on_click=_cb_new,
            help="Очистить экран для нового проекта. Сохранённые проекты не удаляются.",
        )
    with sb3:
        with st.popover("", icon=":material/delete:", disabled=_picked_save is None, help="Удалить выбранный проект", width="stretch"):
            st.write(f"Удалить «{_picked}»?")
            if st.button("Удалить", key="_del_confirm"):
                storage_delete(_picked_save["ref"])
                if STORAGE_MODE == "local":
                    pointer = _read_pointer()
                    if (
                        pointer
                        and pointer.get("user_name") == _picked_save["user_name"]
                        and pointer.get("project_name") == _picked_save["project_name"]
                    ):
                        try:
                            POINTER_PATH.unlink(missing_ok=True)
                        except Exception:
                            pass
                st.session_state.pop("_project_picker", None)
                st.session_state.pop("_last_saved_sig", None)
                st.rerun()

    with st.expander("Данные из Excel", icon=":material/table_view:"):
        st.caption(
            "Один файл заполняет все вкладки. Загрузка заменяет данные на экране; "
            "сохранённые проекты не трогает. Проект сохранится под названием из файла."
        )
        if EMPTY_TEMPLATE_BYTES is None:
            EMPTY_TEMPLATE_BYTES = build_template_bytes(None)
        st.download_button(
            "Пустой шаблон", data=EMPTY_TEMPLATE_BYTES, file_name="Шаблон загрузки финмодели.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
            icon=":material/download:",
        )
        try:
            _export_bytes = build_template_bytes(collect_template_data())
        except Exception:
            _export_bytes = None
        if _export_bytes:
            st.download_button(
                "Текущий проект в шаблон", data=_export_bytes,
                file_name=f"{_slugify(st.session_state.get('project_name_input', 'проект'))} — шаблон.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
                icon=":material/download:",
            )
        st.file_uploader("Файл шаблона (.xlsx)", type=["xlsx"], key="_tpl_upload")
        st.button(
            "Загрузить данные", on_click=_cb_import_template, width="stretch",
            disabled=st.session_state.get("_tpl_upload") is None,
        )
    _tpl_msgs = st.session_state.pop("_tpl_msgs", None)
    if _tpl_msgs:
        if _tpl_msgs[0].startswith("Загружено"):
            st.success(_tpl_msgs[0])
        else:
            st.error(_tpl_msgs[0])
        for _w in _tpl_msgs[1:12]:
            st.warning(_w)
        if len(_tpl_msgs) > 12:
            st.warning(f"… и ещё {len(_tpl_msgs) - 12} предупреждений.")

    if not st.session_state["user_name_input"].strip():
        st.caption("⚠️ Введите имя — иначе данные не будут сохраняться.")
    elif st.session_state.get("_autosave_loaded"):
        st.caption("💾 Данные восстановлены из сохранения")
    else:
        st.caption("💾 Автосохранение включено")

    if STORAGE_MODE == "yadisk":
        _cur_ref = current_save_path()
        _writer = _get_ya_writer()
        _err = _writer.errors.get(_cur_ref) if _cur_ref else None
        _err = _err or st.session_state.get("_storage_error")
        if _err:
            st.error(f"Яндекс Диск: ошибка сохранения или чтения — {_err}")
        elif _cur_ref and _cur_ref in _writer.last_ok:
            st.caption(f"☁️ Яндекс Диск — сохранено в {_writer.last_ok[_cur_ref]}")
        else:
            st.caption("☁️ Хранилище: Яндекс Диск")
    else:
        st.caption(
            "🖥️ Хранилище: локальная папка. На Streamlit Cloud данные теряются при "
            "перезапуске — задайте YANDEX_DISK_TOKEN в Secrets."
        )


# ======================================================================
# 5. ШАПКА (на всех вкладках): проект, город, стадия, сценарий, Excel + KPI
# ======================================================================
h1, h2, h3, h4, h5 = st.columns([3, 1.3, 1.1, 1.4, 1.5], vertical_alignment="bottom")
with h1:
    project_name = st.text_input("Проект", key="project_name_input")
with h2:
    project_city = st.text_input("Город", key="project_city_input")
with h3:
    if st.session_state.get("_stage_seg") not in ("МП", "ЭП"):
        st.session_state["_stage_seg"] = "МП" if st.session_state.design_stage == STAGE_MP else "ЭП"
    _stage = st.segmented_control(
        "Стадия", ["МП", "ЭП"], key="_stage_seg",
        help="МП — коробка одной ставкой руб/м² NSA. ЭП — детальный расчет по статьям A–E.",
    )
    if _stage in ("МП", "ЭП"):
        st.session_state.design_stage = STAGE_MP if _stage == "МП" else STAGE_EP
with h4:
    scenario_name = st.selectbox("Сценарий", list(SCENARIOS.keys()), key="scenario_select")
with h5:
    excel_slot = st.empty()
rev_factor = SCENARIOS[scenario_name]["revenue"]
cost_factor = SCENARIOS[scenario_name]["cost"]
is_mp_stage = st.session_state.design_stage == STAGE_MP
kpi_slot = st.container()

# ======================================================================
# 6. ХРАНИЛИЩА ВВОДА (session_state)
# ======================================================================
if "indirect_items" not in st.session_state:
    st.session_state.indirect_items = {}  # ключ раздела -> DataFrame статей (Наименование, Сумма)
if "blocks_df" not in st.session_state:
    st.session_state.blocks_df = generate_default_table()
if "block_rates" not in st.session_state:
    st.session_state.block_rates = {}  # имя жилого блока -> DataFrame ставок (32 статьи A-E)
if "block_g_area" not in st.session_state:
    st.session_state.block_g_area = {}  # имя блока (жилого и паркинга) -> ставки G (площадь участка блока)
if "block_g_length" not in st.session_state:
    st.session_state.block_g_length = {}  # имя блока -> ставки G (сети/кабели)
if "block_z_pct" not in st.session_state:
    st.session_state.block_z_pct = {}  # имя блока -> ставки Z (% от базы блока)
if "block_z_fixed" not in st.session_state:
    st.session_state.block_z_fixed = {}  # имя блока -> статьи Z (прямой ввод суммы)
if "block_infra" not in st.session_state:
    st.session_state.block_infra = {}  # имя ЛЮБОГО урбан-блока -> DataFrame статей инфраструктуры (Z.50.10)
if "block_mp_rate" not in st.session_state:
    st.session_state.block_mp_rate = {}  # имя жилого блока -> ставка коробки, руб/м2 NSA (этап МП)
if "block_elevator_stops" not in st.session_state:
    st.session_state.block_elevator_stops = {}  # имя жилого блока -> кол-во остановок лифтов (статья D.10.10)
if "tep_store" not in st.session_state:
    st.session_state.tep_store = {}  # имя блока -> {поле ТЭП: значение}

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["1. Блоки", "2. Цены", "3. СМР", "4. Инфраструктура и косвенные", "5. Итоги"]
)

# ------------------------------------------------------------------
# ВКЛАДКА 1: урбан-блоки. Единственное место, где добавляются/удаляются блоки —
# остальные вкладки синхронизируются по «Название блока».
# ------------------------------------------------------------------
NAZEMNY_PARKING_COL = "Наземный/Многоуровневый паркинг, м/м"
LOCKED_S_COLS = ["S квартир, м2", "S коммерции 1 эт., м2", "S кладовых, м2"]
with tab1:
    st.subheader(
        "Урбан-блоки",
        help=(
            "Блоки добавляются строкой «+» внизу таблицы, удаляются — выделить строку и корзина. "
            "Площади квартир, коммерции и кладовых (🔒) вводятся в «3. СМР → Данные ЭП». "
            "Наземные м/м — только у блока типа «Наземный/Многоуровневый паркинг»; в паркинге "
            "доступны также коммерция и кладовые — они входят в NSA и получают долю косвенных. "
            "Названия блоков должны быть уникальны."
        ),
    )
    _view = st.session_state.blocks_df.copy()
    for col in MAIN_TABLE_NUMERIC_COLS:
        if col not in _view.columns:
            _view[col] = 0.0
        _view[col] = pd.to_numeric(_view[col], errors="coerce").fillna(0.0)
    _view_res = _view["Тип блока"] == TYPE_RESIDENTIAL
    _view["NSA, м2"] = np.where(
        _view_res,
        _view["S квартир, м2"] + _view["S коммерции 1 эт., м2"] + _view["S кладовых, м2"],
        _view["S коммерции 1 эт., м2"] + _view["S кладовых, м2"],
    )
    column_config = {
        "Название блока": st.column_config.TextColumn("Название", required=True, width="small"),
        "Тип блока": st.column_config.SelectboxColumn(options=BLOCK_TYPES, required=True, width="medium"),
        "NSA, м2": st.column_config.NumberColumn("NSA, м²", format="localized", disabled=True, width="small",
                                                 help="Продаваемая площадь: квартиры + коммерция + кладовые"),
    }
    for col in MAIN_TABLE_NUMERIC_COLS:
        if col in LOCKED_S_COLS:
            column_config[col] = st.column_config.NumberColumn(
                {"S квартир, м2": "Квартиры, м² 🔒", "S коммерции 1 эт., м2": "Коммерция, м² 🔒",
                 "S кладовых, м2": "Кладовые, м² 🔒"}[col], format="localized", disabled=True, width="small",
                help="Вводится в «3. СМР → Данные ЭП» — здесь только просмотр.",
            )
        elif col == NAZEMNY_PARKING_COL:
            column_config[col] = st.column_config.NumberColumn(
                "Наземный, м/м", min_value=0, format="localized",
                help="Только для блока типа «Наземный/Многоуровневый паркинг». В жилом блоке обнуляется.",
            )
        else:
            column_config[col] = st.column_config.NumberColumn("Подземный, м/м", min_value=0, format="localized")

    edited = st.data_editor(
        _view,
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key="blocks_editor",
        column_order=MAIN_TABLE_COLS + ["NSA, м2"],
        column_config=column_config,
    )
    blocks = edited[MAIN_TABLE_COLS].copy()
    for col in MAIN_TABLE_NUMERIC_COLS:
        blocks[col] = pd.to_numeric(blocks[col], errors="coerce").fillna(0.0)
    blocks["Название блока"] = blocks["Название блока"].fillna("").astype(str).str.strip()
    blocks["Тип блока"] = blocks["Тип блока"].fillna(TYPE_RESIDENTIAL)
    blocks = blocks[~((blocks["Название блока"] == "") & (blocks[MAIN_TABLE_NUMERIC_COLS].sum(axis=1) == 0))]
    blocks = blocks.reset_index(drop=True)
    # Уникальные непустые названия — по ним связаны все вкладки
    _seen, _names, _renamed = {}, [], False
    for _i, _n in enumerate(blocks["Название блока"]):
        if not _n:
            _n, _renamed = f"Блок {_i + 1}", True
        base_n, k = _n, 2
        while _n in _seen:
            _n, _renamed = f"{base_n} ({k})", True
            k += 1
        _seen[_n] = True
        _names.append(_n)
    blocks["Название блока"] = _names
    if _renamed:
        st.warning("Пустые и повторяющиеся названия блоков переименованы — названия должны быть уникальны.")
    N_ROWS = len(blocks)

    is_res = (blocks["Тип блока"] == TYPE_RESIDENTIAL)
    is_park = (blocks["Тип блока"] == TYPE_PARKING)
    blocks.loc[is_res, NAZEMNY_PARKING_COL] = 0.0
    st.caption("🔒 — вводится в «3. СМР → Данные ЭП». NSA — расчет, не редактируется.")

all_block_names = list(blocks["Название блока"])
res_block_names = list(blocks.loc[is_res, "Название блока"])
park_block_names = list(blocks.loc[is_park, "Название блока"])
block_type_of = dict(zip(blocks["Название блока"], blocks["Тип блока"]))

# ТЭП-хранилище — по каждому блоку (у паркинга используются коммерция, кладовые,
# площадь участка)
st.session_state.tep_store = {k: v for k, v in st.session_state.tep_store.items() if k in all_block_names}
EP_EXTRA_DEFAULTS = {c: 0.0 for c in EXTRA_EP_PARAM_COLS}
for _name in all_block_names:
    _tep_defaults = {**{tep_col: 0.0 for tep_col in TEP_TO_MAIN_COL}, **EP_EXTRA_DEFAULTS}
    if _name not in st.session_state.tep_store:
        st.session_state.tep_store[_name] = dict(_tep_defaults)
    else:
        for _tep_col, _tep_default in _tep_defaults.items():
            st.session_state.tep_store[_name].setdefault(_tep_col, _tep_default)

# ------------------------------------------------------------------
# ВКЛАДКА 2: цены продаж — строка «База» + блоки
# ------------------------------------------------------------------
PRICE_LABELS = {
    "Цена жилья, руб/м2": "Жильё, ₽/м²",
    "Цена коммерции, руб/м2": "Коммерция, ₽/м²",
    "Цена кладовых, руб/м2": "Кладовые, ₽/м²",
    "Цена подземного м/м, руб": "Подземное м/м, ₽",
    "Цена наземного м/м, руб": "Наземное м/м, ₽",
}
with tab2:
    st.subheader(
        "Цены продаж",
        help=(
            "Строка «База» задаёт цену всем блокам. Цена блока, изменённая вручную, сохраняется "
            "при смене базы. «↺ к базе» — вернуть строку блока к базе. Для паркинга учитываются "
            "наземные м/м, коммерция и кладовые; для жилого блока — всё, кроме наземных м/м."
        ),
    )
    if all_block_names:
        price_store = base_row_table(
            "prices", PRICE_COLS, {c: f"_price_default_{c}" for c in PRICE_COLS},
            all_block_names, "price_editor", PRICE_LABELS,
        )
    else:
        price_store = {}
        st.info("Добавьте хотя бы один блок на Вкладке 1.")

for col in PRICE_COLS:
    if col not in blocks.columns:
        blocks[col] = 0.0
for _i, _name in enumerate(all_block_names):
    for col in PRICE_COLS:
        blocks.at[_i, col] = price_store[_name][col]

# ------------------------------------------------------------------
# Хранилища ставок по блокам: A-E, МП, лифты — жилые блоки; G и Z — все блоки.
# Новый блок наследует ставки ПРЕДЫДУЩЕГО блока того же типа (первый — пустой
# шаблон). Паркинги старых проектов получают пустые G/Z — их цифры не меняются.
# ------------------------------------------------------------------
item_codes_master = list(DEFAULT_RATES_DF["Код"])
code_to_basis = dict(zip(DEFAULT_RATES_DF["Код"], DEFAULT_RATES_DF["_basis"]))
code_to_group = dict(zip(DEFAULT_RATES_DF["Код"], DEFAULT_RATES_DF["Группа"]))
code_to_name = dict(zip(DEFAULT_RATES_DF["Код"], DEFAULT_RATES_DF["Статья затрат"]))
code_to_unit = dict(zip(DEFAULT_RATES_DF["Код"], DEFAULT_RATES_DF["Единица измерения"]))

st.session_state.block_rates = {k: v for k, v in st.session_state.block_rates.items() if k in res_block_names}
st.session_state.block_elevator_stops = {
    k: v for k, v in st.session_state.block_elevator_stops.items() if k in res_block_names
}
for _idx, _name in enumerate(res_block_names):
    _prev = res_block_names[_idx - 1] if _idx > 0 else None
    if _name not in st.session_state.block_rates:
        st.session_state.block_rates[_name] = (
            st.session_state.block_rates[_prev].copy() if _prev in st.session_state.block_rates
            else DEFAULT_RATES_DF.copy()
        )
    if _name not in st.session_state.block_mp_rate:
        st.session_state.block_mp_rate[_name] = st.session_state.block_mp_rate.get(_prev, MP_RATE_DEFAULT)
    if _name not in st.session_state.block_elevator_stops:
        st.session_state.block_elevator_stops[_name] = st.session_state.block_elevator_stops.get(_prev, 0.0)

GZ_STORES = [
    ("block_g_area", DEFAULT_G_AREA_DF), ("block_g_length", DEFAULT_G_LENGTH_DF),
    ("block_z_pct", DEFAULT_Z_PCT_DF), ("block_z_fixed", DEFAULT_Z_FIXED_DF),
]
for _store_name, _default_df in GZ_STORES:
    _d = {k: v for k, v in st.session_state[_store_name].items() if k in all_block_names}
    _prev_same = {}
    for _name in all_block_names:
        _t = block_type_of[_name]
        if _name not in _d:
            _p = _prev_same.get(_t)
            _d[_name] = _d[_p].copy() if _p is not None else _default_df.copy()
        else:
            # статьи, добавленные в каталог позже (старые сохранения) — с нулём
            _miss = _default_df[~_default_df["Код"].isin(_d[_name]["Код"])]
            if len(_miss):
                _d[_name] = pd.concat([_d[_name], _miss], ignore_index=True)
        _prev_same[_t] = _name
    st.session_state[_store_name] = _d

# ------------------------------------------------------------------
# ВКЛАДКА 3: СМР — подвкладки «Данные ЭП», «Коробка A–E», «Паркинг», «G и Z»
# ------------------------------------------------------------------
STOPS_KEY = "__elevator_stops__"
AREA_PARAMS = [
    ("Общая площадь квартир (с летними, с коэф.), м2", "Общая площадь квартир (с коэф.)", "м²"),
    ("Площадь коммерции, м2", "Площадь коммерции", "м²"),
    ("Площадь кладовых в доме, м2", "Площадь кладовых", "м²"),
    ("Площадь участка блока, га", "Площадь участка блока", "га"),
]
EP_ONLY_PARAMS = [
    ("Кол-во секций, шт", "Секций (подъездов)", "шт"),
    ("Кол-во квартир, шт", "Квартир", "шт"),
    ("Кол-во лифтов, шт", "Лифтов", "шт"),
    (STOPS_KEY, "Остановок лифтов (D.10.10)", "шт"),
    ("Площадь застройки, м2", "Площадь застройки", "м²"),
    ("Площадь фасада, м2", "Площадь фасада", "м²"),
    ("Площадь остекления окон, м2", "Остекление окон", "м²"),
    ("Площадь остекления лоджий, м2", "Остекление лоджий", "м²"),
    ("Объем здания ниже 0, м3", "Объём ниже 0", "м³"),
    ("Объем здания выше 0, м3", "Объём выше 0", "м³"),
]
PARKING_TEP_KEYS = {"Площадь коммерции, м2", "Площадь кладовых в доме, м2", "Площадь участка блока, га"}


def _tep_get(n, k):
    if k == STOPS_KEY:
        return _num(st.session_state.block_elevator_stops.get(n, 0.0))
    return _num(st.session_state.tep_store[n].get(k, 0.0))


def _tep_set(n, k, v):
    if k == STOPS_KEY:
        st.session_state.block_elevator_stops[n] = v
    else:
        st.session_state.tep_store[n][k] = v


with tab3:
    sub_ep, sub_box, sub_park, sub_gz = st.tabs(["Данные ЭП", "Коробка A–E", "Паркинг", "G и Z"])

with sub_ep:
    if all_block_names:
        st.markdown("**Площади — нужны на МП и ЭП**")
        transposed_editor(
            AREA_PARAMS, all_block_names, _tep_get, _tep_set, "ep_areas_editor",
            allowed=lambda n, k: block_type_of[n] == TYPE_RESIDENTIAL or k in PARKING_TEP_KEYS,
        )
        if res_block_names and not is_mp_stage:
            st.markdown("**Только ЭП — база статей A–E**")
            transposed_editor(
                EP_ONLY_PARAMS, res_block_names, _tep_get, _tep_set, "ep_only_editor",
                allowed=lambda n, k: True,
            )
        elif is_mp_stage:
            st.caption("На стадии МП параметры для статей A–E не нужны — переключите стадию на ЭП в шапке.")
        st.caption(
            "Строки — параметры, столбцы — блоки. У паркинга вводятся коммерция, кладовые и "
            "площадь участка; остальные ячейки закрыты."
        )
    else:
        st.info("Добавьте хотя бы один блок на Вкладке 1.")

# Данные ЭП — источник истины для площадей и параметров блоков
for col in EXTRA_EP_PARAM_COLS + PARAM_COLS:
    if col not in blocks.columns:
        blocks[col] = 0.0
for _i, _name in enumerate(all_block_names):
    store_row = st.session_state.tep_store[_name]
    if is_res[_i]:
        for tep_col, main_col in TEP_TO_MAIN_COL.items():
            blocks.at[_i, main_col] = store_row[tep_col]
        for col in EXTRA_EP_PARAM_COLS:
            blocks.at[_i, col] = store_row[col]
    else:
        blocks.at[_i, "S квартир, м2"] = 0.0
        blocks.at[_i, "S коммерции 1 эт., м2"] = store_row["Площадь коммерции, м2"]
        blocks.at[_i, "S кладовых, м2"] = store_row["Площадь кладовых в доме, м2"]
        blocks.at[_i, "Площадь участка блока, га"] = store_row["Площадь участка блока, га"]

st.session_state.blocks_df = blocks[MAIN_TABLE_COLS].copy()

# NSA: жилой блок — квартиры + коммерция + кладовые; паркинг — коммерция + кладовые.
# На NSA распределяется пул косвенных расходов.
nsa = np.where(
    is_res,
    blocks["S квартир, м2"] + blocks["S коммерции 1 эт., м2"] + blocks["S кладовых, м2"],
    blocks["S коммерции 1 эт., м2"] + blocks["S кладовых, м2"],
).astype(float)
blocks["NSA, м2"] = nsa

elevator_stops_arr = np.array([
    float(st.session_state.block_elevator_stops.get(blocks["Название блока"].iloc[i], 0.0)) if is_res[i] else 0.0
    for i in range(N_ROWS)
])
qty_by_basis = {
    b: (elevator_stops_arr if b == "ELEVATOR_STOPS" else get_basis_values(b, blocks, nsa))
    for b in set(code_to_basis.values())
}


def _block_group_sum(name: str, group: str) -> float:
    i = all_block_names.index(name)
    df = st.session_state.block_rates[name]
    rates = pd.to_numeric(df.set_index("Код")["Ставка, руб/ед."], errors="coerce").fillna(0.0)
    return float(sum(
        qty_by_basis[code_to_basis[c]][i] * rates.get(c, 0.0)
        for c in item_codes_master if code_to_group[c] == group
    ))


with sub_box:
    box_card = None
    if not res_block_names:
        st.info("Добавьте хотя бы один «Жилой блок» на Вкладке 1.")
    elif is_mp_stage:
        selected_block = fixed_selector("Блок (карточка справа)", res_block_names, "smr_block_selector")
        box_left, box_right = st.columns([3, 1.1], gap="medium")
        with box_left:
            st.subheader(
                "Коробка — укрупнённо (МП)",
                help="Ставка руб/м² NSA. Строка «База» — для всех жилых блоков, корректировка — по блоку.",
            )
            korobka_store = base_row_table(
                "mp_korobka", ["Ставка, руб/м2 NSA"], {"Ставка, руб/м2 NSA": "_mp_korobka_default"},
                res_block_names, "mp_korobka_editor", {"Ставка, руб/м2 NSA": "Ставка, ₽/м² NSA"},
            )
            for _name in res_block_names:
                st.session_state.block_mp_rate[_name] = korobka_store[_name]["Ставка, руб/м2 NSA"]
        box_card = box_right.container(border=True)
    else:
        sel_row = st.columns([4, 1.3], vertical_alignment="bottom")
        with sel_row[0]:
            selected_block = fixed_selector("Блок", res_block_names, "smr_block_selector")
        with sel_row[1]:
            with st.popover("Копировать ставки в блок…", width="stretch"):
                _targets = st.multiselect(
                    "В какие блоки", [n for n in res_block_names if n != selected_block], key="copy_box_targets",
                )
                if st.button("Копировать ставки A–E", key="copy_box_btn", disabled=not _targets):
                    for _t in _targets:
                        st.session_state.block_rates[_t] = st.session_state.block_rates[selected_block].copy()
                    st.session_state.pop("copy_box_targets", None)
                    st.rerun()
        box_left, box_right = st.columns([3, 1.1], gap="medium")
        with box_left:
            _groups = list(GROUP_LABELS.keys())
            _gsums = {g: _block_group_sum(selected_block, g) for g in _groups}
            if st.session_state.get("smr_group_selector") not in _groups:
                st.session_state["smr_group_selector"] = _groups[0]
            _grp = st.segmented_control(
                "Группа работ", _groups, key="smr_group_selector",
                format_func=lambda g: f"{g} · {GROUP_LABELS[g]} — {fmt_mln(_gsums[g], unit=False)}",
            ) or _groups[0]
            _bi = all_block_names.index(selected_block)
            _rates_df = st.session_state.block_rates[selected_block]
            _g_codes = [c for c in item_codes_master if code_to_group[c] == _grp]
            _rate_map = pd.to_numeric(_rates_df.set_index("Код")["Ставка, руб/ед."], errors="coerce").fillna(0.0)
            _gdf = pd.DataFrame([{
                "Код": c, "Статья": code_to_name[c], "Ед.": code_to_unit[c],
                "База": float(qty_by_basis[code_to_basis[c]][_bi]),
                "Ставка, руб/ед.": float(_rate_map.get(c, 0.0)),
                "Сумма, руб": float(qty_by_basis[code_to_basis[c]][_bi] * _rate_map.get(c, 0.0)),
            } for c in _g_codes])
            _ged = st.data_editor(
                _gdf, key=f"rates_{selected_block}_{_grp}", hide_index=True, num_rows="fixed",
                width="stretch",
                disabled=["Код", "Статья", "Ед.", "База", "Сумма, руб"],
                column_config={
                    "База": st.column_config.NumberColumn(format="localized", help="Из «Данных ЭП»"),
                    "Ставка, руб/ед.": st.column_config.NumberColumn(min_value=0, format="localized"),
                    "Сумма, руб": st.column_config.NumberColumn(format="localized"),
                },
            )
            _changed = False
            _new_rates = _rates_df.copy()
            for _, _r in _ged.iterrows():
                _v = _num(_r["Ставка, руб/ед."])
                if _v != float(_rate_map.get(_r["Код"], 0.0)):
                    _new_rates.loc[_new_rates["Код"] == _r["Код"], "Ставка, руб/ед."] = _v
                    _changed = True
            if _changed:
                st.session_state.block_rates[selected_block] = _new_rates
                st.rerun()
            st.caption("Базы статей — из «Данных ЭП»; остановки лифтов (D.10.10) — там же.")
        box_card = box_right.container(border=True)

with sub_park:
    st.subheader(
        "Ставки СМР паркинга",
        help=(
            "Подземный м/м — для жилых блоков (подземный паркинг в составе блока). Наземный м/м, "
            "коммерция и кладовые — для блоков-паркингов. У жилых блоков коммерция и кладовые "
            "считаются в коробке A–E — эти колонки для них не используются."
        ),
    )
    if all_block_names:
        parking_rate_store = base_row_table(
            "smr_parking", SMR_PARKING_RATE_COLS, {c: f"_smr_parking_default_{c}" for c in SMR_PARKING_RATE_COLS},
            all_block_names, "parking_rate_editor",
            {
                "Ставка СМР подземного м/м, руб": "Подземный, ₽/м/м",
                "Ставка СМР наземного м/м, руб": "Наземный / многоур., ₽/м/м",
                "Ставка СМР коммерции паркинга, руб/м2": "Коммерция паркинга, ₽/м²",
                "Ставка СМР кладовых паркинга, руб/м2": "Кладовые паркинга, ₽/м²",
            },
        )
    else:
        parking_rate_store = {}
        st.info("Добавьте хотя бы один блок на Вкладке 1.")
    park_card = st.container()

for col in SMR_PARKING_RATE_COLS:
    if col not in blocks.columns:
        blocks[col] = 0.0
for _i, _name in enumerate(all_block_names):
    for col in SMR_PARKING_RATE_COLS:
        blocks.at[_i, col] = parking_rate_store[_name][col]

with sub_gz:
    gz_strip = None
    if not all_block_names:
        st.info("Добавьте хотя бы один блок на Вкладке 1.")
    else:
        gz_row = st.columns([4, 1.3], vertical_alignment="bottom")
        with gz_row[0]:
            gz_block = fixed_selector("Блок", all_block_names, "gz_block_selector")
        with gz_row[1]:
            with st.popover("Копировать ставки в блок…", width="stretch"):
                _gz_targets = st.multiselect(
                    "В какие блоки", [n for n in all_block_names if n != gz_block], key="copy_gz_targets",
                )
                if st.button("Копировать ставки G и Z", key="copy_gz_btn", disabled=not _gz_targets):
                    for _t in _gz_targets:
                        for _store_name, _ in GZ_STORES:
                            st.session_state[_store_name][_t] = st.session_state[_store_name][gz_block].copy()
                    st.session_state.pop("copy_gz_targets", None)
                    st.rerun()
        gz_strip = st.container()
        _plot = float(blocks.loc[blocks["Название блока"] == gz_block, "Площадь участка блока, га"].iloc[0])
        gA, gB = st.columns(2, gap="medium")
        with gA:
            st.markdown("**G · площадь участка**")
            st.caption(f"База — площадь участка блока: {_plot:.2f} га (из «Данных ЭП»)".replace(".", ","))
            st.session_state.block_g_area[gz_block] = st.data_editor(
                st.session_state.block_g_area[gz_block], width="stretch", num_rows="fixed",
                hide_index=True, key=f"g_area_editor_{gz_block}",
                disabled=["Код", "Статья затрат", "Единица измерения"],
                column_config={
                    "Единица измерения": None,
                    "Ставка, руб/ед.": st.column_config.NumberColumn("Ставка, ₽/га", min_value=0, format="localized"),
                },
            )
        with gB:
            st.markdown("**G · сети и кабели**")
            st.caption("Длины вводятся вручную")
            st.session_state.block_g_length[gz_block] = st.data_editor(
                st.session_state.block_g_length[gz_block], width="stretch", num_rows="fixed",
                hide_index=True, key=f"g_length_editor_{gz_block}",
                disabled=["Код", "Статья затрат", "Единица измерения"],
                column_config={
                    "Единица измерения": None,
                    "Кол-во": st.column_config.NumberColumn("Кол-во, п.м.", min_value=0, format="%.1f"),
                    "Ставка, руб/ед.": st.column_config.NumberColumn("Ставка, ₽/п.м.", min_value=0, format="localized"),
                },
            )
        zC, zD = st.columns(2, gap="medium")
        with zC:
            st.markdown("**Z · % от базы**")
            _zbase_txt = (
                "коробка A–E + G + подземный паркинг" if block_type_of[gz_block] == TYPE_RESIDENTIAL
                else "СМР паркинга (м/м + коммерция + кладовые) + G"
            )
            st.caption(f"База блока = {_zbase_txt}")
            st.session_state.block_z_pct[gz_block] = st.data_editor(
                st.session_state.block_z_pct[gz_block], width="stretch", num_rows="fixed",
                hide_index=True, key=f"z_pct_editor_{gz_block}",
                disabled=["Код", "Статья затрат"],
                column_config={
                    "Ставка, доля от СМР+G": st.column_config.NumberColumn(
                        "Ставка, доля от базы", min_value=0, max_value=1, format="%.3f",
                    ),
                },
            )
        with zD:
            st.markdown("**Z · прямой ввод суммы**")
            st.caption("По объектам-аналогам")
            st.session_state.block_z_fixed[gz_block] = st.data_editor(
                st.session_state.block_z_fixed[gz_block], width="stretch", num_rows="fixed",
                hide_index=True, key=f"z_fixed_editor_{gz_block}",
                disabled=["Код", "Статья затрат"],
                column_config={"Сумма, руб": st.column_config.NumberColumn(min_value=0, format="localized")},
            )

# ------------------------------------------------------------------
# Расчет себестоимости коробки (жилые блоки): у каждого блока свой каталог ставок
# ------------------------------------------------------------------
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

if is_mp_stage:
    smr_korobka_raw = np.array([
        nsa[i] * float(st.session_state.block_mp_rate.get(blocks["Название блока"].iloc[i], 0.0)) if is_res[i] else 0.0
        for i in range(N_ROWS)
    ])
else:
    smr_korobka_raw = np.sum(list(item_cost_matrix.values()), axis=0) if item_cost_matrix and N_ROWS > 0 else np.zeros(N_ROWS)
smr_korobka_scaled = smr_korobka_raw * cost_factor
blocks["Себестоимость коробки (методика)"] = smr_korobka_scaled
blocks["Эффективная ставка коробки, руб/м2"] = np.where(nsa > 0, smr_korobka_scaled / np.where(nsa > 0, nsa, 1), 0.0)

# СМР паркинга (без сценария): жилой блок — подземные м/м; блок-паркинг —
# наземные м/м + коммерция + кладовые по ставкам подвкладки «Паркинг»
underground_cost_raw = (blocks["Подземный паркинг, м/м"] * blocks["Ставка СМР подземного м/м, руб"]).to_numpy(dtype=float)
park_block_smr_raw = np.where(
    is_park,
    blocks[NAZEMNY_PARKING_COL] * blocks["Ставка СМР наземного м/м, руб"]
    + blocks["S коммерции 1 эт., м2"] * blocks["Ставка СМР коммерции паркинга, руб/м2"]
    + blocks["S кладовых, м2"] * blocks["Ставка СМР кладовых паркинга, руб/м2"],
    0.0,
).astype(float)
parking_cost_block = np.where(is_res, underground_cost_raw, park_block_smr_raw)

# ------------------------------------------------------------------
# Наружные работы (G) и прочие затраты (Z) — ПО КАЖДОМУ БЛОКУ (жилому и паркингу)
# ------------------------------------------------------------------
plot_area_col = blocks["Площадь участка блока, га"].to_numpy(dtype=float)
g_area_total = np.zeros(N_ROWS)
g_length_total = np.zeros(N_ROWS)
z_fixed_total = np.zeros(N_ROWS)
z_pct_total = np.zeros(N_ROWS)
for i in range(N_ROWS):
    name = blocks["Название блока"].iloc[i]
    g_area_block = st.session_state.block_g_area.get(name, DEFAULT_G_AREA_DF)
    g_area_total[i] = plot_area_col[i] * float(pd.to_numeric(g_area_block["Ставка, руб/ед."], errors="coerce").fillna(0.0).sum())
    g_length_block = st.session_state.block_g_length.get(name, DEFAULT_G_LENGTH_DF)
    g_qty = pd.to_numeric(g_length_block["Кол-во"], errors="coerce").fillna(0.0)
    g_rate = pd.to_numeric(g_length_block["Ставка, руб/ед."], errors="coerce").fillna(0.0)
    g_length_total[i] = float((g_qty * g_rate).sum())
    z_fixed_block = st.session_state.block_z_fixed.get(name, DEFAULT_Z_FIXED_DF)
    z_fixed_total[i] = float(pd.to_numeric(z_fixed_block["Сумма, руб"], errors="coerce").fillna(0.0).sum())

g_block_total = g_area_total + g_length_total
# База % статей Z: жилой — коробка + G + подземный паркинг; паркинг — СМР паркинга + G
zg_base = smr_korobka_raw + g_block_total + parking_cost_block
for i in range(N_ROWS):
    name = blocks["Название блока"].iloc[i]
    z_pct_block = st.session_state.block_z_pct.get(name, DEFAULT_Z_PCT_DF)
    z_pct_total[i] = zg_base[i] * float(pd.to_numeric(z_pct_block["Ставка, доля от СМР+G"], errors="coerce").fillna(0.0).sum())

z_block_total = z_pct_total + z_fixed_total
g_block_total_scaled = g_block_total * cost_factor
z_block_total_scaled = z_block_total * cost_factor
blocks["Наружные работы блока (G), руб"] = g_block_total_scaled
blocks["Прочие затраты блока (Z), руб"] = z_block_total_scaled

# ------------------------------------------------------------------
# ВКЛАДКА 4: косвенные расходы проекта + инфраструктура Z.50.10 по блокам
# ------------------------------------------------------------------
st.session_state.block_infra = {k: v for k, v in st.session_state.block_infra.items() if k in all_block_names}
for _name in all_block_names:
    _existing_infra = st.session_state.block_infra.get(_name)
    if _existing_infra is None or "Сумма, руб" not in _existing_infra.columns:
        st.session_state.block_infra[_name] = EMPTY_INFRA_DF.copy()

with tab4:
    ind_col, infra_col = st.columns(2, gap="large")
    with ind_col:
        st.subheader(
            "Косвенные расходы проекта",
            help=(
                "Только затраты, общие на весь участок. Распределяются по NSA — на жилые блоки и на "
                "коммерцию/кладовые паркингов. Локальные затраты блока — в «3. СМР → G и Z»."
            ),
        )
        cost_land = render_indirect_category("land", "Земля", 0.0)
        cost_infra = render_indirect_category("infra", "Магистральные сети (на весь участок)", 0.0)
        cost_landscape = render_indirect_category("landscape", "Благоустройство мест общего пользования", 0.0)
        cost_social = render_indirect_category("social", "Социальные объекты (школы/сады)", 0.0)
        cost_soft = render_indirect_category("soft", "Прочие Soft Costs", 0.0)
        indirect_pool_sidebar = cost_land + cost_infra + cost_landscape + cost_social + cost_soft
        st.markdown(f"#### Итого пул: {fmt_money(indirect_pool_sidebar)} ₽")
    with infra_col:
        st.subheader(
            "Инфраструктура Z.50.10",
            help=(
                "«Покупка объекта недвижимости (Инфраструктура)» — по каждому блоку, статьи свободные "
                "(плейхаб, парк, дороги и т.п.), до 10 на блок. Итог прибавляется к прямым затратам "
                "блока и выгружается в Excel строкой Z5010."
            ),
        )
        if all_block_names:
            selected_infra_block = fixed_selector("Блок", all_block_names, "infra_block_selector")
            infra_hdr = st.empty()
            infra_edited = st.data_editor(
                st.session_state.block_infra[selected_infra_block],
                width="stretch",
                num_rows="dynamic",
                hide_index=True,
                key=f"infra_editor_{selected_infra_block}",
                column_config={
                    "Статья затрат": st.column_config.TextColumn(),
                    "Сумма, руб": st.column_config.NumberColumn(min_value=0, format="localized"),
                },
            )
            if len(infra_edited) > MAX_INFRA_ITEMS_PER_BLOCK:
                st.warning(f"Максимум {MAX_INFRA_ITEMS_PER_BLOCK} статей на блок — лишние строки не учитываются.")
                infra_edited = infra_edited.iloc[:MAX_INFRA_ITEMS_PER_BLOCK].reset_index(drop=True)
            st.session_state.block_infra[selected_infra_block] = infra_edited
            _infra_sel_total = float(pd.to_numeric(infra_edited["Сумма, руб"], errors="coerce").fillna(0.0).sum())
            infra_hdr.markdown(f"**Итого {selected_infra_block}: {fmt_money(_infra_sel_total)} ₽** · строка «+» — новая статья")
            infra_summary_slot = st.container()
        else:
            infra_summary_slot = None
            st.info("Добавьте хотя бы один блок на Вкладке 1.")

infra_total = np.array([
    float(pd.to_numeric(
        st.session_state.block_infra.get(blocks["Название блока"].iloc[i], EMPTY_INFRA_DF)["Сумма, руб"],
        errors="coerce",
    ).fillna(0.0).sum())
    for i in range(N_ROWS)
])
blocks["Инфраструктура блока (Z5010), руб"] = infra_total * cost_factor
if infra_summary_slot is not None:
    with infra_summary_slot:
        st.markdown("**Z.50.10 по всем блокам**")
        st.dataframe(
            pd.DataFrame({"Блок": blocks["Название блока"], "Z.50.10, руб": infra_total}),
            width="stretch", hide_index=True,
            column_config={"Z.50.10, руб": st.column_config.NumberColumn(format="localized")},
        )

# ======================================================================
# 7. РАСЧЕТ ЭКОНОМИКИ
# ======================================================================
indirect_pool_total = indirect_pool_sidebar

total_nsa = nsa.sum()
share = np.zeros_like(nsa, dtype=float)
if total_nsa > 0:
    share = nsa / total_nsa
blocks["Доля аллокации"] = share

indirect_pool_scenario = indirect_pool_total * cost_factor
blocks["Аллоцированные затраты"] = blocks["Доля аллокации"] * indirect_pool_scenario

# Прямые затраты: жилой блок = коробка + G + Z + подземный паркинг + Z.50.10;
# блок-паркинг = СМР паркинга (м/м + коммерция + кладовые) + G + Z + Z.50.10.
direct_res = (
    blocks["Себестоимость коробки (методика)"] + blocks["Наружные работы блока (G), руб"]
    + blocks["Прочие затраты блока (Z), руб"]
    + blocks["Подземный паркинг, м/м"] * blocks["Ставка СМР подземного м/м, руб"] * cost_factor
    + blocks["Инфраструктура блока (Z5010), руб"]
)
direct_park = (
    park_block_smr_raw * cost_factor
    + blocks["Наружные работы блока (G), руб"] + blocks["Прочие затраты блока (Z), руб"]
    + blocks["Инфраструктура блока (Z5010), руб"]
)
blocks["Прямые затраты"] = np.where(is_res, direct_res, direct_park)
blocks["Полные затраты"] = blocks["Прямые затраты"] + blocks["Аллоцированные затраты"]

revenue_res = (
    blocks["S квартир, м2"] * blocks["Цена жилья, руб/м2"]
    + blocks["S коммерции 1 эт., м2"] * blocks["Цена коммерции, руб/м2"]
    + blocks["S кладовых, м2"] * blocks["Цена кладовых, руб/м2"]
    + blocks["Подземный паркинг, м/м"] * blocks["Цена подземного м/м, руб"]
) * rev_factor
revenue_park = (
    blocks[NAZEMNY_PARKING_COL] * blocks["Цена наземного м/м, руб"]
    + blocks["S коммерции 1 эт., м2"] * blocks["Цена коммерции, руб/м2"]
    + blocks["S кладовых, м2"] * blocks["Цена кладовых, руб/м2"]
) * rev_factor
blocks["Выручка"] = np.where(is_res, revenue_res, revenue_park)

blocks["Валовая прибыль"] = blocks["Выручка"] - blocks["Полные затраты"]
blocks["Валовая рентабельность"] = np.where(blocks["Выручка"] > 0, blocks["Валовая прибыль"] / blocks["Выручка"], 0.0)

revenue_components = {
    "Жилье": (blocks.loc[is_res, "S квартир, м2"] * blocks.loc[is_res, "Цена жилья, руб/м2"] * rev_factor).sum(),
    "Коммерция": (blocks["S коммерции 1 эт., м2"] * blocks["Цена коммерции, руб/м2"] * rev_factor).sum(),
    "Кладовые": (blocks["S кладовых, м2"] * blocks["Цена кладовых, руб/м2"] * rev_factor).sum(),
    "Подземные м/м": (blocks.loc[is_res, "Подземный паркинг, м/м"] * blocks.loc[is_res, "Цена подземного м/м, руб"] * rev_factor).sum(),
    "Наземные/Многоур. паркинги": (
        blocks.loc[is_park, NAZEMNY_PARKING_COL] * blocks.loc[is_park, "Цена наземного м/м, руб"] * rev_factor
    ).sum(),
}

total_revenue = blocks["Выручка"].sum()
total_cost = blocks["Полные затраты"].sum()
total_profit = total_revenue - total_cost
avg_margin = (total_profit / total_revenue) if total_revenue > 0 else 0.0
no_residential_warning = (blocks.shape[0] > 0) and (total_nsa == 0) and (indirect_pool_total > 0)

parking_smr_scaled = np.where(is_res, underground_cost_raw, park_block_smr_raw) * cost_factor

# ======================================================================
# 8. ЗАПОЛНЕНИЕ МЕСТ НА ЭКРАНЕ: KPI, карточки блоков, Итоги
# ======================================================================
with kpi_slot:
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Выручка", fmt_mln(total_revenue))
    k2.metric("Полные затраты", fmt_mln(total_cost))
    k3.metric("Валовая прибыль", fmt_mln(total_profit))
    k4.metric("Валовая рентабельность", f"{avg_margin * 100:.1f} %".replace(".", ","))
    k5.metric("NSA проекта", f"{fmt_money(total_nsa)} м²")
    if no_residential_warning:
        st.warning("NSA проекта = 0 — пул косвенных расходов не на что распределить, он не учтен в итогах.")


def _block_cost_lines(i: int) -> list:
    if is_res[i]:
        return [
            ("Коробка A–E" if not is_mp_stage else "Коробка (МП)", smr_korobka_scaled[i]),
            ("Подземный паркинг", underground_cost_raw[i] * cost_factor),
            ("Наружные работы G", g_block_total_scaled[i]),
            ("Прочие Z", z_block_total_scaled[i]),
            ("Инфраструктура Z.50.10", blocks["Инфраструктура блока (Z5010), руб"].iloc[i]),
        ]
    return [
        ("СМР паркинга", park_block_smr_raw[i] * cost_factor),
        ("Наружные работы G", g_block_total_scaled[i]),
        ("Прочие Z", z_block_total_scaled[i]),
        ("Инфраструктура Z.50.10", blocks["Инфраструктура блока (Z5010), руб"].iloc[i]),
    ]


if box_card is not None and res_block_names:
    _i = all_block_names.index(selected_block)
    with box_card:
        st.markdown(f"**{selected_block} · прямые затраты, млн ₽**")
        st.dataframe(
            pd.DataFrame([{"Статья": a, "млн ₽": b / 1e6} for a, b in _block_cost_lines(_i)]),
            hide_index=True, width="stretch",
            column_config={"млн ₽": st.column_config.NumberColumn(format="%.1f")},
        )
        st.metric("Итого прямые", fmt_mln(blocks["Прямые затраты"].iloc[_i]))
        st.caption(
            f"на 1 м² NSA: {fmt_money(blocks['Прямые затраты'].iloc[_i] / nsa[_i]) if nsa[_i] > 0 else '—'} ₽"
        )

with park_card:
    if park_block_names:
        st.markdown("**Блоки-паркинги · СМР, руб** (со сценарием)")
        _pc = []
        for i in range(N_ROWS):
            if not is_park[i]:
                continue
            _pc.append({
                "Блок": blocks["Название блока"].iloc[i],
                "Машиноместа": blocks[NAZEMNY_PARKING_COL].iloc[i] * blocks["Ставка СМР наземного м/м, руб"].iloc[i] * cost_factor,
                "Коммерция": blocks["S коммерции 1 эт., м2"].iloc[i] * blocks["Ставка СМР коммерции паркинга, руб/м2"].iloc[i] * cost_factor,
                "Кладовые": blocks["S кладовых, м2"].iloc[i] * blocks["Ставка СМР кладовых паркинга, руб/м2"].iloc[i] * cost_factor,
                "СМР паркинга": park_block_smr_raw[i] * cost_factor,
            })
        st.dataframe(
            pd.DataFrame(_pc), hide_index=True, width="stretch",
            column_config={c: st.column_config.NumberColumn(format="localized") for c in ["Машиноместа", "Коммерция", "Кладовые", "СМР паркинга"]},
        )
        st.caption("Наружные работы G и прочие Z паркинга — на подвкладке «G и Z».")

if gz_strip is not None:
    _gi = all_block_names.index(gz_block)
    with gz_strip:
        s1, s2, s3, s4, s5 = st.columns(5)
        s1.metric(f"{gz_block} · прямые затраты", fmt_mln(blocks["Прямые затраты"].iloc[_gi]))
        s2.metric("Наружные G", fmt_mln(g_block_total_scaled[_gi]))
        s3.metric("Прочие Z", fmt_mln(z_block_total_scaled[_gi]))
        s4.metric("База % Z (без сценария)", fmt_mln(zg_base[_gi]))
        s5.metric("Z.50.10", fmt_mln(blocks["Инфраструктура блока (Z5010), руб"].iloc[_gi]))

with tab5:
    st.subheader(
        f"{project_name}, {project_city} — сценарий «{scenario_name}»",
        help="Валовая прибыль и валовая рентабельность. Налоги и кредиты не учитываются.",
    )
    tbl_cols = [
        "Название блока", "Тип блока", "NSA, м2", "Доля аллокации",
        "Прямые затраты", "Аллоцированные затраты", "Полные затраты",
        "Выручка", "Валовая прибыль", "Валовая рентабельность",
    ]
    _money_cols = ["Прямые затраты", "Аллоцированные затраты", "Полные затраты", "Выручка", "Валовая прибыль"]
    _res_view = blocks[tbl_cols].copy()
    _res_view["Тип блока"] = np.where(is_res, "Жилой", "Паркинг")
    _res_view["NSA, м2"] = _res_view["NSA, м2"].round(0)
    for _c in _money_cols:
        _res_view[_c] = _res_view[_c] / 1e6
    _tot = {"Название блока": "Итого", "Тип блока": "", "NSA, м2": round(total_nsa),
            "Доля аллокации": 1.0 if total_nsa > 0 else 0.0,
            **{_c: _res_view[_c].sum() for _c in _money_cols}, "Валовая рентабельность": avg_margin}
    _res_view = pd.concat([_res_view, pd.DataFrame([_tot])], ignore_index=True)
    def _ru(v, d=1):
        return f"{v:,.{d}f}".replace(",", " ").replace(".", ",")

    _last = len(_res_view) - 1
    _styler = (
        _res_view.style
        .format({
            "NSA, м2": lambda v: _ru(v, 0),
            "Доля аллокации": lambda v: _ru(v * 100, 1) + " %",
            "Валовая рентабельность": lambda v: _ru(v * 100, 1) + " %",
            **{_c: (lambda v: _ru(v, 1)) for _c in _money_cols},
        })
        .apply(lambda r: ["background-color: #EEF7EC; color: #2E7D27" if r.name == _last else "" for _ in r], axis=1)
        .map(lambda v: "color: #AE4B67" if isinstance(v, (int, float)) and v < 0 else "",
             subset=["Валовая прибыль", "Валовая рентабельность"])
    )
    st.dataframe(
        _styler, width="stretch", hide_index=True,
        column_config={
            "Название блока": st.column_config.TextColumn("Блок"),
            "Тип блока": st.column_config.TextColumn("Тип"),
            "NSA, м2": st.column_config.NumberColumn("NSA, м²"),
            "Доля аллокации": st.column_config.NumberColumn("Доля"),
            "Валовая рентабельность": st.column_config.NumberColumn("Вал. рентаб."),
            "Прямые затраты": st.column_config.NumberColumn("Прямые, млн ₽"),
            "Аллоцированные затраты": st.column_config.NumberColumn("Косвенные, млн ₽"),
            "Полные затраты": st.column_config.NumberColumn("Полные, млн ₽"),
            "Выручка": st.column_config.NumberColumn("Выручка, млн ₽"),
            "Валовая прибыль": st.column_config.NumberColumn("Вал. прибыль, млн ₽"),
        },
    )
    _chart_layout = dict(template="plotly_white+talan", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", height=300, margin=dict(t=48, b=24, l=8, r=8),
                         legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="left", x=0))
    c1, c2 = st.columns(2)
    with c1:
        fig_rev_cost = go.Figure()
        fig_rev_cost.add_bar(name="Выручка", x=blocks["Название блока"], y=blocks["Выручка"] / 1e6, marker_color=f"#{COLOR_REVENUE}")
        fig_rev_cost.add_bar(name="Полные затраты", x=blocks["Название блока"], y=blocks["Полные затраты"] / 1e6, marker_color=f"#{COLOR_COST}")
        fig_rev_cost.update_layout(title="Выручка и затраты по блокам, млн ₽", barmode="group", **_chart_layout)
        st.plotly_chart(fig_rev_cost, width="stretch", theme=None)
    with c2:
        margin_colors = [f"#{COLOR_POSITIVE}" if v >= 0 else f"#{COLOR_NEGATIVE}" for v in blocks["Валовая рентабельность"]]
        fig_margin = go.Figure(go.Bar(
            x=blocks["Название блока"], y=blocks["Валовая рентабельность"] * 100, marker_color=margin_colors,
            text=[f"{v * 100:.1f} %".replace(".", ",") for v in blocks["Валовая рентабельность"]], textposition="outside",
        ))
        fig_margin.update_layout(title="Валовая рентабельность по блокам, %", yaxis_title="%", **_chart_layout)
        fig_margin.add_hline(y=0, line_color="#898781", line_width=1)
        st.plotly_chart(fig_margin, width="stretch", theme=None)

    def _stack_chart(title: str, parts: dict, colors: list):
        total = sum(v for v in parts.values() if v > 0)
        fig = go.Figure()
        for (lbl, v), col in zip(parts.items(), colors):
            pct = (v / total * 100) if total > 0 else 0.0
            fig.add_bar(name=f"{lbl} {pct:.0f}%", x=[pct], y=[""], orientation="h", marker_color=f"#{col}",
                        hovertemplate=f"{lbl}: {fmt_money(v)} ₽ ({pct:.1f}%)<extra></extra>")
        fig.update_layout(template="plotly_white+talan", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", title=title, barmode="stack", height=190, margin=dict(t=48, b=8, l=8, r=8),
                          xaxis=dict(range=[0, 100], ticksuffix="%"),
                          legend=dict(orientation="h", yanchor="top", y=-0.25, xanchor="left", x=0))
        return fig

    c3, c4 = st.columns(2)
    with c3:
        st.plotly_chart(_stack_chart("Структура выручки", revenue_components, PIE_COLORS), width="stretch", theme=None)
    with c4:
        cost_parts = {
            "Коробка": float(smr_korobka_scaled.sum()),
            "G": float(g_block_total_scaled.sum()),
            "Z": float(z_block_total_scaled.sum()),
            "Паркинг": float(parking_smr_scaled.sum()),
            "Z.50.10": float(blocks["Инфраструктура блока (Z5010), руб"].sum()),
            "Косвенные": float(blocks["Аллоцированные затраты"].sum()),
        }
        st.plotly_chart(
            _stack_chart("Структура полных затрат", cost_parts,
                         [PALETTE["green_dark"], PALETTE["green"], PALETTE["green2"], PALETTE["bordo"], PALETTE["green3"], PALETTE["gray"]]),
            width="stretch", theme=None,
        )
    if not is_mp_stage and res_block_names:
        with st.expander("Детализация коробки по статьям A–E и блокам"):
            st.caption("Значения по статьям — на базовых ставках (без коэффициента сценария).")
            detail_df = pd.DataFrame(item_cost_matrix, index=blocks["Название блока"]).T
            detail_df.insert(0, "Статья затрат", [code_to_name[c] for c in detail_df.index])
            st.dataframe(detail_df, width="stretch")

# ======================================================================
# 10. ЭКСПОРТ В EXCEL (openpyxl) — ЖИВЫЕ ФОРМУЛЫ, 2 ЛИСТА, ВСТРОЕННЫЕ ГРАФИКИ
# ======================================================================
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
THICK_TOP_GREEN = Side(style="medium", color=XL_GREEN)
HEADER_FILL = PatternFill(start_color=XL_GREEN, end_color=XL_GREEN, fill_type="solid")
HEADER_FONT = Font(name=XL_FONT, bold=True, color=XL_WHITE, size=10)
TITLE_FONT = Font(name=XL_FONT, bold=True, size=14, color=XL_GREEN)
SUBTITLE_FONT = Font(name=XL_FONT, size=11, color=XL_TEXT2)
PARAM_FONT = Font(name=XL_FONT, bold=True, color=XL_TEXT2)
TOTAL_FILL = PatternFill(start_color=XL_GREEN3, end_color=XL_GREEN3, fill_type="solid")
TOTAL_FONT = Font(name=XL_FONT, bold=True, color=XL_TEXT)
ZEBRA_FILL = PatternFill(start_color=XL_GREEN4, end_color=XL_GREEN4, fill_type="solid")
WHITE_FILL = PatternFill(start_color=XL_WHITE, end_color=XL_WHITE, fill_type="solid")
DATA_FONT = Font(name=XL_FONT, size=10, color=XL_TEXT)
MONEY_FMT = "#,##0"
PERCENT_FMT = "0.0%"
MONEY_MM_FMT = '#,##0.0,, "млн ₽"'


def zebra(i: int) -> PatternFill:
    """Чередующаяся заливка строк данных по индексу i (0-based)."""
    return ZEBRA_FILL if i % 2 == 1 else WHITE_FILL

HELPER_SHEET_NAME = "Расчет по блокам (служебный)"
SHEET1_NAME = "Экономика проекта"
SHEET2_NAME = "СМР по методике"

EXCEL_INPUT_HEADERS = [
    "Название блока", "Тип блока",
    "S квартир, м2", "S коммерции 1 эт., м2", "S кладовых, м2",
    "Подземный паркинг, м/м", "Наземный/Многоуровневый паркинг, м/м",
    "Цена жилья, руб/м2", "Цена коммерции, руб/м2", "Цена кладовых, руб/м2",
    "Цена подземного м/м, руб", "Цена наземного м/м, руб",
    "Ставка СМР подземного м/м, руб", "Ставка СМР наземного м/м, руб",
    "Ставка СМР коммерции паркинга, руб/м2", "Ставка СМР кладовых паркинга, руб/м2",
] + PARAM_COLS
EXCEL_CALC_HEADERS = [
    "NSA, м2", "Доля аллокации", "Аллоцированные затраты",
    "Себестоимость коробки (методика)",
    "Наружные работы блока (G), руб", "Прочие затраты блока (Z), руб",
    "Инфраструктура блока (Z5010), руб",
    "Прямые затраты", "Полные затраты", "Выручка", "Валовая прибыль", "Валовая рентабельность",
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
    cell.font = TOTAL_FONT if bold else DATA_FONT
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
    ws1.title = HELPER_SHEET_NAME

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
    style_header_row(ws1, header_row1, N_COLS_EXCEL)

    # ------------------------------------------------------------------
    # Лист 2: каталог расценок + расчет себестоимости коробки по блокам
    # (строится ПЕРЕД заполнением листа 1, т.к. лист 1 на него ссылается)
    # ------------------------------------------------------------------
    ws2 = wb.create_sheet(SHEET2_NAME)
    ws2["A1"] = "Расчет СМР по укрупненным видам работ (действующая методика)"
    ws2["A1"].font = TITLE_FONT
    ws2["A2"] = f"Проект: {project_name} | Сценарий: {scenario_name}"

    if is_mp_stage:
        # -- Этап МП: коробка = NSA блока x укрупненная ставка руб/м2 NSA --
        ws2["A4"] = "Себестоимость коробки — этап МП (укрупненная ставка руб/м2 NSA)"
        ws2["A4"].font = PARAM_FONT
        calc_header_row = 5
        calc_headers2 = ["Название блока", "Тип блока", "NSA, м2", "Ставка коробки, руб/м2 NSA", "ИТОГО СМР коробки, руб"]
        for j, h in enumerate(calc_headers2, start=1):
            ws2.cell(row=calc_header_row, column=j, value=h)
        style_header_row(ws2, calc_header_row, len(calc_headers2))

        calc_first_row = calc_header_row + 1
        total_col_idx2 = len(calc_headers2)
        for i, (_, row) in enumerate(blocks.iterrows()):
            r2 = calc_first_row + i
            r1 = first_row1 + i
            block_name = row["Название блока"]
            row_fill = zebra(i)
            style_cell(ws2.cell(row=r2, column=1, value=block_name), fill=row_fill)
            style_cell(ws2.cell(row=r2, column=2, value=row["Тип блока"]), fill=row_fill)
            nsa_ref = f"'{HELPER_SHEET_NAME}'!{COL['NSA, м2']}{r1}"
            style_cell(ws2.cell(row=r2, column=3, value=f"={nsa_ref}"), number_format=MONEY_FMT, fill=row_fill)
            rate_val = float(st.session_state.block_mp_rate.get(block_name, 0.0)) if row["Тип блока"] == TYPE_RESIDENTIAL else 0.0
            style_cell(ws2.cell(row=r2, column=4, value=rate_val), number_format=MONEY_FMT, fill=row_fill)
            total_cell = ws2.cell(row=r2, column=total_col_idx2, value=f"=C{r2}*D{r2}*'{HELPER_SHEET_NAME}'!$E$4")
            style_cell(total_cell, number_format=MONEY_MM_FMT, bold=True, fill=TOTAL_FILL)
        calc_last_row = calc_first_row + N_ROWS - 1

        autosize(ws2, len(calc_headers2), width=17)
        ws2.column_dimensions["A"].width = 16
        ws2.column_dimensions["B"].width = 24

        if N_ROWS > 0:
            bar2 = BarChart()
            bar2.type = "col"
            bar2.title = "Себестоимость коробки по блокам (МП)"
            bar2.y_axis.title = "руб"
            bar2.style = 10
            data_ref2 = Reference(ws2, min_col=total_col_idx2, max_col=total_col_idx2, min_row=calc_header_row, max_row=calc_last_row)
            cats_ref2 = Reference(ws2, min_col=1, min_row=calc_first_row, max_row=calc_last_row)
            bar2.add_data(data_ref2, titles_from_data=True)
            bar2.set_categories(cats_ref2)
            bar2.series[0].graphicalProperties.solidFill = XL_GREEN
            bar2.height, bar2.width = 10, 22
            ws2.add_chart(bar2, f"{get_column_letter(total_col_idx2 + 2)}{calc_header_row}")

        group_last_row = calc_last_row  # на МП разбивки по статьям A-E нет — группы не строятся
    else:
        # Каталог расценок — ставка ИНДИВИДУАЛЬНА для каждого жилого блока: строки —
        # 32 статьи методики, колонки — Код/Статья/Группа/Единица + одна колонка
        # ставки на каждый жилой блок (в том же порядке, что на листе 1).
        ws2["A4"] = "Каталог расценок (ставка — своя колонка на каждый жилой блок)"
        ws2["A4"].font = PARAM_FONT
        rates_header_row = 5
        rate_col_for_block = {name: get_column_letter(5 + idx) for idx, name in enumerate(res_block_names)}
        rates_headers = ["Код", "Статья затрат", "Группа", "Единица измерения"] + [
            f"Ставка «{name}», руб/ед." for name in res_block_names
        ]
        for j, h in enumerate(rates_headers, start=1):
            ws2.cell(row=rates_header_row, column=j, value=h)
        style_header_row(ws2, rates_header_row, len(rates_headers))

        code_to_unit = dict(zip(DEFAULT_RATES_DF["Код"], DEFAULT_RATES_DF["Единица измерения"]))
        rates_first_row = rates_header_row + 1
        rate_row_by_code = {}
        for i, code in enumerate(item_codes_master):
            r = rates_first_row + i
            rate_row_by_code[code] = r
            row_fill = zebra(i)
            vals = [code, code_to_name[code], code_to_group[code], code_to_unit[code]]
            for j, v in enumerate(vals, start=1):
                style_cell(ws2.cell(row=r, column=j, value=v), fill=row_fill)
            for block_idx, name in enumerate(res_block_names):
                rate_val = float(block_rate_series.get(name, pd.Series(dtype=float)).get(code, 0.0))
                cell = ws2.cell(row=r, column=5 + block_idx, value=rate_val)
                style_cell(cell, number_format=MONEY_FMT, fill=row_fill)
        rates_last_row = rates_first_row + len(item_codes_master) - 1

        # -- Кол-во остановок лифтов (статья D.10.10) — отдельный ручной ввод по блоку,
        #    не колонка листа 1 (в разных секциях разное число лифтов/остановок). --
        stops_row = rates_last_row + 1
        style_cell(ws2.cell(row=stops_row, column=1, value="—"))
        style_cell(ws2.cell(row=stops_row, column=2, value="Кол-во остановок (для D.10.10)"))
        style_cell(ws2.cell(row=stops_row, column=3, value="D"))
        style_cell(ws2.cell(row=stops_row, column=4, value="шт"))
        for block_idx, name in enumerate(res_block_names):
            stops_val = float(st.session_state.block_elevator_stops.get(name, 0.0))
            style_cell(ws2.cell(row=stops_row, column=5 + block_idx, value=stops_val), number_format="0")

        # -- Расчет по блокам: строки = блоки (в том же порядке, что на листе 1) --
        calc_title_row = stops_row + 3
        ws2.cell(row=calc_title_row, column=1, value="Себестоимость коробки по блокам (только «Жилой блок»)")
        ws2.cell(row=calc_title_row, column=1).font = PARAM_FONT
        calc_header_row = calc_title_row + 1
        item_codes = item_codes_master
        calc_headers2 = ["Название блока", "Тип блока"] + item_codes + ["ИТОГО СМР коробки, руб"]
        for j, h in enumerate(calc_headers2, start=1):
            ws2.cell(row=calc_header_row, column=j, value=h)
        style_header_row(ws2, calc_header_row, len(calc_headers2))

        calc_first_row = calc_header_row + 1
        total_col_idx2 = len(calc_headers2)  # последняя колонка — ИТОГО
        for i, (_, row) in enumerate(blocks.iterrows()):
            r2 = calc_first_row + i
            r1 = first_row1 + i  # соответствующая строка на листе 1 (тот же порядок блоков)
            block_name = row["Название блока"]
            row_fill = zebra(i)
            ws2.cell(row=r2, column=1, value=block_name)
            ws2.cell(row=r2, column=2, value=row["Тип блока"])
            style_cell(ws2.cell(row=r2, column=1), fill=row_fill)
            style_cell(ws2.cell(row=r2, column=2), fill=row_fill)

            type_ref = f"'{HELPER_SHEET_NAME}'!{COL['Тип блока']}{r1}"
            rate_col_letter = rate_col_for_block.get(block_name)
            for k, code in enumerate(item_codes):
                basis_key = code_to_basis[code]
                if basis_key == "NSA":
                    qty_ref = f"'{HELPER_SHEET_NAME}'!{COL['NSA, м2']}{r1}"
                elif basis_key == "VOL_TOTAL":
                    qty_ref = f"('{HELPER_SHEET_NAME}'!{COL['Объем здания ниже 0, м3']}{r1}+'{HELPER_SHEET_NAME}'!{COL['Объем здания выше 0, м3']}{r1})"
                elif basis_key == "STORAGE_AREA":
                    qty_ref = f"'{HELPER_SHEET_NAME}'!{COL['S кладовых, м2']}{r1}"
                elif basis_key == "ELEVATOR_STOPS":
                    qty_ref = f"{rate_col_letter}${stops_row}" if rate_col_letter is not None else "0"
                else:
                    basis_col_name = BASIS_COLUMN[basis_key]
                    qty_ref = f"'{HELPER_SHEET_NAME}'!{COL[basis_col_name]}{r1}"
                if rate_col_letter is not None:
                    rate_ref = f"{rate_col_letter}${rate_row_by_code[code]}"
                    formula = f'=IF({type_ref}="{TYPE_RESIDENTIAL}",{qty_ref}*{rate_ref},0)'
                else:
                    formula = 0  # блок-паркинг — каталога ставок методики у него нет
                cell = ws2.cell(row=r2, column=3 + k, value=formula)
                style_cell(cell, number_format=MONEY_FMT, fill=row_fill)

            first_item_col = get_column_letter(3)
            last_item_col = get_column_letter(2 + len(item_codes))
            total_cell = ws2.cell(row=r2, column=total_col_idx2, value=f"=SUM({first_item_col}{r2}:{last_item_col}{r2})*'{HELPER_SHEET_NAME}'!$E$4")
            style_cell(total_cell, number_format=MONEY_MM_FMT, bold=True, fill=TOTAL_FILL)
        calc_last_row = calc_first_row + N_ROWS - 1

        autosize(ws2, len(calc_headers2), width=13)
        ws2.column_dimensions["A"].width = 16
        ws2.column_dimensions["B"].width = 24
        ws2.column_dimensions[get_column_letter(total_col_idx2)].width = 15
        for name, col_letter in rate_col_for_block.items():
            ws2.column_dimensions[col_letter].width = 16

        # -- Группы работ (для круговой диаграммы) — суммы по прямоугольным
        #    диапазонам колонок статей внутри каждой группы (колонки статей
        #    идут подряд, сгруппированы по буквенному коду) --
        group_order, group_col_ranges = [], {}
        prev_group = None
        for k, code in enumerate(item_codes_master):
            g = code_to_group[code]
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
            bar2.series[0].graphicalProperties.solidFill = XL_GREEN
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
            pie2.dataLabels.showCatName = False
            pie2.dataLabels.showSerName = False
            pie2.dataLabels.showVal = False
            pie2.dataLabels.showLegendKey = False
            pie2.series[0].data_points = [
                DataPoint(idx=i, spPr=GraphicalProperties(solidFill=XL_GREEN_RAMP5[i % len(XL_GREEN_RAMP5)]))
                for i in range(len(group_order))
            ]
            pie2.height, pie2.width = 10, 14
            ws2.add_chart(pie2, f"{get_column_letter(total_col_idx2 + 2)}{calc_header_row + 22}")

    # ------------------------------------------------------------------
    # Наружные работы (G) и прочие затраты по СМР (Z) — ИНДИВИДУАЛЬНО по
    # каждому жилому блоку (свой локальный участок под блоком). Ставка —
    # своя колонка на каждый блок (как в каталоге A-E выше); кол-во для
    # статей на площадь участка — площадь участка САМОГО БЛОКА (лист 1).
    # ------------------------------------------------------------------
    gz_block_names = list(blocks["Название блока"])
    g_area_rate_series = {
        name: pd.to_numeric(
            st.session_state.block_g_area.get(name, DEFAULT_G_AREA_DF).set_index("Код")["Ставка, руб/ед."],
            errors="coerce",
        ).fillna(0.0)
        for name in gz_block_names
    }
    g_length_qty_series = {
        name: pd.to_numeric(
            st.session_state.block_g_length.get(name, DEFAULT_G_LENGTH_DF).set_index("Код")["Кол-во"], errors="coerce"
        ).fillna(0.0)
        for name in gz_block_names
    }
    g_length_rate_series = {
        name: pd.to_numeric(
            st.session_state.block_g_length.get(name, DEFAULT_G_LENGTH_DF).set_index("Код")["Ставка, руб/ед."],
            errors="coerce",
        ).fillna(0.0)
        for name in gz_block_names
    }
    z_pct_rate_series = {
        name: pd.to_numeric(
            st.session_state.block_z_pct.get(name, DEFAULT_Z_PCT_DF).set_index("Код")["Ставка, доля от СМР+G"],
            errors="coerce",
        ).fillna(0.0)
        for name in gz_block_names
    }
    z_fixed_sum_series = {
        name: pd.to_numeric(
            st.session_state.block_z_fixed.get(name, DEFAULT_Z_FIXED_DF).set_index("Код")["Сумма, руб"],
            errors="coerce",
        ).fillna(0.0)
        for name in gz_block_names
    }

    g_title_row = group_last_row + 3
    ws2.cell(row=g_title_row, column=1, value="Наружные работы (код G) — своя колонка ставок на каждый блок (жилой и паркинг)")
    ws2.cell(row=g_title_row, column=1).font = PARAM_FONT

    g_area_header_row = g_title_row + 1
    rate_col_for_g_area = {name: get_column_letter(4 + idx) for idx, name in enumerate(gz_block_names)}
    g_area_headers = ["Код", "Статья затрат", "Единица измерения"] + [
        f"Ставка «{name}», руб/ед." for name in gz_block_names
    ]
    for j, h in enumerate(g_area_headers, start=1):
        ws2.cell(row=g_area_header_row, column=j, value=h)
    style_header_row(ws2, g_area_header_row, len(g_area_headers))

    g_area_first_row = g_area_header_row + 1
    for i, (code, name, unit, _default_rate) in enumerate(ITEMS_G_AREA):
        r = g_area_first_row + i
        style_cell(ws2.cell(row=r, column=1, value=code))
        style_cell(ws2.cell(row=r, column=2, value=name))
        style_cell(ws2.cell(row=r, column=3, value=unit))
        for idx, block_name in enumerate(gz_block_names):
            rate_val = float(g_area_rate_series[block_name].get(code, 0.0))
            cell = ws2.cell(row=r, column=4 + idx, value=rate_val)
            style_cell(cell, number_format=MONEY_FMT)
    g_area_last_row = g_area_first_row + len(ITEMS_G_AREA) - 1
    for col_letter in rate_col_for_g_area.values():
        ws2.column_dimensions[col_letter].width = 16

    g_length_title_row = g_area_last_row + 2
    ws2.cell(row=g_length_title_row, column=1, value="Сети и кабели блока (кол-во вводится вручную по каждому блоку)")
    ws2.cell(row=g_length_title_row, column=1).font = PARAM_FONT
    g_length_header_row = g_length_title_row + 1
    glen_cols_for_block = {}
    g_length_headers = ["Код", "Статья затрат", "Единица измерения"]
    for idx, name in enumerate(gz_block_names):
        qty_col = get_column_letter(4 + 2 * idx)
        rate_col = get_column_letter(5 + 2 * idx)
        glen_cols_for_block[name] = (qty_col, rate_col)
        g_length_headers += [f"Кол-во «{name}»", f"Ставка «{name}», руб/ед."]
    for j, h in enumerate(g_length_headers, start=1):
        ws2.cell(row=g_length_header_row, column=j, value=h)
    style_header_row(ws2, g_length_header_row, len(g_length_headers))

    g_length_first_row = g_length_header_row + 1
    for i, (code, name, unit, _default_qty, _default_rate) in enumerate(ITEMS_G_LENGTH):
        r = g_length_first_row + i
        style_cell(ws2.cell(row=r, column=1, value=code))
        style_cell(ws2.cell(row=r, column=2, value=name))
        style_cell(ws2.cell(row=r, column=3, value=unit))
        for block_name, (qty_col, rate_col) in glen_cols_for_block.items():
            qty_val = float(g_length_qty_series[block_name].get(code, 0.0))
            rate_val = float(g_length_rate_series[block_name].get(code, 0.0))
            style_cell(ws2.cell(row=r, column=column_index_from_string(qty_col), value=qty_val), number_format="0.0")
            style_cell(ws2.cell(row=r, column=column_index_from_string(rate_col), value=rate_val), number_format=MONEY_FMT)
    g_length_last_row = g_length_first_row + len(ITEMS_G_LENGTH) - 1
    for qty_col, rate_col in glen_cols_for_block.values():
        ws2.column_dimensions[qty_col].width = 12
        ws2.column_dimensions[rate_col].width = 16

    # ------------------------------------------------------------------
    # Прочие затраты, связанные с СМР (Z) — своя колонка ставок/сумм на
    # каждый жилой блок; % статьи считаются от базы (коробка блока + G блока).
    # ------------------------------------------------------------------
    z_pct_title_row = g_length_last_row + 2
    ws2.cell(row=z_pct_title_row, column=1, value="Прочие затраты, связанные с СМР (код Z) — % от базы (жилой: коробка + G + подземный паркинг; паркинг: СМР паркинга + G)")
    ws2.cell(row=z_pct_title_row, column=1).font = PARAM_FONT
    z_pct_header_row = z_pct_title_row + 1
    rate_col_for_z_pct = {name: get_column_letter(3 + idx) for idx, name in enumerate(gz_block_names)}
    z_pct_headers = ["Код", "Статья затрат"] + [f"Ставка «{name}», доля" for name in gz_block_names]
    for j, h in enumerate(z_pct_headers, start=1):
        ws2.cell(row=z_pct_header_row, column=j, value=h)
    style_header_row(ws2, z_pct_header_row, len(z_pct_headers))

    z_pct_first_row = z_pct_header_row + 1
    for i, (code, name, _default_rate) in enumerate(ITEMS_Z_PCT):
        r = z_pct_first_row + i
        style_cell(ws2.cell(row=r, column=1, value=code))
        style_cell(ws2.cell(row=r, column=2, value=name))
        for idx, block_name in enumerate(gz_block_names):
            rate_val = float(z_pct_rate_series[block_name].get(code, 0.0))
            style_cell(ws2.cell(row=r, column=3 + idx, value=rate_val), number_format=PERCENT_FMT)
    z_pct_last_row = z_pct_first_row + len(ITEMS_Z_PCT) - 1
    for col_letter in rate_col_for_z_pct.values():
        ws2.column_dimensions[col_letter].width = 14

    z_fixed_title_row = z_pct_last_row + 2
    ws2.cell(row=z_fixed_title_row, column=1, value="Статьи Z с прямым вводом суммы (нет формульной базы) — по каждому блоку")
    ws2.cell(row=z_fixed_title_row, column=1).font = PARAM_FONT
    z_fixed_header_row = z_fixed_title_row + 1
    sum_col_for_z_fixed = {name: get_column_letter(3 + idx) for idx, name in enumerate(gz_block_names)}
    z_fixed_headers = ["Код", "Статья затрат"] + [f"Сумма «{name}», руб" for name in gz_block_names]
    for j, h in enumerate(z_fixed_headers, start=1):
        ws2.cell(row=z_fixed_header_row, column=j, value=h)
    style_header_row(ws2, z_fixed_header_row, len(z_fixed_headers))

    z_fixed_first_row = z_fixed_header_row + 1
    for i, (code, name, _default_sum) in enumerate(ITEMS_Z_FIXED):
        r = z_fixed_first_row + i
        style_cell(ws2.cell(row=r, column=1, value=code))
        style_cell(ws2.cell(row=r, column=2, value=name))
        for idx, block_name in enumerate(gz_block_names):
            sum_val = float(z_fixed_sum_series[block_name].get(code, 0.0))
            style_cell(ws2.cell(row=r, column=3 + idx, value=sum_val), number_format=MONEY_FMT)
    z_fixed_last_row = z_fixed_first_row + len(ITEMS_Z_FIXED) - 1
    for col_letter in sum_col_for_z_fixed.values():
        ws2.column_dimensions[col_letter].width = 16

    # ------------------------------------------------------------------
    # Сводная таблица ИТОГО G и Z по блокам (строки = все блоки, тот же
    # порядок, что на листе 1) — простые формулы: площадь участка блока x
    # SUM(ставок) для G-площади, SUMPRODUCT(кол-во,ставка) для сетей.
    # ------------------------------------------------------------------
    gz_title_row = z_fixed_last_row + 3
    ws2.cell(row=gz_title_row, column=1, value="ИТОГО наружные работы (G) и прочие затраты СМР (Z) по блокам")
    ws2.cell(row=gz_title_row, column=1).font = PARAM_FONT
    gz_header_row = gz_title_row + 1
    gz_headers = [
        "Название блока", "Тип блока", "Площадь участка блока, га",
        "G: площадь участка, руб", "G: сети/кабели, руб", "ИТОГО G, руб",
        "Себестоимость коробки (A-E), руб", "Паркинг блока (подз. / СМР паркинга), руб",
        "База для % Z (короб.+G+паркинг), руб",
        "Z: % от базы, руб", "Z: прямой ввод, руб", "ИТОГО Z, руб",
    ]
    for j, h in enumerate(gz_headers, start=1):
        ws2.cell(row=gz_header_row, column=j, value=h)
    style_header_row(ws2, gz_header_row, len(gz_headers))
    GZ_COL_PLOT, GZ_COL_GAREA, GZ_COL_GLEN, GZ_COL_GTOTAL = 3, 4, 5, 6
    GZ_COL_KOROBKA, GZ_COL_PARKING, GZ_COL_ZBASE = 7, 8, 9
    GZ_COL_ZPCT, GZ_COL_ZFIXED, GZ_COL_ZTOTAL = 10, 11, 12

    gz_first_row = gz_header_row + 1
    for i, (_, row) in enumerate(blocks.iterrows()):
        r = gz_first_row + i
        r1 = first_row1 + i
        r2 = calc_first_row + i
        block_name = row["Название блока"]
        row_fill = zebra(i)
        style_cell(ws2.cell(row=r, column=1, value=block_name), fill=row_fill)
        style_cell(ws2.cell(row=r, column=2, value=row["Тип блока"]), fill=row_fill)

        plot_area_ref = f"'{HELPER_SHEET_NAME}'!{COL['Площадь участка блока, га']}{r1}"
        style_cell(ws2.cell(row=r, column=GZ_COL_PLOT, value=f"={plot_area_ref}"), number_format="0.00", fill=row_fill)

        g_area_rate_col = rate_col_for_g_area.get(block_name)
        g_area_formula = (
            f"={plot_area_ref}*SUM({g_area_rate_col}{g_area_first_row}:{g_area_rate_col}{g_area_last_row})"
            f"*'{HELPER_SHEET_NAME}'!$E$4"
            if g_area_rate_col is not None else 0
        )
        style_cell(ws2.cell(row=r, column=GZ_COL_GAREA, value=g_area_formula), number_format=MONEY_FMT, fill=row_fill)

        qty_col, rate_col = glen_cols_for_block.get(block_name, (None, None))
        g_length_formula = (
            f"=SUMPRODUCT({qty_col}{g_length_first_row}:{qty_col}{g_length_last_row},"
            f"{rate_col}{g_length_first_row}:{rate_col}{g_length_last_row})"
            f"*'{HELPER_SHEET_NAME}'!$E$4"
            if qty_col is not None else 0
        )
        style_cell(ws2.cell(row=r, column=GZ_COL_GLEN, value=g_length_formula), number_format=MONEY_FMT, fill=row_fill)

        g_total_formula = f"={get_column_letter(GZ_COL_GAREA)}{r}+{get_column_letter(GZ_COL_GLEN)}{r}"
        style_cell(ws2.cell(row=r, column=GZ_COL_GTOTAL, value=g_total_formula), number_format=MONEY_FMT, bold=True, fill=TOTAL_FILL)

        korobka_formula = f"={get_column_letter(total_col_idx2)}{r2}"
        style_cell(ws2.cell(row=r, column=GZ_COL_KOROBKA, value=korobka_formula), number_format=MONEY_MM_FMT, fill=row_fill)

        # Подземный паркинг в составе урбан-блока — генподряд/непредвиденные (Z%)
        # начисляются в т.ч. на его стоимость (та же % ставка блока).
        # Паркинг-блок: СМР паркинга = наземные м/м + коммерция + кладовые.
        _h = f"'{HELPER_SHEET_NAME}'!"
        parking_formula = (
            f'=IF({_h}{COL["Тип блока"]}{r1}="{TYPE_RESIDENTIAL}",'
            f"{_h}{COL['Подземный паркинг, м/м']}{r1}*{_h}{COL['Ставка СМР подземного м/м, руб']}{r1},"
            f"{_h}{COL['Наземный/Многоуровневый паркинг, м/м']}{r1}*{_h}{COL['Ставка СМР наземного м/м, руб']}{r1}"
            f"+{_h}{COL['S коммерции 1 эт., м2']}{r1}*{_h}{COL['Ставка СМР коммерции паркинга, руб/м2']}{r1}"
            f"+{_h}{COL['S кладовых, м2']}{r1}*{_h}{COL['Ставка СМР кладовых паркинга, руб/м2']}{r1})"
            f"*{_h}$E$4"
        )
        style_cell(ws2.cell(row=r, column=GZ_COL_PARKING, value=parking_formula), number_format=MONEY_FMT, fill=row_fill)

        zbase_formula = (
            f"={get_column_letter(GZ_COL_KOROBKA)}{r}+{get_column_letter(GZ_COL_GTOTAL)}{r}"
            f"+{get_column_letter(GZ_COL_PARKING)}{r}"
        )
        style_cell(ws2.cell(row=r, column=GZ_COL_ZBASE, value=zbase_formula), number_format=MONEY_MM_FMT, fill=row_fill)

        z_pct_rate_col = rate_col_for_z_pct.get(block_name)
        z_pct_formula = (
            f"={get_column_letter(GZ_COL_ZBASE)}{r}*SUM({z_pct_rate_col}{z_pct_first_row}:{z_pct_rate_col}{z_pct_last_row})"
            if z_pct_rate_col is not None else 0
        )
        style_cell(ws2.cell(row=r, column=GZ_COL_ZPCT, value=z_pct_formula), number_format=MONEY_FMT, fill=row_fill)

        z_fixed_col = sum_col_for_z_fixed.get(block_name)
        z_fixed_formula = (
            f"=SUM({z_fixed_col}{z_fixed_first_row}:{z_fixed_col}{z_fixed_last_row})"
            f"*'{HELPER_SHEET_NAME}'!$E$4"
            if z_fixed_col is not None else 0
        )
        style_cell(ws2.cell(row=r, column=GZ_COL_ZFIXED, value=z_fixed_formula), number_format=MONEY_FMT, fill=row_fill)

        z_total_formula = f"={get_column_letter(GZ_COL_ZPCT)}{r}+{get_column_letter(GZ_COL_ZFIXED)}{r}"
        style_cell(ws2.cell(row=r, column=GZ_COL_ZTOTAL, value=z_total_formula), number_format=MONEY_FMT, bold=True, fill=TOTAL_FILL)
    gz_last_row = gz_first_row + N_ROWS - 1

    autosize(ws2, len(gz_headers), width=15)
    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 24

    # Пул косвенных расходов на Листе 1 = статьи сайдбара (только затраты,
    # общие на весь участок). G и Z теперь считаются по блокам (Лист 2) и
    # входят напрямую в прямые затраты каждого блока, а не в этот пул.
    ws1["H4"] = indirect_pool_sidebar
    ws1["H4"].number_format = MONEY_FMT

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
        f_nsa = f'=IF({c["Тип блока"]}{r1}="{TYPE_RESIDENTIAL}",{c["S квартир, м2"]}{r1},0)+{c["S коммерции 1 эт., м2"]}{r1}+{c["S кладовых, м2"]}{r1}'
        f_share = f'=IF(SUM(${c["NSA, м2"]}${first_row1}:${c["NSA, м2"]}${last_row1})=0,0,{c["NSA, м2"]}{r1}/SUM(${c["NSA, м2"]}${first_row1}:${c["NSA, м2"]}${last_row1}))'
        f_alloc = f'={c["Доля аллокации"]}{r1}*$H$4*$E$4'
        f_korobka = f"='{SHEET2_NAME}'!{get_column_letter(total_col_idx2)}{r2}"
        r_gz = gz_first_row + i
        f_g = f"='{SHEET2_NAME}'!{get_column_letter(GZ_COL_GTOTAL)}{r_gz}"
        f_z = f"='{SHEET2_NAME}'!{get_column_letter(GZ_COL_ZTOTAL)}{r_gz}"
        # Инфраструктура блока (Z.50.10, Вкладка 4) — детализация по статьям
        # не выгружается (только итог по блоку), поэтому пишется как значение,
        # а не как формула со ссылкой на отдельный лист-каталог.
        f_infra = float(row["Инфраструктура блока (Z5010), руб"])
        f_direct = (
            f'=IF({c["Тип блока"]}{r1}="{TYPE_RESIDENTIAL}",'
            f'{c["Себестоимость коробки (методика)"]}{r1}+{c["Наружные работы блока (G), руб"]}{r1}'
            f'+{c["Прочие затраты блока (Z), руб"]}{r1}'
            f'+{c["Подземный паркинг, м/м"]}{r1}*{c["Ставка СМР подземного м/м, руб"]}{r1}*$E$4'
            f'+{c["Инфраструктура блока (Z5010), руб"]}{r1},'
            f'({c["Наземный/Многоуровневый паркинг, м/м"]}{r1}*{c["Ставка СМР наземного м/м, руб"]}{r1}'
            f'+{c["S коммерции 1 эт., м2"]}{r1}*{c["Ставка СМР коммерции паркинга, руб/м2"]}{r1}'
            f'+{c["S кладовых, м2"]}{r1}*{c["Ставка СМР кладовых паркинга, руб/м2"]}{r1})*$E$4'
            f'+{c["Наружные работы блока (G), руб"]}{r1}+{c["Прочие затраты блока (Z), руб"]}{r1}'
            f'+{c["Инфраструктура блока (Z5010), руб"]}{r1})'
        )
        f_full = f'={c["Аллоцированные затраты"]}{r1}+{c["Прямые затраты"]}{r1}'
        f_revenue = (
            f'=IF({c["Тип блока"]}{r1}="{TYPE_RESIDENTIAL}",'
            f'({c["S квартир, м2"]}{r1}*{c["Цена жилья, руб/м2"]}{r1}'
            f'+{c["S коммерции 1 эт., м2"]}{r1}*{c["Цена коммерции, руб/м2"]}{r1}'
            f'+{c["S кладовых, м2"]}{r1}*{c["Цена кладовых, руб/м2"]}{r1}'
            f'+{c["Подземный паркинг, м/м"]}{r1}*{c["Цена подземного м/м, руб"]}{r1})*$B$4,'
            f'({c["Наземный/Многоуровневый паркинг, м/м"]}{r1}*{c["Цена наземного м/м, руб"]}{r1}'
            f'+{c["S коммерции 1 эт., м2"]}{r1}*{c["Цена коммерции, руб/м2"]}{r1}'
            f'+{c["S кладовых, м2"]}{r1}*{c["Цена кладовых, руб/м2"]}{r1})*$B$4)'
        )
        f_profit = f'={c["Выручка"]}{r1}-{c["Полные затраты"]}{r1}'
        f_margin = f'=IF({c["Выручка"]}{r1}=0,0,{c["Валовая прибыль"]}{r1}/{c["Выручка"]}{r1})'

        calc_formulas = [
            f_nsa, f_share, f_alloc, f_korobka, f_g, f_z, f_infra, f_direct, f_full, f_revenue, f_profit, f_margin,
        ]
        for k, formula in enumerate(calc_formulas):
            col_idx = len(EXCEL_INPUT_HEADERS) + 1 + k
            cell = ws1.cell(row=r1, column=col_idx, value=formula)
            header_name = EXCEL_CALC_HEADERS[k]
            fmt = PERCENT_FMT if header_name in ("Доля аллокации", "Валовая рентабельность") else MONEY_FMT
            style_cell(cell, number_format=fmt)

    total_row1 = last_row1 + 1 if N_ROWS > 0 else first_row1
    ws1.cell(row=total_row1, column=1, value="ИТОГО")
    if N_ROWS > 0:
        for header_name in ["NSA, м2", "Доля аллокации", "Аллоцированные затраты", "Себестоимость коробки (методика)",
                             "Наружные работы блока (G), руб", "Прочие затраты блока (Z), руб",
                             "Инфраструктура блока (Z5010), руб",
                             "Прямые затраты", "Полные затраты", "Выручка", "Валовая прибыль"]:
            col_letter = COL[header_name]
            ws1[f"{col_letter}{total_row1}"] = f"=SUM({col_letter}{first_row1}:{col_letter}{last_row1})"
        ws1[f'{COL["Валовая рентабельность"]}{total_row1}'] = (
            f'=IF({COL["Выручка"]}{total_row1}=0,0,{COL["Валовая прибыль"]}{total_row1}/{COL["Выручка"]}{total_row1})'
        )
    else:
        for header_name in EXCEL_CALC_HEADERS:
            ws1[f"{COL[header_name]}{total_row1}"] = 0
    for j in range(1, N_COLS_EXCEL + 1):
        cell = ws1.cell(row=total_row1, column=j)
        header_name = EXCEL_HEADERS[j - 1]
        fmt = PERCENT_FMT if header_name in ("Доля аллокации", "Валовая рентабельность") else (MONEY_FMT if j >= 3 else None)
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
        bar.series[0].graphicalProperties.solidFill = XL_GRAY
        bar.series[1].graphicalProperties.solidFill = XL_GREEN
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
            ("Коммерция", f'=SUMPRODUCT({rng_c1},{rng_p_c1})*$B$4'),
            ("Кладовые", f'=SUMPRODUCT({rng_storage},{rng_p_storage})*$B$4'),
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
    pie.dataLabels.showCatName = False
    pie.dataLabels.showSerName = False
    pie.dataLabels.showVal = False
    pie.dataLabels.showLegendKey = False
    pie.series[0].data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=XL_GREEN_RAMP5[i % len(XL_GREEN_RAMP5)]))
        for i in range(len(struct_formulas))
    ]
    pie.height, pie.width = 10, 16
    ws1.add_chart(pie, f"{get_column_letter(N_COLS_EXCEL + 2)}{struct_header_row}")

    # ------------------------------------------------------------------
    # Логотип Талан в шапке листов 1 и 2
    # ------------------------------------------------------------------
    logo_bytes = base64.b64decode(TALAN_LOGO_PNG_B64)
    logo_img1 = XLImage(PILImage.open(BytesIO(logo_bytes)))
    logo_img1.width, logo_img1.height = 170, 32
    ws1.add_image(logo_img1, f"{get_column_letter(N_COLS_EXCEL + 2)}1")
    ws1.sheet_view.showGridLines = False

    logo_img2 = XLImage(PILImage.open(BytesIO(logo_bytes)))
    logo_img2.width, logo_img2.height = 170, 32
    ws2.add_image(logo_img2, f"{get_column_letter(total_col_idx2 + 2)}1")
    ws2.sheet_view.showGridLines = False

    # Служебный лист (данные по блокам) не показываем пользователю — он
    # используется только формулами "СМР по методике" и нового отчета ниже.
    ws1.sheet_state = "hidden"

    # ==================================================================
    # НОВЫЙ ЛИСТ "Экономика проекта" — свод по форме заказчика: пара колонок
    # (Сумма + на 1 кв.м. NSA) на каждый блок + колонка ИТОГО по проекту;
    # строки — иерархия статей затрат методики (группы A-Z), код разбит на
    # 3 колонки (буква/№1/№2), сверху блок продаж/выручки, снизу — прямые
    # /полные затраты, выручка, прибыль, рентабельность. Все ячейки — живые
    # формулы (ссылки на служебный лист и на «СМР по методике»).
    # ------------------------------------------------------------------
    ws_report = wb.create_sheet(SHEET1_NAME)
    ws_report.sheet_view.showGridLines = False

    block_names_all = list(blocks["Название блока"])
    block_types_all = dict(zip(blocks["Название блока"], blocks["Тип блока"]))
    n_blocks_r = len(block_names_all)
    FIRST_BLOCK_COL = 5  # A-D заняты кодом статьи (буква/№1/№2) и наименованием
    r_sum_col = {name: FIRST_BLOCK_COL + 2 * i for i, name in enumerate(block_names_all)}
    r_pm2_col = {name: FIRST_BLOCK_COL + 2 * i + 1 for i, name in enumerate(block_names_all)}
    r_total_sum_col = FIRST_BLOCK_COL + 2 * n_blocks_r
    r_total_pm2_col = r_total_sum_col + 1
    n_cols_report = r_total_pm2_col

    calc_row_for_block = {name: calc_first_row + i for i, name in enumerate(block_names_all)}
    gz_row_for_block = {name: gz_first_row + i for i, name in enumerate(block_names_all)}
    helper_row_for_block = {name: first_row1 + i for i, name in enumerate(block_names_all)}
    item_col_in_calc = (
        {} if is_mp_stage else {code: get_column_letter(3 + k) for k, code in enumerate(item_codes_master)}
    )
    g_area_row_for_code = {code: g_area_first_row + i for i, (code, *_r) in enumerate(ITEMS_G_AREA)}
    g_length_row_for_code = {code: g_length_first_row + i for i, (code, *_r) in enumerate(ITEMS_G_LENGTH)}
    z_pct_row_for_code = {code: z_pct_first_row + i for i, (code, *_r) in enumerate(ITEMS_Z_PCT)}
    z_fixed_row_for_code = {code: z_fixed_first_row + i for i, (code, *_r) in enumerate(ITEMS_Z_FIXED)}

    def r_nsa_ref(name):
        return f"'{HELPER_SHEET_NAME}'!{COL['NSA, м2']}{helper_row_for_block[name]}"

    r_nsa_total_ref = f"'{HELPER_SHEET_NAME}'!{COL['NSA, м2']}{total_row1}" if N_ROWS > 0 else "0"

    def r_pm2_formula(sum_ref, nsa_ref):
        return f"=IF({nsa_ref}=0,0,{sum_ref}/{nsa_ref})"

    def r_helper_ref(header_name, name):
        return f"'{HELPER_SHEET_NAME}'!{COL[header_name]}{helper_row_for_block[name]}"

    def r_leaf_formula(source, name):
        """Формула суммы по одной статье затрат методики для блока name.
        Каталог A-Z считается только по жилым блокам (паркинг — вне каталога
        СМР, у него своя ставка паркинга) и только там, где статья есть в
        модели — иначе 0 (по решению: каталог модели не меняем). Исключение —
        Z.50.10 (Инфраструктура, Вкладка 4): считается по ЛЮБОМУ типу блока."""
        if source is None:
            return 0
        kind, code = source
        if kind == "INFRA_Z5010":
            return f"='{HELPER_SHEET_NAME}'!{COL['Инфраструктура блока (Z5010), руб']}{helper_row_for_block[name]}"
        if kind == "ITEMS" and block_types_all.get(name) != TYPE_RESIDENTIAL:
            return 0  # каталог A-E — только жилые блоки; G и Z — все блоки
        if kind == "ITEMS":
            col = item_col_in_calc.get(code)
            if col is None:
                return 0
            return f"='{SHEET2_NAME}'!{col}{calc_row_for_block[name]}"
        if kind == "G_AREA":
            rate_col = rate_col_for_g_area.get(name)
            if rate_col is None:
                return 0
            row = g_area_row_for_code[code]
            plot_ref = f"'{HELPER_SHEET_NAME}'!{COL['Площадь участка блока, га']}{helper_row_for_block[name]}"
            return f"={plot_ref}*'{SHEET2_NAME}'!{rate_col}{row}*'{HELPER_SHEET_NAME}'!$E$4"
        if kind == "G_LENGTH":
            cols = glen_cols_for_block.get(name)
            if cols is None:
                return 0
            qty_col, rate_col = cols
            row = g_length_row_for_code[code]
            return f"='{SHEET2_NAME}'!{qty_col}{row}*'{SHEET2_NAME}'!{rate_col}{row}*'{HELPER_SHEET_NAME}'!$E$4"
        if kind == "Z_PCT":
            rate_col = rate_col_for_z_pct.get(name)
            if rate_col is None:
                return 0
            row = z_pct_row_for_code[code]
            zbase_ref = f"'{SHEET2_NAME}'!{get_column_letter(GZ_COL_ZBASE)}{gz_row_for_block[name]}"
            return f"={zbase_ref}*'{SHEET2_NAME}'!{rate_col}{row}"
        if kind == "Z_FIXED":
            sum_col = sum_col_for_z_fixed.get(name)
            if sum_col is None:
                return 0
            row = z_fixed_row_for_code[code]
            return f"='{SHEET2_NAME}'!{sum_col}{row}*'{HELPER_SHEET_NAME}'!$E$4"
        return 0

    def rr_write_labels(row, letter, num1, num2, label, bold=False, section=False):
        if section:
            ws_report.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            c = ws_report.cell(row=row, column=1, value=label)
            c.font = TITLE_FONT if bold else PARAM_FONT
            return
        if letter is not None:
            style_cell(ws_report.cell(row=row, column=1, value=letter))
        if num1 is not None:
            style_cell(ws_report.cell(row=row, column=2, value=num1))
        if num2 is not None:
            style_cell(ws_report.cell(row=row, column=3, value=num2))
        cell_d = ws_report.cell(row=row, column=4, value=label)
        style_cell(cell_d)
        if bold:
            cell_d.font = PARAM_FONT

    def rr_write_data_row(row, block_formula_fn, fill_pm2=True, bold=False, fmt=MONEY_FMT,
                           total_mode="sum", total_ref=None):
        """total_mode: 'sum' — ИТОГО = сумма по блокам (по умолчанию);
        'ref' — ИТОГО = total_ref (формула-строка), напр. для % показателей;
        'none' — колонка ИТОГО не заполняется (напр. для средних цен)."""
        fill = TOTAL_FILL if bold else None
        for name in block_names_all:
            sc, pc = r_sum_col[name], r_pm2_col[name]
            val = block_formula_fn(name)
            style_cell(ws_report.cell(row=row, column=sc, value=val), number_format=fmt, bold=bold, fill=fill)
            if fill_pm2:
                nsa_ref = r_nsa_ref(name)
                sum_ref = f"{get_column_letter(sc)}{row}"
                pm2_val = r_pm2_formula(sum_ref, nsa_ref)
                style_cell(ws_report.cell(row=row, column=pc, value=pm2_val), number_format=MONEY_FMT, bold=bold, fill=fill)
        if total_mode == "none":
            return
        if total_mode == "ref" and total_ref is not None:
            total_formula = total_ref
        elif N_ROWS > 0:
            total_formula = "=SUM(" + ",".join(f"{get_column_letter(r_sum_col[n])}{row}" for n in block_names_all) + ")"
        else:
            total_formula = 0
        style_cell(ws_report.cell(row=row, column=r_total_sum_col, value=total_formula),
                   number_format=fmt, bold=True, fill=TOTAL_FILL)
        if fill_pm2:
            if total_mode == "ref" and total_ref is not None:
                tpm2_val = total_ref  # % и т.п. — то же значение, доля от м2 не нужна
            else:
                total_sum_ref = f"{get_column_letter(r_total_sum_col)}{row}"
                tpm2_val = r_pm2_formula(total_sum_ref, r_nsa_total_ref)
            style_cell(ws_report.cell(row=row, column=r_total_pm2_col, value=tpm2_val),
                       number_format=MONEY_FMT, bold=True, fill=TOTAL_FILL)
        else:
            style_cell(ws_report.cell(row=row, column=r_total_pm2_col), fill=TOTAL_FILL)

    # -- шапка отчета (проект/город/сценарий/параметры) --
    ws_report["A1"] = f"Проект: {project_name}"
    ws_report["A1"].font = TITLE_FONT
    ws_report["A2"] = f"Город: {project_city}"
    ws_report["A2"].font = TITLE_FONT
    ws_report["A3"] = f"Сценарий: {scenario_name}"
    ws_report["A3"].font = TITLE_FONT
    ws_report["A4"] = "Коэфф. выручки (сценарий):"
    ws_report["A4"].font = PARAM_FONT
    ws_report["B4"] = rev_factor
    ws_report["D4"] = "Коэфф. затрат (сценарий):"
    ws_report["D4"].font = PARAM_FONT
    ws_report["E4"] = cost_factor

    # -- шапка колонок: наименование блока на пару (Сумма | на 1 кв.м.) --
    ROW_BLOCK_HDR, ROW_SUBHDR = 6, 7
    ws_report.cell(row=ROW_BLOCK_HDR, column=1, value="Код затрат").font = HEADER_FONT
    ws_report.merge_cells(start_row=ROW_SUBHDR, start_column=1, end_row=ROW_SUBHDR, end_column=3)
    ws_report.cell(row=ROW_SUBHDR, column=1, value="код затрат")
    ws_report.cell(row=ROW_SUBHDR, column=4, value="Статьи затрат / показатели")
    for name in block_names_all:
        sc, pc = r_sum_col[name], r_pm2_col[name]
        ws_report.merge_cells(start_row=ROW_BLOCK_HDR, start_column=sc, end_row=ROW_BLOCK_HDR, end_column=pc)
        hcell = ws_report.cell(row=ROW_BLOCK_HDR, column=sc, value=f"{name} ({block_types_all[name]})")
        ws_report.cell(row=ROW_SUBHDR, column=sc, value="Сумма, руб")
        ws_report.cell(row=ROW_SUBHDR, column=pc, value="на 1 кв.м. NSA, руб/м2")
        ws_report.column_dimensions[get_column_letter(sc)].width = 16
        ws_report.column_dimensions[get_column_letter(pc)].width = 16
    ws_report.merge_cells(start_row=ROW_BLOCK_HDR, start_column=r_total_sum_col, end_row=ROW_BLOCK_HDR, end_column=r_total_pm2_col)
    ws_report.cell(row=ROW_BLOCK_HDR, column=r_total_sum_col, value="ИТОГО ПО ПРОЕКТУ")
    ws_report.cell(row=ROW_SUBHDR, column=r_total_sum_col, value="Сумма, руб")
    ws_report.cell(row=ROW_SUBHDR, column=r_total_pm2_col, value="на 1 кв.м. NSA, руб/м2")
    ws_report.column_dimensions[get_column_letter(r_total_sum_col)].width = 16
    ws_report.column_dimensions[get_column_letter(r_total_pm2_col)].width = 16
    style_header_row(ws_report, ROW_BLOCK_HDR, n_cols_report)
    style_header_row(ws_report, ROW_SUBHDR, n_cols_report)
    ws_report.column_dimensions["A"].width = 6
    ws_report.column_dimensions["B"].width = 6
    ws_report.column_dimensions["C"].width = 6
    ws_report.column_dimensions["D"].width = 46

    type_ref_cache = {name: r_helper_ref("Тип блока", name) for name in block_names_all}
    apt_ref = {name: r_helper_ref("S квартир, м2", name) for name in block_names_all}
    com_ref = {name: r_helper_ref("S коммерции 1 эт., м2", name) for name in block_names_all}
    storage_ref = {name: r_helper_ref("S кладовых, м2", name) for name in block_names_all}
    park_u_ref = {name: r_helper_ref("Подземный паркинг, м/м", name) for name in block_names_all}
    park_g_ref = {name: r_helper_ref("Наземный/Многоуровневый паркинг, м/м", name) for name in block_names_all}
    price_apt_ref = {name: r_helper_ref("Цена жилья, руб/м2", name) for name in block_names_all}
    price_com_ref = {name: r_helper_ref("Цена коммерции, руб/м2", name) for name in block_names_all}
    price_storage_ref = {name: r_helper_ref("Цена кладовых, руб/м2", name) for name in block_names_all}
    price_park_u_ref = {name: r_helper_ref("Цена подземного м/м, руб", name) for name in block_names_all}
    price_park_g_ref = {name: r_helper_ref("Цена наземного м/м, руб", name) for name in block_names_all}

    # -- Продаваемая площадь --
    ROW_SEC_AREA = 9
    ROW_AREA_RES, ROW_AREA_COM, ROW_AREA_STOR, ROW_AREA_PARK = 10, 11, 12, 13
    rr_write_labels(ROW_SEC_AREA, None, None, None, "ПРОДАВАЕМАЯ ПЛОЩАДЬ", bold=True, section=True)
    rr_write_labels(ROW_AREA_RES, None, None, None, "Площадь жилых, м2")
    rr_write_labels(ROW_AREA_COM, None, None, None, "Площадь нежилых, м2")
    rr_write_labels(ROW_AREA_STOR, None, None, None, "Площадь кладовых, м2")
    rr_write_labels(ROW_AREA_PARK, None, None, None, "Паркинг, м/м")
    rr_write_data_row(ROW_AREA_RES, lambda n: f"={apt_ref[n]}", fill_pm2=False, fmt=MONEY_FMT)
    rr_write_data_row(ROW_AREA_COM, lambda n: f"={com_ref[n]}", fill_pm2=False, fmt=MONEY_FMT)
    rr_write_data_row(ROW_AREA_STOR, lambda n: f"={storage_ref[n]}", fill_pm2=False, fmt=MONEY_FMT)
    rr_write_data_row(ROW_AREA_PARK, lambda n: f"={park_u_ref[n]}+{park_g_ref[n]}", fill_pm2=False, fmt=MONEY_FMT)

    # -- Средневзвешенная цена --
    ROW_SEC_PRICE = 15
    ROW_PRICE_RES, ROW_PRICE_COM, ROW_PRICE_STOR, ROW_PRICE_PARK = 16, 17, 18, 19
    rr_write_labels(ROW_SEC_PRICE, None, None, None, "СРЕДНЕВЗВЕШЕННАЯ ЦЕНА", bold=True, section=True)
    rr_write_labels(ROW_PRICE_RES, None, None, None, "Средневзвешенная жилых, руб/м2")
    rr_write_labels(ROW_PRICE_COM, None, None, None, "Средневзвешенная нежилых, руб/м2")
    rr_write_labels(ROW_PRICE_STOR, None, None, None, "Средневзвешенная кладовых, руб/м2")
    rr_write_labels(ROW_PRICE_PARK, None, None, None, "Паркинг, руб/м/м")
    rr_write_data_row(ROW_PRICE_RES, lambda n: f"={price_apt_ref[n]}", fill_pm2=False, fmt=MONEY_FMT, total_mode="none")
    rr_write_data_row(ROW_PRICE_COM, lambda n: f"={price_com_ref[n]}", fill_pm2=False, fmt=MONEY_FMT, total_mode="none")
    rr_write_data_row(ROW_PRICE_STOR, lambda n: f"={price_storage_ref[n]}", fill_pm2=False, fmt=MONEY_FMT, total_mode="none")
    rr_write_data_row(
        ROW_PRICE_PARK,
        lambda n: f'=IF({type_ref_cache[n]}="{TYPE_RESIDENTIAL}",{price_park_u_ref[n]},{price_park_g_ref[n]})',
        fill_pm2=False, fmt=MONEY_FMT, total_mode="none",
    )

    # -- Выручка (доходы) --
    ROW_REV_TOTAL, ROW_REV_RES, ROW_REV_COM, ROW_REV_STOR, ROW_REV_PARK = 21, 22, 23, 24, 25
    rr_write_labels(ROW_REV_TOTAL, None, None, None, "ИТОГО ДОХОДЫ", bold=True)
    rr_write_labels(ROW_REV_RES, None, None, None, "Доходы жилые, руб")
    rr_write_labels(ROW_REV_COM, None, None, None, "Доходы нежилые, руб")
    rr_write_labels(ROW_REV_STOR, None, None, None, "Доходы кладовые, руб")
    rr_write_labels(ROW_REV_PARK, None, None, None, "Доходы паркинг, руб")
    rev_b4 = f"'{HELPER_SHEET_NAME}'!$B$4"
    rr_write_data_row(
        ROW_REV_RES,
        lambda n: f'=IF({type_ref_cache[n]}="{TYPE_RESIDENTIAL}",{apt_ref[n]}*{price_apt_ref[n]}*{rev_b4},0)',
    )
    rr_write_data_row(
        ROW_REV_COM,
        lambda n: f"={com_ref[n]}*{price_com_ref[n]}*{rev_b4}",
    )
    rr_write_data_row(
        ROW_REV_STOR,
        lambda n: f"={storage_ref[n]}*{price_storage_ref[n]}*{rev_b4}",
    )
    rr_write_data_row(
        ROW_REV_PARK,
        lambda n: (
            f'=IF({type_ref_cache[n]}="{TYPE_RESIDENTIAL}",{park_u_ref[n]}*{price_park_u_ref[n]}*{rev_b4},'
            f'{park_g_ref[n]}*{price_park_g_ref[n]}*{rev_b4})'
        ),
    )
    rr_write_data_row(
        ROW_REV_TOTAL,
        lambda n: (
            f"={get_column_letter(r_sum_col[n])}{ROW_REV_RES}+{get_column_letter(r_sum_col[n])}{ROW_REV_COM}"
            f"+{get_column_letter(r_sum_col[n])}{ROW_REV_STOR}+{get_column_letter(r_sum_col[n])}{ROW_REV_PARK}"
        ),
        bold=True,
    )

    # -- Расходы: заголовки секции + сводные строки СМР/коробка --
    ROW_SEC_EXPENSES, ROW_SEC_SMR = 27, 28
    ROW_SMR_TOTAL, ROW_KOROBKA_TOTAL = 29, 30
    rr_write_labels(ROW_SEC_EXPENSES, None, None, None, "РАСХОДЫ", bold=True, section=True)
    rr_write_labels(ROW_SEC_SMR, None, None, None, "РАСХОДЫ по СМР", bold=True, section=True)
    rr_write_labels(ROW_SMR_TOTAL, None, None, None, "СЕБЕСТОИМОСТЬ СМР, В Т.Ч.", bold=True)
    rr_write_labels(ROW_KOROBKA_TOTAL, None, None, None, "СЕБЕСТОИМОСТЬ КОРОБКИ, В Т.Ч.", bold=True)
    rr_write_data_row(
        ROW_KOROBKA_TOTAL,
        lambda n: f"={r_helper_ref('Себестоимость коробки (методика)', n)}",
        bold=True,
    )
    rr_write_data_row(
        ROW_SMR_TOTAL,
        lambda n: (
            f"={get_column_letter(r_sum_col[n])}{ROW_KOROBKA_TOTAL}+{r_helper_ref('Наружные работы блока (G), руб', n)}"
        ),
        bold=True,
    )

    # -- Каталог статей затрат (группы A-Z, иерархия из REPORT_ROWS) --
    CATALOG_START_ROW = 31
    catalog_children, catalog_source, catalog_rows_meta = {}, {}, []
    row_cursor = CATALOG_START_ROW
    cur_group_row, cur_subgroup_row = None, None
    for level, letter, num1, num2, label, source in REPORT_ROWS:
        catalog_rows_meta.append((row_cursor, level, letter, num1, num2, label))
        catalog_source[row_cursor] = source
        if level == "group":
            cur_group_row = row_cursor
            catalog_children[row_cursor] = []
        elif level == "subgroup":
            cur_subgroup_row = row_cursor
            catalog_children[row_cursor] = []
            catalog_children[cur_group_row].append(row_cursor)
        else:  # leaf
            parent = cur_group_row if num1 is None else cur_subgroup_row
            catalog_children.setdefault(parent, []).append(row_cursor)
        row_cursor += 1
    catalog_end_row = row_cursor - 1

    def catalog_formula(row, name):
        source = catalog_source[row]
        if source is not None:
            return r_leaf_formula(source, name)
        children = catalog_children.get(row) or []
        if not children:
            return 0
        return "=SUM(" + ",".join(f"{get_column_letter(r_sum_col[name])}{cr}" for cr in children) + ")"

    for row, level, letter, num1, num2, label in catalog_rows_meta:
        rr_write_labels(row, letter, num1, num2, label, bold=(level in ("group", "subgroup")))
        rr_write_data_row(row, (lambda name, row=row: catalog_formula(row, name)), bold=(level == "group"))

    # -- Прямые/полные затраты, выручка, прибыль, рентабельность (по блокам и по проекту) --
    ROW_SPACER_TAIL = catalog_end_row + 2
    ROW_DIRECT = catalog_end_row + 3
    ROW_ALLOC = catalog_end_row + 4
    ROW_FULL = catalog_end_row + 5
    ROW_REVENUE2 = catalog_end_row + 7
    ROW_PROFIT = catalog_end_row + 8
    ROW_MARGIN = catalog_end_row + 9
    rr_write_labels(ROW_DIRECT, None, None, None, "ПРЯМЫЕ ЗАТРАТЫ", bold=True)
    rr_write_labels(ROW_ALLOC, None, None, None, "АЛЛОЦИРОВАННЫЕ ЗАТРАТЫ (косвенные расходы)", bold=True)
    rr_write_labels(ROW_FULL, None, None, None, "ПОЛНЫЕ ЗАТРАТЫ", bold=True)
    rr_write_labels(ROW_REVENUE2, None, None, None, "ВЫРУЧКА", bold=True)
    rr_write_labels(ROW_PROFIT, None, None, None, "ВАЛОВАЯ ПРИБЫЛЬ", bold=True)
    rr_write_labels(ROW_MARGIN, None, None, None, "ВАЛОВАЯ РЕНТАБЕЛЬНОСТЬ", bold=True)
    rr_write_data_row(ROW_DIRECT, lambda n: f"={r_helper_ref('Прямые затраты', n)}", bold=True)
    rr_write_data_row(ROW_ALLOC, lambda n: f"={r_helper_ref('Аллоцированные затраты', n)}", bold=True)
    rr_write_data_row(ROW_FULL, lambda n: f"={r_helper_ref('Полные затраты', n)}", bold=True)
    rr_write_data_row(ROW_REVENUE2, lambda n: f"={r_helper_ref('Выручка', n)}", bold=True)
    rr_write_data_row(ROW_PROFIT, lambda n: f"={r_helper_ref('Валовая прибыль', n)}", bold=True)
    rr_write_data_row(
        ROW_MARGIN,
        lambda n: f"={r_helper_ref('Валовая рентабельность', n)}",
        fill_pm2=False, bold=True, fmt=PERCENT_FMT,
        total_mode="ref",
        total_ref=(f"='{HELPER_SHEET_NAME}'!{COL['Валовая рентабельность']}{total_row1}" if N_ROWS > 0 else 0),
    )

    report_last_row = ROW_MARGIN
    logo_report = XLImage(PILImage.open(BytesIO(logo_bytes)))
    logo_report.width, logo_report.height = 170, 32
    ws_report.add_image(logo_report, f"{get_column_letter(n_cols_report + 2)}1")
    ws_report.freeze_panes = ws_report.cell(row=ROW_SUBHDR + 1, column=FIRST_BLOCK_COL)

    # ------------------------------------------------------------------
    # Лист "Дашборд" — ключевые метрики проекта одним экраном (для ГД)
    # ------------------------------------------------------------------
    ws_dash = wb.create_sheet("Дашборд", 0)
    ws_dash.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 23, "C": 23, "D": 23, "E": 23, "F": 23, "G": 23, "H": 3}.items():
        ws_dash.column_dimensions[col].width = w

    logo_dash = XLImage(PILImage.open(BytesIO(logo_bytes)))
    logo_dash.width, logo_dash.height = 220, 41
    ws_dash.add_image(logo_dash, "B2")
    ws_dash.row_dimensions[1].height = 8
    for r in range(2, 6):
        ws_dash.row_dimensions[r].height = 20

    ws_dash["B7"] = f'="Финансовая модель — "&\'{HELPER_SHEET_NAME}\'!A1'
    ws_dash["B7"].font = Font(name=XL_FONT, size=16, bold=True, color=XL_TEXT)
    ws_dash.merge_cells("B7:G7")
    ws_dash["B8"] = (
        f'="Город: "&SUBSTITUTE(\'{HELPER_SHEET_NAME}\'!A2,"Город: ","")&"   |   "&\'{HELPER_SHEET_NAME}\'!A3'
    )
    ws_dash["B8"].font = Font(name=XL_FONT, size=11, color=XL_TEXT2)
    ws_dash.merge_cells("B8:G8")
    ws_dash.row_dimensions[7].height = 24

    FILL_TILE = PatternFill("solid", fgColor=XL_GREEN4)
    THICK_TOP = Side(style="medium", color=XL_GREEN)

    kpi_row = 10
    tile_h = 4
    tiles = [
        ("Выручка", f"='{HELPER_SHEET_NAME}'!{COL['Выручка']}{total_row1}", MONEY_MM_FMT),
        ("Полные затраты", f"='{HELPER_SHEET_NAME}'!{COL['Полные затраты']}{total_row1}", MONEY_MM_FMT),
        ("Валовая прибыль", f"='{HELPER_SHEET_NAME}'!{COL['Валовая прибыль']}{total_row1}", MONEY_MM_FMT),
        ("Валовая рентабельность", f"='{HELPER_SHEET_NAME}'!{COL['Валовая рентабельность']}{total_row1}", PERCENT_FMT),
        ("NSA проекта, м2", f"='{HELPER_SHEET_NAME}'!{COL['NSA, м2']}{total_row1}", '#,##0 "м²"'),
        ("Пул косвенных, база", "='" + HELPER_SHEET_NAME + "'!H4", MONEY_MM_FMT),
    ]
    dash_cols = ["B", "C", "D", "E", "F", "G"]
    for col, (label, formula, fmt) in zip(dash_cols, tiles):
        top = kpi_row
        bottom = kpi_row + tile_h - 1
        for r in range(top, bottom + 1):
            ws_dash[f"{col}{r}"].fill = FILL_TILE
        ws_dash[f"{col}{top}"].border = Border(top=THICK_TOP)

        ws_dash.merge_cells(f"{col}{top}:{col}{top + 1}")
        lbl_cell = ws_dash[f"{col}{top}"]
        lbl_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        lbl_cell.font = Font(name=XL_FONT, size=10, bold=True, color=XL_TEXT2)
        lbl_cell.value = label

        val_row = top + 2
        ws_dash.merge_cells(f"{col}{val_row}:{col}{bottom}")
        val_cell = ws_dash[f"{col}{val_row}"]
        val_cell.value = formula
        val_cell.number_format = fmt
        val_cell.font = Font(name=XL_FONT, size=15, bold=True, color=XL_GREEN if fmt != PERCENT_FMT else XL_BORDO)
        val_cell.alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
    for r in range(kpi_row, kpi_row + tile_h):
        ws_dash.row_dimensions[r].height = 16

    chart_top = kpi_row + tile_h + 2

    # --- диаграмма 1: выручка / затраты / прибыль по блокам ---
    ws_dash[f"B{chart_top}"] = "Выручка, затраты и прибыль по блокам"
    ws_dash[f"B{chart_top}"].font = Font(name=XL_FONT, size=12, bold=True, color=XL_GREEN)

    if N_ROWS > 0:
        dash_bar = BarChart()
        dash_bar.type = "col"
        dash_bar.grouping = "clustered"
        dash_bar.style = 2
        dash_bar.y_axis.numFmt = MONEY_MM_FMT
        dash_bar.height, dash_bar.width = 8.5, 17
        dash_cats = Reference(ws1, min_col=1, min_row=first_row1, max_row=last_row1)
        dash_data_cost = Reference(ws1, min_col=EXCEL_HEADERS.index("Полные затраты") + 1,
                                    min_row=header_row1, max_row=last_row1)
        dash_data_rev = Reference(ws1, min_col=EXCEL_HEADERS.index("Выручка") + 1,
                                   min_row=header_row1, max_row=last_row1)
        dash_data_profit = Reference(ws1, min_col=EXCEL_HEADERS.index("Валовая прибыль") + 1,
                                      min_row=header_row1, max_row=last_row1)
        dash_bar.add_data(dash_data_cost, titles_from_data=True)
        dash_bar.add_data(dash_data_rev, titles_from_data=True)
        dash_bar.add_data(dash_data_profit, titles_from_data=True)
        dash_bar.set_categories(dash_cats)
        dash_bar.series[0].graphicalProperties.solidFill = XL_GRAY
        dash_bar.series[1].graphicalProperties.solidFill = XL_GREEN
        dash_bar.series[2].graphicalProperties.solidFill = XL_BORDO
        dash_bar.legend.position = "b"
        ws_dash.add_chart(dash_bar, f"B{chart_top + 1}")

    # --- диаграмма 2: структура полных затрат проекта ---
    struct_row = chart_top + 1
    ws_dash[f"F{chart_top}"] = "Структура полных затрат"
    ws_dash[f"F{chart_top}"].font = Font(name=XL_FONT, size=12, bold=True, color=XL_GREEN)
    cost_labels = ["Косвенные (пул)", "Коробка (методика)", "Наружные работы (G)", "Прочие СМР (Z)", "Паркинги"]
    ws_dash[f"F{struct_row}"] = "Статья"
    ws_dash[f"G{struct_row}"] = "Сумма, руб"
    for i, lbl in enumerate(cost_labels):
        ws_dash[f"F{struct_row + 1 + i}"] = lbl
    for r in range(struct_row, struct_row + 1 + len(cost_labels)):
        ws_dash[f"F{r}"].font = Font(name=XL_FONT, size=10, color=XL_TEXT)
        ws_dash[f"G{r}"].font = Font(name=XL_FONT, size=10, color=XL_TEXT)

    if N_ROWS > 0:
        rng_park_u = f"'{HELPER_SHEET_NAME}'!${COL['Подземный паркинг, м/м']}${first_row1}:${COL['Подземный паркинг, м/м']}${last_row1}"
        rng_park_u_rate = f"'{HELPER_SHEET_NAME}'!${COL['Ставка СМР подземного м/м, руб']}${first_row1}:${COL['Ставка СМР подземного м/м, руб']}${last_row1}"
        rng_park_g = f"'{HELPER_SHEET_NAME}'!${COL['Наземный/Многоуровневый паркинг, м/м']}${first_row1}:${COL['Наземный/Многоуровневый паркинг, м/м']}${last_row1}"
        rng_park_g_rate = f"'{HELPER_SHEET_NAME}'!${COL['Ставка СМР наземного м/м, руб']}${first_row1}:${COL['Ставка СМР наземного м/м, руб']}${last_row1}"
        _hs = f"'{HELPER_SHEET_NAME}'!"
        _rng = lambda col: f"{_hs}${COL[col]}${first_row1}:${COL[col]}${last_row1}"
        _rng_type = _rng("Тип блока")
        parking_total_formula = (
            f'=(SUMPRODUCT(({_rng_type}="{TYPE_RESIDENTIAL}")*{rng_park_u}*{rng_park_u_rate})'
            f"+SUMPRODUCT({rng_park_g},{rng_park_g_rate})"
            f'+SUMPRODUCT(({_rng_type}="{TYPE_PARKING}")*{_rng("S коммерции 1 эт., м2")}*{_rng("Ставка СМР коммерции паркинга, руб/м2")})'
            f'+SUMPRODUCT(({_rng_type}="{TYPE_PARKING}")*{_rng("S кладовых, м2")}*{_rng("Ставка СМР кладовых паркинга, руб/м2")}))'
            f"*{_hs}$E$4"
        )
    else:
        parking_total_formula = 0
    ws_dash[f"G{struct_row + 1}"] = f"='{HELPER_SHEET_NAME}'!{COL['Аллоцированные затраты']}{total_row1}"
    ws_dash[f"G{struct_row + 2}"] = f"='{HELPER_SHEET_NAME}'!{COL['Себестоимость коробки (методика)']}{total_row1}"
    ws_dash[f"G{struct_row + 3}"] = f"='{HELPER_SHEET_NAME}'!{COL['Наружные работы блока (G), руб']}{total_row1}"
    ws_dash[f"G{struct_row + 4}"] = f"='{HELPER_SHEET_NAME}'!{COL['Прочие затраты блока (Z), руб']}{total_row1}"
    ws_dash[f"G{struct_row + 5}"] = parking_total_formula
    for r in range(struct_row + 1, struct_row + 1 + len(cost_labels)):
        ws_dash[f"G{r}"].number_format = MONEY_MM_FMT

    dash_pie = PieChart()
    dash_pie.height, dash_pie.width = 8.5, 10.5
    dash_data_pie = Reference(ws_dash, min_col=7, min_row=struct_row, max_row=struct_row + len(cost_labels))
    dash_cats_pie = Reference(ws_dash, min_col=6, min_row=struct_row + 1, max_row=struct_row + len(cost_labels))
    dash_pie.add_data(dash_data_pie, titles_from_data=True)
    dash_pie.set_categories(dash_cats_pie)
    dash_pie.dataLabels = DataLabelList()
    dash_pie.dataLabels.showPercent = True
    dash_pie.dataLabels.showCatName = False
    dash_pie.dataLabels.showSerName = False
    dash_pie.dataLabels.showVal = False
    dash_pie.dataLabels.showLegendKey = False
    dash_pie.legend.position = "b"
    dash_pie.series[0].data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=XL_GREEN_RAMP5[i % len(XL_GREEN_RAMP5)]))
        for i in range(len(cost_labels))
    ]
    ws_dash.add_chart(dash_pie, f"F{struct_row + len(cost_labels) + 1}")

    note_row = struct_row + len(cost_labels) + 22
    ws_dash[f"B{note_row}"] = (
        "Источник: листы «Экономика проекта» и «СМР по методике». Диаграммы — с учетом коэфф. выручки/затрат сценария."
    )
    ws_dash[f"B{note_row}"].font = Font(name=XL_FONT, size=8, italic=True, color=XL_GRAY)

    ws_dash.page_setup.orientation = "landscape"
    ws_dash.page_setup.fitToWidth = 1
    ws_dash.page_setup.fitToHeight = 1
    ws_dash.sheet_properties.pageSetUpPr.fitToPage = True
    ws_dash.print_area = f"A1:H{note_row + 2}"
    ws_dash.page_margins.left = 0.3
    ws_dash.page_margins.right = 0.3
    ws_dash.page_margins.top = 0.3
    ws_dash.page_margins.bottom = 0.3

    # ------------------------------------------------------------------
    # Порядок листов: Дашборд, Экономика проекта, СМР по методике, затем
    # скрытый служебный лист с данными по блокам.
    # ------------------------------------------------------------------
    desired_order = ["Дашборд", SHEET1_NAME, SHEET2_NAME, HELPER_SHEET_NAME]
    wb._sheets = [wb[name] for name in desired_order if name in wb.sheetnames]
    wb.active = 0

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


excel_bytes = build_excel_report()
excel_slot.download_button(
    label="Скачать Excel",
    data=excel_bytes,
    file_name=f"financial_model_{scenario_name}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    icon=":material/download:",
    width="stretch",
    help="Дашборд, Экономика проекта, СМР по методике — живые формулы и графики",
)

# ======================================================================
# АВТОСОХРАНЕНИЕ ТЕКУЩЕГО СОСТОЯНИЯ (в конце каждого rerun)
# ======================================================================
_save_user, _save_project = current_user_and_project()
_save_path = current_save_path()
if _save_path is not None:
    autosave_write(_save_path, _save_user, _save_project)
