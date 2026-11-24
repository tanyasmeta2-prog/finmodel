"""Version 2 forms, Excel inputs and stage adapter for the existing financial model."""
from copy import deepcopy
from io import BytesIO
import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from methodology_v2 import *

REVISION = 2
SAVE_TABLES = ['block_mp_rates', 'mp_sections', 'ep_details', 'legacy_rates']
SAVE_PLAIN = ['methodology_revision', 'thermal_included', 'legacy_warnings']
DETAIL_CODES = {v[0] for v in DETAIL_SPEC.values()}
YELLOW = PatternFill('solid', fgColor='FFF2CC')
BLUE = PatternFill('solid', fgColor='DDEBF7')
GRAY = PatternFill('solid', fgColor='F2F2F2')
GREEN = PatternFill('solid', fgColor='E2F0D9')


def prepare(ss, names, ep_items):
    for key in SAVE_TABLES:
        ss.setdefault(key, {})
    ss.setdefault('thermal_included', {})
    ss.setdefault('legacy_warnings', [])
    if ss.get('methodology_revision') != REVISION:
        for name, old in ss.get('block_rates', {}).items():
            problems = legacy_issues(old)
            if problems:
                ss.legacy_rates[name] = old.copy()
                ss.legacy_warnings.extend(f'{name}: {p}' for p in problems)
                if 'Код' in old:
                    ss.block_rates[name] = old.copy()
                    ss.block_rates[name].loc[old['Код'] == 'E.20.10', RATE] = 0.0
        if any(number(v) for v in ss.get('block_mp_rate', {}).values()):
            ss.legacy_warnings.append('Прежняя единая ставка МП сохранена в старом шаблоне, но не применяется: заполните ставки A–E по новым базам.')
        ss.methodology_revision = REVISION
    for name in names:
        ss.block_rates[name] = normalise_rates(ss.block_rates.get(name), ep_items, 'EP')
        ss.block_mp_rates[name] = normalise_rates(ss.block_mp_rates.get(name), MP_ITEMS, 'MP')
        ss.mp_sections.setdefault(name, pd.DataFrame(columns=SECTION_COLS))
        ss.ep_details.setdefault(name, pd.DataFrame(columns=DETAIL_COLS))
        ss.thermal_included.setdefault(name, False)
        for key, cols, numeric in [('mp_sections', SECTION_COLS, SECTION_COLS[1:]), ('ep_details', DETAIL_COLS, ['Объем', RATE])]:
            frame = ss[key][name].reindex(columns=cols).copy()
            for col in cols:
                if col in numeric:
                    frame[col] = pd.to_numeric(frame[col], errors='coerce').fillna(0.0)
                else:
                    frame[col] = frame[col].fillna('').astype(str)
            ss[key][name] = frame


def legend(st):
    st.markdown('<style>[class*="st-key-seo_"] {background:#fffbeb;padding:0.75rem;border-radius:0.5rem;}</style>', unsafe_allow_html=True)
    st.markdown('🟨 **СЭО:** ставки, аналог, дата цен и состав работ. 🟦 **Проектировщик:** ТЭП и объёмы. ⬜ **Авторасчёт:** суммы.')
    if st.session_state.get('legacy_warnings'):
        with st.expander('Проверьте перенос ставок прежней методики', expanded=True):
            for msg in st.session_state.legacy_warnings:
                st.warning(msg)
            st.caption('Ставки с изменёнными единицами нельзя переносить автоматически. Архив исходных ставок включён в шаблон выгрузки.')


def _edit(st, df, key, **kwargs):
    edited = st.data_editor(df, key=key, hide_index=True, width='stretch', **kwargs)
    # Widget state owns row editing; equality prevents unnecessary reruns.
    return edited


def render_geometry(st, names):
    name = st.selectbox('Блок — геометрия МП', names, key='mp_geometry_block')
    st.markdown('**Геометрия МП — заполняет проектировщик**')
    st.caption('Секции одинаковой геометрии можно объединить количеством. Стилобат задайте отдельной строкой, без перекрытия пятен секций. Торцы — суммарная исключаемая площадь для строки, каждый закрытый фасад учитывается один раз. Высота этажа для МП — 3,0 м.')
    old = st.session_state.mp_sections[name]
    edited = _edit(st, old, f'mp_sections_editor_{name}', num_rows='dynamic', column_config={
        c: st.column_config.NumberColumn(min_value=0, step=1 if c in ['Количество', 'Этажей', 'Стилобат (1/0)'] else None)
        for c in SECTION_COLS[1:]
    })
    if not edited.equals(old):
        st.session_state.mp_sections[name] = edited
        st.rerun()
    totals, issues = geometry(edited.to_dict('records'))
    for issue in issues:
        st.error(issue)
    labels = ['Застройка, м²', 'Подземная часть, м²', 'Надземная площадь, м²', 'Фасад, м²', 'Секции, шт']
    st.dataframe(pd.DataFrame([dict(zip(labels, totals.values()))]), hide_index=True, width='stretch')


def render_details(st, names):
    name = st.selectbox('Блок — наружные конструкции ЭП', names, key='ep_detail_block')
    st.markdown('**Наружные конструкции B20.10 — объёмы из ЭП, ставки заполняет СЭО**')
    st.caption('Кладку и отделку задавайте отдельными строками по каждому типу фасада. Для корзин и экранов укажите материал и размеры. Входная группа: комплексная ставка включает ограждения и облицовку полов; расчёт за штуку используйте только для групп без площади. Террасы: отделка и оборудование, уже учтённый каркас повторно не включайте.')
    old = st.session_state.ep_details[name]
    # Units are derived from category, never entered by hand.
    display = old.copy()
    display.insert(2, 'Ед.', display['Категория'].map(lambda c: DETAIL_SPEC.get(c, ('', ''))[1]))
    display['Сумма, руб'] = pd.to_numeric(display['Объем'], errors='coerce').fillna(0) * pd.to_numeric(display[RATE], errors='coerce').fillna(0)
    with st.container(border=True, key=f'seo_facades_{name}'):
        edited = _edit(st, display, f'ep_details_editor_{name}', num_rows='dynamic', disabled=['Ед.', 'Сумма, руб'], column_config={
            'Категория': st.column_config.SelectboxColumn(options=list(DETAIL_SPEC), required=True, width='large'),
            'Объем': st.column_config.NumberColumn('Объём · проект', min_value=0),
            RATE: st.column_config.NumberColumn('🟨 Ставка · СЭО', min_value=0),
            'Аналог': st.column_config.TextColumn('🟨 Аналог · СЭО'),
            'Дата цен': st.column_config.TextColumn('🟨 Дата цен · СЭО'),
            'Состав работ / комментарий': st.column_config.TextColumn('🟨 Состав работ · СЭО'),
            'Сумма, руб': st.column_config.NumberColumn(format='localized'),
        })
    result = edited.reindex(columns=DETAIL_COLS)
    if not result.equals(old):
        st.session_state.ep_details[name] = result
        st.rerun()
    amounts, _, issues = detail_totals(result.to_dict('records'))
    st.caption(f'Наружные конструкции B20.10, без сценария: {sum(amounts.values()):,.0f} руб.'.replace(',', ' '))
    for issue in issues:
        st.error(issue)
    included = st.checkbox('ИТП и котельная уже включены в комплексные ставки инженерии', value=bool(st.session_state.thermal_included.get(name)), key=f'thermal_included_widget_{name}')
    if included != st.session_state.thermal_included[name]:
        st.session_state.thermal_included[name] = included
        st.rerun()
    if included:
        st.caption('D30.20 показан в детализации с нулевым начислением. Состав комплексной ставки укажите в комментарии к инженерии.')


def quantities(items, ss, names, is_res, blocks, nsa, get_basis, mp):
    qty, costs_override, rates_override, issues = {}, {}, {}, []
    geoms = {n: geometry(ss.mp_sections[n].to_dict('records')) for n in names if n in ss.mp_sections}
    details = {n: detail_totals(ss.ep_details[n].to_dict('records')) for n in names if n in ss.ep_details}
    for code, label, group, unit, basis, _ in items:
        arr = np.zeros(len(names))
        for i, name in enumerate(names):
            if not is_res[i]:
                continue
            if mp and basis.startswith('MP_'):
                arr[i] = geoms[name][0][basis]
            elif code in DETAIL_CODES:
                arr[i] = details[name][1].get(code, 0)
                amount = details[name][0].get(code, 0)
                costs_override.setdefault(code, np.zeros(len(names)))[i] = amount
                rates_override.setdefault(code, np.zeros(len(names)))[i] = amount / arr[i] if arr[i] else 0
            elif basis in dict((p[0], p) for p in EXTRA_PARAMS):
                arr[i] = number(ss.tep_store[name].get(basis))
                if code.startswith('D.30.20.') and ss.thermal_included.get(name):
                    arr[i] = 0
            elif basis == 'ELEVATOR_STOPS':
                arr[i] = number(ss.block_elevator_stops.get(name))
            else:
                arr[i] = get_basis(basis, blocks, nsa)[i]
        qty[code] = arr
    for i, n in enumerate(names):
        if is_res[i] and nsa[i] > 0 and mp and not geoms[n][0]['MP_FOOTPRINT']:
            issues.append(f'{n}: заполните геометрию секций МП.')
        if is_res[i] and not mp and number(ss.tep_store.get(n, {}).get('Площадь фасада, м2')) > 0 and not details[n][1]:
            issues.append(f'{n}: площадь фасада задана, но детализация B20.10 не заполнена.')
        if n in geoms and mp:
            issues += [f'{n}: {s}' for s in geoms[n][1]]
        elif n in details and not mp:
            issues += [f'{n}: {s}' for s in details[n][2]]
    return qty, costs_override, rates_override, issues


def render_rates(st, items, mp, names, qty, override_rates, group_labels, fmt_mln):
    name = st.selectbox('Блок — ставки A–E', names, key='v2_rates_block')
    store = 'block_mp_rates' if mp else 'block_rates'
    frame = st.session_state[store][name]
    ni = list(st.session_state.blocks_df['Название блока']).index(name)
    st.markdown('**МП · статьи второго уровня**' if mp else '**ЭП · ставки по видам работ**')
    st.caption('Ставки МП и ЭП хранятся раздельно. Продаваемая площадь (ПП) = квартиры с коэффициентами + коммерция + кладовые — действующая база NSA. Ставки вводятся без коэффициента сценария; он применяется к сумме один раз.')
    group = st.segmented_control('Группа A–E', list(group_labels), default='A', format_func=lambda g: f'{g} · {group_labels[g]}', key=f'v2_group_{mp}') or 'A'
    frame = frame.copy()
    frame['База'] = [float(qty[c][ni]) for c in frame['Код']]
    frame['Статус'] = ['Из детализации ЭП' if c in DETAIL_CODES else 'СЭО: заполнить ставку' if frame.iloc[i]['База'] > 0 and not number(frame.iloc[i][RATE]) else 'Проект: заполнить объём' if number(frame.iloc[i][RATE]) > 0 and not frame.iloc[i]['База'] else '' for i, c in enumerate(frame['Код'])]
    for i, c in enumerate(frame['Код']):
        if c in override_rates:
            frame.at[i, RATE] = float(override_rates[c][ni])
    frame['Сумма, руб'] = frame['База'] * pd.to_numeric(frame[RATE], errors='coerce').fillna(0)
    frame.loc[frame['Применимость'] == 'Не требуется', 'Сумма, руб'] = 0
    frame.loc[frame['Применимость'] == 'Не требуется', 'Статус'] = 'Не требуется — подтверждено СЭО'
    group_frame = frame[frame['Группа'] == group].drop(columns=['_basis', 'Группа'])
    manual = group_frame[~group_frame['Код'].isin(DETAIL_CODES)]
    detail = group_frame[group_frame['Код'].isin(DETAIL_CODES)]
    if len(manual):
        with st.container(border=True, key=f'seo_rates_{mp}_{name}'):
            edited = _edit(st, manual, f'v2_rates_{mp}_{name}_{group}', num_rows='fixed',
                disabled=['Код', 'Статья затрат', 'Единица измерения', 'База', 'Статус', 'Сумма, руб'],
                column_config={'Применимость': st.column_config.SelectboxColumn('🟨 Применимость · СЭО', options=['Уточнить','Применяется','Не требуется']), RATE: st.column_config.NumberColumn('🟨 Ставка · СЭО', min_value=0, format='localized'),
                               'Аналог': st.column_config.TextColumn('🟨 Аналог · СЭО'),
                               'Дата цен': st.column_config.TextColumn('🟨 Дата цен · СЭО'),
                               'Состав работ / комментарий': st.column_config.TextColumn('🟨 Состав работ · СЭО'),
                               'Сумма, руб': st.column_config.NumberColumn(format='localized')})
        updated = st.session_state[store][name].copy()
        for _, row in edited.iterrows():
            for c in [RATE, 'Применимость', 'Аналог', 'Дата цен', 'Состав работ / комментарий']:
                updated.loc[updated['Код'] == row['Код'], c] = row[c]
        if not updated.equals(st.session_state[store][name]):
            st.session_state[store][name] = updated
            st.rerun()
    if len(detail):
        st.caption('B20.10 — суммы строк из «Данные МП / ЭП → Наружные конструкции». Ставка здесь — средневзвешенная справочно.')
        st.dataframe(detail, hide_index=True, width='stretch')
    st.metric(f'{group} · сумма без сценария', fmt_mln(float(frame.loc[frame['Группа'] == group, 'Сумма, руб'].sum())))
    with st.expander('Копировать ставки A–E в другие блоки'):
        targets = st.multiselect('Блоки', [n for n in names if n != name], key=f'v2_copy_{mp}')
        if st.button('Копировать ставки', disabled=not targets, key=f'v2_copy_btn_{mp}'):
            for target in targets:
                st.session_state[store][target] = st.session_state[store][name].copy()
            st.rerun()
    return name


def gz_editor(st, frame, key, *, qty=None, base=None, rate_col=RATE):
    display = frame.copy()
    if rate_col == RATE:
        volume = number(qty) if qty is not None else pd.to_numeric(frame['Кол-во'], errors='coerce').fillna(0)
        display['Сумма, руб'] = volume * pd.to_numeric(frame[RATE], errors='coerce').fillna(0)
    elif base is not None:
        display['База, руб'] = base
        display['Сумма, руб'] = base * pd.to_numeric(frame[rate_col], errors='coerce').fillna(0)
    disabled = ['Код', 'Статья затрат'] + [c for c in ['Единица измерения', 'База, руб', 'Сумма, руб'] if c in display and c not in frame]
    if 'Единица измерения' in display and 'Единица измерения' not in disabled:
        disabled.append('Единица измерения')
    cfg = {rate_col: st.column_config.NumberColumn('🟨 СЭО · ' + ('доля от базы' if base is not None else 'ставка'), min_value=0, max_value=1 if base is not None else None, format='localized'),
           'Сумма, руб': st.column_config.NumberColumn(format='localized')}
    edited = st.data_editor(display, key=key, hide_index=True, width='stretch', num_rows='fixed', disabled=disabled, column_config=cfg)
    result = edited.reindex(columns=frame.columns)
    if not result.equals(frame):
        return result, True
    return result, False


def v2_snapshot(ss):
    return {**{k: {n: f.to_dict('records') for n, f in ss.get(k, {}).items()} for k in SAVE_TABLES},
            **{k: ss.get(k) for k in SAVE_PLAIN}}


def add_template_sheets(wb, snapshot):
    specs = [('МП — секции', 'mp_sections', SECTION_COLS),
             ('МП — ставки A–E', 'block_mp_rates', list(rates_frame(MP_ITEMS).columns)),
             ('ЭП — фасады и элементы', 'ep_details', DETAIL_COLS),
             ('Архив прежних ставок', 'legacy_rates', ['Код', 'Статья затрат', RATE]),
             ('ЭП — источники ставок', 'ep_rates_metadata', ["Код", "Применимость", "Аналог", "Дата цен", "Состав работ / комментарий"])]
    for title, key, cols in specs:
        ws = wb.create_sheet(title)
        ws.append(['Блок'] + cols)
        for name, rows in snapshot.get(key, {}).items():
            for row in rows:
                ws.append([name] + [row.get(c) for c in cols])
        if ws.max_row == 1:
            for _ in range(20):
                ws.append([None] * (len(cols) + 1))
        style_input_sheet(ws, seo=key not in ('mp_sections',))
    ws = wb.create_sheet('Настройки методики V2')
    ws.append(['Параметр', 'Блок', 'Значение'])
    ws.append(['Ревизия', '', REVISION])
    for name, value in (snapshot.get('thermal_included') or {}).items():
        ws.append(['ИТП и котельная включены в ставку', name, int(value)])
    style_input_sheet(ws, seo=True)


def restore_template(wb, ss, warns):
    sheets = {'МП — секции': ('mp_sections', SECTION_COLS), 'МП — ставки A–E': ('block_mp_rates', list(rates_frame(MP_ITEMS).columns)),
              'ЭП — фасады и элементы': ('ep_details', DETAIL_COLS), 'Архив прежних ставок': ('legacy_rates', ['Код','Статья затрат',RATE])}
    names = set(ss['blocks_df']['Название блока'])
    for title, (key, cols) in sheets.items():
        ss.setdefault(key, {})
        if title not in wb:
            continue
        ss[key] = {}
        ws = wb[title]
        headers = [c.value for c in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            rec = dict(zip(headers, row))
            name = rec.pop('Блок', None)
            if not name:
                continue
            if name not in names:
                warns.append(f'{title}: неизвестный блок «{name}», строка пропущена.')
                continue
            ss[key].setdefault(name, []).append({c: rec.get(c) for c in cols})
        ss[key] = {n: pd.DataFrame(rows, columns=cols) for n, rows in ss[key].items()}
    ss['thermal_included'] = {}
    if 'Настройки методики V2' in wb:
        ss['methodology_revision'] = REVISION
        for row in wb['Настройки методики V2'].iter_rows(min_row=2, values_only=True):
            if row[0] == 'ИТП и котельная включены в ставку' and row[1] in names:
                ss['thermal_included'][row[1]] = bool(number(row[2]))
    if 'ЭП — источники ставок' in wb:
        for name, code, applies, analog, date, comment in wb['ЭП — источники ставок'].iter_rows(min_row=2, values_only=True):
            if name in ss.get('block_rates', {}):
                frame = ss['block_rates'][name]
                for col, val in [('Применимость', applies or 'Уточнить'), ('Аналог', analog), ('Дата цен', date), ('Состав работ / комментарий', comment)]:
                    frame.loc[frame['Код'] == code, col] = val or ''
    return ss, warns


def style_input_sheet(ws, seo=False):
    ws.freeze_panes = 'C2'
    for c in ws[1]:
        c.fill = PatternFill('solid', fgColor='41AA37')
        c.font = Font(name='Arial', bold=True, color='FFFFFF')
        c.alignment = Alignment(wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            is_seo = seo and ws.cell(1,c.column).value in (RATE, 'Аналог','Дата цен','Состав работ / комментарий','Значение','Применимость')
            c.fill = YELLOW if is_seo else BLUE
            c.comment = Comment('Заполняет СЭО' if is_seo else 'Данные проекта / идентификатор', 'ТАЛАН')
    for col in range(1,ws.max_column+1):
        ws.column_dimensions[get_column_letter(col)].width = 24 if col != 2 else 42


def stage_workbook(wb, ws, ss, items, blocks, res_names, mp, helper, first_row, colmap, factor_ref):
    """Live quantities and amounts; returns row/column coordinates used by existing report."""
    # Input quantities not present in the original hidden helper live on a separate visible sheet.
    inp = wb.create_sheet('ТЭП A–E')
    extra = [p[0] for p in EXTRA_PARAMS]
    inp.append(['Блок'] + extra + ['ИТП включены (1/0)'])
    name_row = {}
    for i, name in enumerate(blocks['Название блока']):
        inp.append([name] + [number(ss.tep_store.get(name, {}).get(k)) for k in extra] + [int(ss.thermal_included.get(name,False))])
        name_row[name] = i+2
    style_input_sheet(inp)
    geom_ws = wb.create_sheet('Геометрия МП')
    geom_ws.append(['Блок'] + SECTION_COLS + ['Застройка, м2','Подземная часть, м2','Надземная площадь, м2','Фасад, м2','Секции, шт'])
    for name in res_names:
        for row in ss.mp_sections[name].to_dict('records'):
            r = geom_ws.max_row+1
            geom_ws.append([name]+[row.get(c) for c in SECTION_COLS])
            for c, formula in enumerate([f'=C{r}*D{r}',f'=I{r}*(1-H{r})',f'=I{r}*E{r}',f'=MAX(0,C{r}*F{r}*E{r}*3-G{r})',f'=C{r}*(1-H{r})'],start=9):
                geom_ws.cell(r,c,formula)
    style_input_sheet(geom_ws)
    for row in geom_ws.iter_rows(min_row=2,min_col=9):
        for c in row: c.fill=GRAY
    det = wb.create_sheet('Детализация ЭП')
    det.append(['Блок','Код','Категория','Тип / материал / размеры','Ед.','Объем',RATE,'Сумма, руб','Аналог','Дата цен','Состав работ / комментарий'])
    for name in res_names:
        for row in ss.ep_details[name].to_dict('records'):
            category=row.get('Категория')
            if category not in DETAIL_SPEC: continue
            code,unit=DETAIL_SPEC[category]
            r=det.max_row+1
            det.append([name,code,category,row.get('Тип / материал / размеры'),unit,number(row.get('Объем')),number(row.get(RATE)),f'=F{r}*G{r}',row.get('Аналог'),row.get('Дата цен'),row.get('Состав работ / комментарий')])
    style_input_sheet(det,seo=True)
    for r in range(2,det.max_row+1): det.cell(r,8).fill=GRAY
    ws['A4']='МП — ставки второго уровня' if mp else 'ЭП — ставки по видам работ'
    ws.append([])
    rates_header=5
    for c,v in enumerate(['Код','Статья','Группа','Единица']+[f'🟨 СЭО · {n}' for n in res_names],1): ws.cell(rates_header,c,v)
    rate_rows={}
    store='block_mp_rates' if mp else 'block_rates'
    for i,(code,label,group,unit,basis,_) in enumerate(items,6):
        rate_rows[code]=i
        for c,v in enumerate([code,label,group,unit],1): ws.cell(i,c,v)
        for j,name in enumerate(res_names,5):
            frame=ss[store][name].set_index('Код')
            v=number(effective_rates(frame.reset_index()).get(code))
            if code in DETAIL_CODES:
                # Informational weighted average; sum itself uses exact detail rows.
                last=max(2,det.max_row)
                v=f'=IFERROR(SUMIFS(\'Детализация ЭП\'!$H$2:$H${last},\'Детализация ЭП\'!$A$2:$A${last},{get_column_letter(j)}$5,\'Детализация ЭП\'!$B$2:$B${last},$A{i})/SUMIFS(\'Детализация ЭП\'!$F$2:$F${last},\'Детализация ЭП\'!$A$2:$A${last},{get_column_letter(j)}$5,\'Детализация ЭП\'!$B$2:$B${last},$A{i}),0)'
                # Header must be exact block name for SUMIFS.
            ws.cell(i,j,v).fill=GRAY if code in DETAIL_CODES else YELLOW
            ws.cell(i,j).comment=Comment('Из детализации ЭП' if code in DETAIL_CODES else 'Заполняет СЭО: ставка на указанную единицу','ТАЛАН')
    # Exact names also make the per-block catalog readable and usable in SUMIFS.
    for j,name in enumerate(res_names,5): ws.cell(5,j,name)
    # Quantity table follows the rates, with one input formula per article and block.
    qty_header=6+len(items)+2
    for j,v in enumerate(['Блок','Тип блока']+[c for c,*_ in items],1): ws.cell(qty_header,j,v)
    qty_first=qty_header+1
    for i,(_,block) in enumerate(blocks.iterrows()):
        r=qty_first+i;name=block['Название блока']; hr=first_row+i
        ws.cell(r,1,name);ws.cell(r,2,block['Тип блока'])
        for j,(code,label,g,u,basis,_) in enumerate(items,3):
            expr='0'
            if name in res_names:
                h=f"'{helper}'!"
                if basis.startswith('MP_'):
                    gi={'MP_FOOTPRINT':'I','MP_UNDERGROUND':'J','MP_ABOVE':'K','MP_FACADE':'L','MP_SECTIONS':'M'}[basis]
                    last=max(2,geom_ws.max_row)
                    expr=f'SUMIF(\'Геометрия МП\'!$A$2:$A${last},A{r},\'Геометрия МП\'!${gi}$2:${gi}${last})'
                elif code in DETAIL_CODES:
                    last=max(2,det.max_row)
                    expr=f'SUMIFS(\'Детализация ЭП\'!$F$2:$F${last},\'Детализация ЭП\'!$A$2:$A${last},A{r},\'Детализация ЭП\'!$B$2:$B${last},{get_column_letter(j)}${qty_header})'
                elif basis in extra:
                    addr=f"'ТЭП A–E'!{get_column_letter(2+extra.index(basis))}{name_row[name]}"
                    expr=f'IF(\'ТЭП A–E\'!{get_column_letter(len(extra)+2)}{name_row[name]}=1,0,{addr})' if code.startswith('D.30.20.') else addr
                elif basis=='NSA':expr=f"{h}{colmap['NSA, м2']}{hr}"
                elif basis=='STORAGE_AREA':expr=f"{h}{colmap['S кладовых, м2']}{hr}"
                elif basis=='VOL_TOTAL':expr=f"{h}{colmap['Объем здания ниже 0, м3']}{hr}+{h}{colmap['Объем здания выше 0, м3']}{hr}"
                elif basis=='ELEVATOR_STOPS':
                    # explicit input in the quantity table
                    expr=None;ws.cell(r,j,number(ss.block_elevator_stops.get(name))).fill=BLUE
                else:
                    mapping={'AREA_FOOTPRINT':'Площадь застройки, м2','VOL_BELOW':'Объем здания ниже 0, м3','VOL_ABOVE':'Объем здания выше 0, м3','AREA_GLAZING_WINDOWS':'Площадь остекления окон, м2','AREA_GLAZING_BALCONY':'Площадь остекления лоджий, м2','APT_COUNT':'Кол-во квартир, шт','ENTRANCE_COUNT':'Кол-во подъездов, шт'}
                    expr=f"{h}{colmap[mapping[basis]]}{hr}"
            if expr is not None:ws.cell(r,j,'='+expr).fill=GRAY
    calc_header=qty_first+len(blocks)+2
    headers=['Название блока','Тип блока']+[c for c,*_ in items]+['ИТОГО СМР коробки, руб']
    for j,v in enumerate(headers,1):ws.cell(calc_header,j,v)
    calc_first=calc_header+1; total_col=len(headers)
    for i,(_,block) in enumerate(blocks.iterrows()):
        r=calc_first+i;qr=qty_first+i;name=block['Название блока']
        ws.cell(r,1,name);ws.cell(r,2,block['Тип блока'])
        for j,(code,*_) in enumerate(items,3):
            formula='=0'
            if name in res_names:
                rate_col=get_column_letter(5+res_names.index(name))
                formula=f'={get_column_letter(j)}{qr}*{rate_col}{rate_rows[code]}'
                if code in DETAIL_CODES:
                    last=max(2,det.max_row)
                    formula=f'=SUMIFS(\'Детализация ЭП\'!$H$2:$H${last},\'Детализация ЭП\'!$A$2:$A${last},A{r},\'Детализация ЭП\'!$B$2:$B${last},{get_column_letter(j)}${calc_header})'
            ws.cell(r,j,formula).fill=GRAY
        ws.cell(r,total_col,f'=SUM(C{r}:{get_column_letter(total_col-1)}{r})*{factor_ref}').fill=GREEN
    for r in [rates_header,qty_header,calc_header]:
        for c in ws[r]:
            c.font=Font(name='Arial',bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='41AA37')
    ws.column_dimensions['A'].width=22;ws.column_dimensions['B'].width=48
    return calc_header,calc_first,calc_first+len(blocks)-1,total_col,calc_first+len(blocks)-1
