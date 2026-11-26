"""Agreed review changes: one input source, explicit facade modes and readable export."""
from uuid import uuid4
from copy import deepcopy
import math
import re
from numbers import Real
import numpy as np
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter as xlcol
from methodology_v2 import *
from report_presentation import report_table, PERCENT

STATUSES = ['Нет данных', 'Рассчитан', 'Не требуется']
HEAT_KEY = 'ИТП / котельная, комплектов на этап'
COMMENT = 'Состав работ / комментарий'


def normalise_details(frame):
    frame = frame.copy().reindex(columns=DETAIL_COLS).astype(object)
    next_type = max([int(str(v).split()[-1]) for v in frame['Тип'].dropna() if str(v).startswith('ТИП ') and str(v).split()[-1].isdigit()] or [0])
    for i in frame.index:
        for c in ['Категория','Тип / материал / размеры', COMMENT,'Аналог','Дата цен','Тип']:
            frame.at[i,c] = '' if pd.isna(frame.at[i,c]) else str(frame.at[i,c])
        if pd.isna(frame.at[i,'ID']) or not str(frame.at[i,'ID']):frame.at[i,'ID']=str(uuid4())
        if frame.at[i,'Категория']=='Отделка фасада' and not frame.at[i,'Тип']:
            next_type+=1;frame.at[i,'Тип']=f'ТИП {next_type}'
        for c in ['Объем',RATE]:frame.at[i,c]=number(frame.at[i,c])
        if frame.at[i,'Статус'] not in STATUSES:
            frame.at[i,'Статус']='Рассчитан' if frame.at[i,'Объем']>0 and frame.at[i,RATE]>0 else 'Нет данных'
    return frame


def prepare(ss, names, ep_items):
    for key in ['block_mp_rates','mp_sections','ep_details','legacy_rates','thermal_kind','facade_modes','thermal_included']:
        ss.setdefault(key,{})
    ss.setdefault('legacy_warnings',[])
    for name in names:
        tep=ss.tep_store.setdefault(name,{})
        if name not in ss.thermal_kind:
            itp,boiler=number(tep.get('ИТП, комплектов на этап')),number(tep.get('Котельная, комплектов на этап'))
            ss.thermal_kind[name]='Котельная' if boiler and not itp else 'ИТП'
            tep.setdefault(HEAT_KEY,boiler if boiler and not itp else itp)
            if itp and boiler:
                msg=f'{name}: в прежних данных заданы и ИТП, и котельная. Выберите один вариант в «Объёмах».'
                if msg not in ss.legacy_warnings:ss.legacy_warnings.append(msg)
        tep.setdefault(HEAT_KEY,0.0)
        old=ss.block_rates.get(name)
        if ss.get('methodology_revision')!=3 and old is not None:
            ss.legacy_rates.setdefault(name,old.copy())
        if old is not None and 'Код' in old and 'D.30.20' not in set(old['Код']):
            selected='D.30.20.02' if ss.thermal_kind[name]=='Котельная' else 'D.30.20.01'
            match=old[old['Код']==selected].copy()
            if len(match):
                match['Код']='D.30.20';old=pd.concat([old,match],ignore_index=True)
        ss.block_rates[name]=normalise_rates(old,ep_items,'EP')
        ss.block_rates[name]['Применимость']='Применяется'
        ss.block_mp_rates[name]=normalise_rates(ss.block_mp_rates.get(name),MP_ITEMS,'MP')
        ss.block_mp_rates[name]['Применимость']='Применяется'
        ss.mp_sections.setdefault(name,pd.DataFrame(columns=SECTION_COLS))
        ss.ep_details[name]=normalise_details(ss.ep_details.get(name,pd.DataFrame(columns=DETAIL_COLS)))
        ss.facade_modes.setdefault(name, 'Укрупнённый')
    ss.methodology_revision=3


def facade_cost(rows, area, rate, mode):
    """No hidden top-up: aggregate OR exact element sum. Unknown is never no-work."""
    if not rows or mode=='Укрупнённый':
        return number(area)*number(rate),([] if number(area)>0 and number(rate)>0 else ['B20.10: заполните общую площадь фасада и комплексную ставку.'])
    total=0.;issues=[]
    for i,row in enumerate(rows):
        if row.get('Категория') not in DETAIL_SPEC:
            issues.append(f'B20.10, строка {i+1}: выберите элемент.');continue
        status=row.get('Статус','Нет данных');q=number(row.get('Объем'));r=number(row.get(RATE))
        if status=='Не требуется':continue
        if status!='Рассчитан':issues.append(f"B20.10 · {row['Категория']} {row.get('Тип','')}: нет данных.");continue
        if q<=0 or r<=0:
            issues.append(f"B20.10 · {row['Категория']}: статус «Рассчитан» требует объём и ставку.");continue
        if DETAIL_SPEC[row['Категория']][1]=='шт' and q%1:
            issues.append(f"B20.10 · {row['Категория']}: количество должно быть целым.");continue
        total+=q*r
    return total,issues


NUMBER_SEPARATORS = str.maketrans({',': ' ', '.': ','})
NUMBER_INPUT = r'^[+-]?(?:(?:[0-9]+|[0-9]{1,3}(?:[ \u00a0\u202f][0-9]{3})+)(?:[,.][0-9]+)?|[,.][0-9]+)$'


def active_facade_cost(rows, area, rate, mode):
    if mode == 'Детальный' and not rows:
        return 0.0, ['B20.10: добавьте элементы или примените укрупнённый расчёт.']
    return facade_cost(rows, area, rate, mode)


def format_number(value):
    if pd.isna(value):
        return '—'
    return f'{value:,.2f}'.translate(NUMBER_SEPARATORS)


def _column_config(frame, config, disabled=(), readonly=False):
    result = dict(config or {})
    locked_columns = set(disabled) if isinstance(disabled, (list, tuple, set)) else set()
    for column in frame.columns:
        if column in result and result[column] is None:
            continue
        settings = dict(result.get(column) or {})
        label = settings.get('label') or str(column)
        label = re.sub(r'^[✎🔒🟨🟦⬜\s]+', '', label).replace(' · СЭО', '').replace('СЭО · ', '').replace(' · проект', '')
        locked = disabled is True or column in locked_columns or settings.get('disabled') is True
        settings['label'] = label if readonly else ('🔒 ' if locked else '✎ ') + label
        if column == frame.columns[0] and column in ('Код', 'Категория', 'Параметр'):
            settings['pinned'] = True
        result[column] = settings
    return result


def _numeric_formats(frame, config):
    for column, settings in config.items():
        if column not in frame or settings is None:
            continue
        type_config = settings.get('type_config') or {}
        values = frame[column].dropna()
        numeric = not values.empty and all(isinstance(value, Real) and not isinstance(value, (bool, np.bool_)) for value in values)
        if type_config.get('type') == 'number' or (not type_config and numeric):
            yield column, type_config


def reset_editors(st):
    st.session_state['_editor_generation'] = st.session_state.get('_editor_generation', 0) + 1
    st.session_state.pop('_dynamic_editor_sources', None)


def _table_values(frame, columns):
    return tuple(
        tuple(None if pd.isna(value) else value for value in row)
        for row in frame[columns].itertuples(index=False, name=None)
    )


def editor(st, data, *, row_identity=None, **kwargs):
    if not isinstance(data, pd.DataFrame):
        return st.data_editor(data, **kwargs)
    config = _column_config(data, kwargs.pop('column_config', None), kwargs.get('disabled', []))
    fields = dict(_numeric_formats(data, config))
    key = kwargs.get('key', 'table')
    generation = st.session_state.get('_editor_generation', 0)
    widget_key = f'{key}_{generation}'
    if row_identity is not None:
        widget_key += '_' + str(tuple(row_identity))

    # Dynamic grids use their input data as part of the Streamlit widget identity.
    # Keep a stable baseline while the frontend submits edits against those rows.
    dynamic = kwargs.get('num_rows') == 'dynamic'
    if dynamic:
        dynamic_sources = st.session_state.setdefault('_dynamic_editor_sources', {})
        editable = [
            column for column in data
            if config.get(column) is not None
            and not config[column].get('disabled')
            and column not in (kwargs.get('disabled') or ())
        ]
        current = _table_values(data, editable)
        source = dynamic_sources.get(key)
        if source is None or source['expected'] != current or source['columns'] != editable:
            source = {'data': data.copy(), 'expected': current, 'columns': editable, 'id': str(uuid4())}
            dynamic_sources[key] = source
        original_frame = source['data']
        widget_key += '_' + source['id']
    else:
        original_frame = data

    view = original_frame.copy()
    for column in fields:
        view[column] = pd.Series(
            [None if pd.isna(value) else format_number(float(value)) for value in original_frame[column]],
            index=original_frame.index, dtype="string",
        )
        settings = config[column]
        settings['alignment'] = 'right'
        settings['type_config'] = {'type': 'text', 'validate': NUMBER_INPUT}

    kwargs['key'] = widget_key
    with st.container(key='manual_table_'+re.sub(r'\W','_',str(key))):
        edited = st.data_editor(view, column_config=config, **kwargs)
        for column, rules in fields.items():
            lower, upper = rules.get('min_value'), rules.get('max_value')
            parsed = []
            for index, value in edited[column].items():
                original = original_frame.at[index, column] if index in original_frame.index else None
                shown = view.at[index, column] if index in view.index else None
                if pd.isna(value) or value == '':
                    parsed.append(original if pd.isna(shown) else None)
                    continue
                if pd.notna(shown) and value == shown:
                    parsed.append(original)
                    continue
                value = str(value).strip()
                raw = re.sub(r'[ \u00a0\u202f]', '', value).replace(',', '.')
                entered = float(raw) if re.fullmatch(NUMBER_INPUT, value) else math.nan
                if (not math.isfinite(entered) or (lower is not None and entered < lower)
                        or (upper is not None and entered > upper)):
                    st.warning(f'«{config[column]["label"][2:]}»: проверьте допустимый диапазон числа.')
                    parsed.append(data.at[index, column] if index in data.index else original)
                else:
                    parsed.append(entered)
            dtype = data[column].dtype
            series = pd.Series(parsed, index=edited.index, dtype=object)
            if pd.api.types.is_numeric_dtype(dtype) and not pd.api.types.is_bool_dtype(dtype):
                series = pd.to_numeric(series)
                if pd.api.types.is_float_dtype(dtype) or (series.notna().all() and (series % 1 == 0).all()):
                    series = series.astype(dtype)
            edited[column] = series
    if dynamic:
        changes = st.session_state[widget_key]
        if changes['added_rows'] or changes['deleted_rows']:
            dynamic_sources.pop(key)
            st.session_state['_refresh_after_row_change'] = True
        else:
            source['expected'] = _table_values(edited, editable)
    return edited


def readonly(st, data, *, percent_columns=(), **kwargs):
    frame = data if isinstance(data, pd.DataFrame) else getattr(data, 'data', None)
    if not isinstance(frame, pd.DataFrame):
        return st.dataframe(data, **kwargs)
    config = _column_config(frame, kwargs.pop('column_config', None), readonly=True)
    numeric_columns = dict(_numeric_formats(frame, config))
    formats = {}
    for column in numeric_columns:
        if column in percent_columns:
            formats[column] = lambda value: format_number(value * 100) + ' %'
        else:
            formats[column] = format_number
        config[column].pop('type_config', None)
        config[column]['alignment'] = 'right'

    def row_style(row):
        total = any(isinstance(value, str) and value.strip().lower() == 'итого' for value in row)
        styles = ['background-color: #eaf3e8; font-weight: bold' if total else
                  'background-color: #f7f9f7' if row.name in alternate_rows else '' for _ in row]
        for i, column in enumerate(row.index):
            if column in ('Валовая прибыль', 'Валовая рентабельность') and pd.notna(row[column]) and row[column] < 0:
                styles[i] += '; color: #9b3654'
        return styles

    alternate_rows = set(frame.index[1::2])
    st.caption('Только просмотр')
    styled = frame.style.format(formats, na_rep='—').apply(row_style, axis=1)
    return st.dataframe(styled, column_config=config, **kwargs)


def legend(st):
    st.markdown('<style>[class*="st-key-manual_table_"] {background:#fffbed;padding:0.65rem;border:1px solid #e2e3dc;border-radius:0.5rem;}</style>',unsafe_allow_html=True)
    st.caption('✎ Поле для ввода · 🔒 Поле только для просмотра')
    if st.session_state.get('legacy_warnings'):
        with st.expander('Проверка переноса прежних данных'):
            for msg in st.session_state.legacy_warnings:st.warning(msg)


def render_geometry(st, name):
    st.caption(f'Урбан-блок · {name}')
    st.markdown('**Геометрия МП**')
    st.info(
        '**Фасад МП = Σ (количество × периметр × этажность × 3,00 м) − сблокированные торцы.**\n\n'
        'Стилобат задаётся отдельно. Площади секций и стилобата не должны перекрываться. '
        'Закрытые торцы исключаются один раз. Одинаковые секции можно объединить количеством.'
    )
    old=st.session_state.mp_sections[name]
    edited=editor(st,old,key='mp_geometry_'+name,num_rows='dynamic',hide_index=True,width='stretch',column_config={c:st.column_config.NumberColumn(min_value=0) for c in SECTION_COLS[1:]})
    st.session_state.mp_sections[name] = edited
    totals,issues=geometry(edited.to_dict('records'))
    labels = {
        'MP_FOOTPRINT': 'Застройка, м²', 'MP_UNDERGROUND': 'Подземная часть, м²',
        'MP_ABOVE': 'Надземная площадь, м²', 'MP_FACADE': 'Фасад, м²', 'MP_SECTIONS': 'Секции, шт',
    }
    readonly(st, pd.DataFrame([totals]).rename(columns=labels), hide_index=True, width='stretch')
    for issue in issues:st.error(issue)


def render_ep_volumes(st,names,params,get,set_):
    columns = {f'volume_{i}': name for i, name in enumerate(names)}
    records = [
        {'Параметр': label, 'Ед.': unit, **{col: get(name, key) for col, name in columns.items()}}
        for key, label, unit in params
    ]
    config = {col: st.column_config.NumberColumn(name, min_value=0) for col, name in columns.items()}
    edited = editor(
        st, pd.DataFrame(records), key='ep_common_volumes', row_identity=names, hide_index=True, width='stretch',
        disabled=['Параметр', 'Ед.'], column_config=config,
    )
    for i, (key, _, _) in enumerate(params):
        for column, name in columns.items():
            value = number(edited.iloc[i][column])
            if value != get(name, key):
                set_(name, key, value)

    st.markdown('**Ед. для строки «ИТП / котельная на этап»**')
    slots = st.columns(min(len(names), 4))
    options = ['ИТП', 'Котельная']
    for i, name in enumerate(names):
        current = st.session_state.thermal_kind[name]
        selected = slots[i % len(slots)].selectbox(
            name, options, index=options.index(current),
            format_func=lambda kind: f'{kind}, компл.', key=f'ep_thermal_kind_{name}_{current}',
        )
        st.session_state.thermal_kind[name] = selected


def render_details(st, name):
    st.caption(f'Урбан-блок · {name}')
    st.markdown('**Наружные конструкции B20.10 · объёмы**')
    old=st.session_state.ep_details[name]
    cols=['ID','Категория','Тип','Тип / материал / размеры','Объем']
    if len(old):
        view=old[cols].copy();view.insert(4,'Ед.',old['Категория'].map(lambda c:DETAIL_SPEC.get(c,('',''))[1]))
        view['Убрать']=False
        ed=editor(st,view.drop(columns='ID'),key='facade_volumes_v3_'+name,row_identity=old['ID'],hide_index=True,width='stretch',num_rows='fixed',disabled=['Категория','Тип','Ед.'],column_config={'Объем':st.column_config.NumberColumn(min_value=0)})
        updated=old.copy()
        for i in old.index:
            updated.at[i,'Объем']=number(ed.at[i,'Объем']);updated.at[i,'Тип / материал / размеры']=str(ed.at[i,'Тип / материал / размеры'] or '')
        updated=updated.loc[~ed['Убрать'].astype(bool)].reset_index(drop=True)
        st.session_state.ep_details[name] = updated
        if len(updated) != len(old):
            st.session_state['_refresh_after_row_change'] = True
        old = updated
    else:st.info('Разбивка не задана: B20.10 считается укрупнённо по общей площади фасада.')
    c1,c2=st.columns([3,1])
    with c1:category=st.selectbox('Элемент',list(DETAIL_SPEC),key='new_facade_category_'+name)
    with c2:
        if st.button('Добавить элемент',key='add_facade_'+name):
            typ=''
            if category=='Отделка фасада':
                numbers=[int(str(v).split()[-1]) for v in old['Тип'] if str(v).startswith('ТИП ') and str(v).split()[-1].isdigit()]
                typ=f'ТИП {max(numbers or [0])+1}'
            row={c:'' for c in DETAIL_COLS};row.update({'ID':str(uuid4()),'Категория':category,'Тип':typ,'Объем':0.,RATE:0.,'Статус':'Нет данных'})
            st.session_state.ep_details[name]=pd.concat([old,pd.DataFrame([row])],ignore_index=True)
            st.rerun()
    st.caption('Отделка фасада: каждый вариант добавляется отдельной строкой ТИП 1, ТИП 2 и далее. Ставки, статусы и применение расчёта — в «Коробка A–E → B».')


def quantities(items,ss,names,is_res,blocks,nsa,get_basis,mp):
    qty={};overrides={};rates_override={};issues=[]
    for code,label,group,unit,basis,_ in items:
        arr=np.zeros(len(names))
        for i,name in enumerate(names):
            if not is_res[i]:continue
            if mp and basis.startswith('MP_'):
                values,errs=geometry(ss.mp_sections[name].to_dict('records'));arr[i]=values[basis]
                for e in errs:
                    msg=f'{name}: {e}'
                    if msg not in issues:issues.append(msg)
            elif basis in [p[0] for p in EXTRA_PARAMS]:arr[i]=number(ss.tep_store[name].get(basis))
            elif basis=='ELEVATOR_STOPS':arr[i]=number(ss.block_elevator_stops.get(name))
            elif code=='B.20.10':
                arr[i]=number(ss.tep_store[name].get('Площадь фасада, м2'))
                r=ss.block_rates[name].set_index('Код')[RATE].get(code,0)
                total,errs=active_facade_cost(ss.ep_details[name].to_dict('records'),arr[i],r,ss.facade_modes[name])
                overrides.setdefault(code,np.zeros(len(names)))[i]=total
                rates_override.setdefault(code,np.zeros(len(names)))[i]=total/arr[i] if arr[i] else 0
                issues.extend(f'{name}: {e}' for e in errs)
            else:arr[i]=get_basis(basis,blocks,nsa)[i]
        qty[code]=arr
    return qty,overrides,rates_override,issues


def render_rates(st, items, mp, names, qty, override_rates, group_labels, fmt_mln, *, name=None):
    name = name or names[0]
    identity = st.session_state.block_ids[name]
    st.caption(f'Урбан-блок · {name}')
    store='block_mp_rates' if mp else 'block_rates';old=st.session_state[store][name]
    ni=list(st.session_state.blocks_df['Название блока']).index(name)
    group=st.segmented_control('Группа A–E',list(group_labels),default='A',format_func=lambda g:f'{g} · {group_labels[g]}',key=f'v3_group_{mp}') or 'A'
    view=old[old['Группа']==group][['Код','Статья затрат','Единица измерения',RATE]].copy()
    view['База']=[qty[c][ni] for c in view['Код']]
    view['Сумма, руб']=[float(qty[c][ni])*number(r) for c,r in zip(view['Код'],view[RATE])]
    if not mp and group == 'B':
        view = view[view['Код'] != 'B.20.10']
    if st.checkbox('Только строки, требующие заполнения', key='rate_filter_' + str(mp)):
        missing = view['База'].gt(0) ^ view[RATE].gt(0)
        view = view[missing]
    if len(view):
        ed=editor(st,view,key=f'rates3_{mp}_{name}_{group}',row_identity=view['Код'],hide_index=True,width='stretch',disabled=[c for c in view if c!=RATE],column_config={RATE:st.column_config.NumberColumn(min_value=0)})
        updated=old.copy()
        for _,row in ed.iterrows():updated.loc[updated['Код']==row['Код'],RATE]=number(row[RATE])
        if not updated.equals(old):st.session_state[store][name]=updated;st.rerun()
    else:
        st.caption('В этой группе нет строк для отображения.')
    if not mp and group=='B':
        details = st.session_state.ep_details[name]
        active_mode = st.session_state.facade_modes[name]
        area = number(st.session_state.tep_store[name].get('Площадь фасада, м2'))
        st.markdown(f'**B20.10 · применяется {active_mode.lower()} расчёт**')
        if not len(details):
            st.caption('Детализация добавляется в «Объёмах».')
        default_view = 'Детальный' if len(details) else 'Укрупнённый'
        editor_mode = st.radio('Редактировать B20.10', ['Укрупнённый', 'Детальный'],
                               index=['Укрупнённый', 'Детальный'].index(default_view),
                               horizontal=True, key=f'facade_editor_{identity}')
        if editor_mode == 'Укрупнённый':
            frame = old.loc[old['Код'] == 'B.20.10', ['Код', RATE]].copy()
            frame['Площадь фасада, м²'] = area
            ed = editor(st, frame, key='facade_aggregate3_' + name, hide_index=True,
                        width='stretch', disabled=['Код', 'Площадь фасада, м²'],
                        column_config={RATE: st.column_config.NumberColumn(min_value=0)})
            value = number(ed.iloc[0][RATE])
            if value != number(frame.iloc[0][RATE]):
                old = old.copy()
                old.loc[old['Код'] == 'B.20.10', RATE] = value
                st.session_state[store][name] = old
                st.rerun()
        elif len(details):
            cols = ['Категория', 'Тип', 'Тип / материал / размеры', 'Объем', RATE, 'Статус']
            ed = editor(st, details[cols], key='facade_rates3_' + name, row_identity=details['ID'], hide_index=True, width='stretch',
                        disabled=cols[:4], column_config={
                            RATE: st.column_config.NumberColumn(min_value=0),
                            'Статус': st.column_config.SelectboxColumn(options=STATUSES, required=True),
                        })
            new = details.copy()
            new[RATE] = ed[RATE]
            new['Статус'] = ed['Статус']
            if not new.equals(details):
                st.session_state.ep_details[name] = new
                st.rerun()
        rows = details.to_dict('records')
        rate = old.set_index('Код')[RATE].get('B.20.10')
        active_total, active_issues = active_facade_cost(rows, area, rate, active_mode)
        preview_total, preview_issues = active_facade_cost(rows, area, rate, editor_mode)
        c1, c2 = st.columns(2)
        c1.metric('Текущая сумма без сценария', fmt_mln(active_total))
        c2.metric('Сумма выбранного режима без сценария', fmt_mln(preview_total))
        for msg in dict.fromkeys(active_issues + preview_issues):
            st.warning(msg)
        if editor_mode != active_mode:
            label = ('Применить детальный расчёт' if editor_mode == 'Детальный'
                     else 'Применить укрупнённый расчёт')
            if st.button(label, key=f'apply_facade_{name}', disabled=bool(preview_issues)):
                st.session_state.facade_modes[name] = editor_mode
                st.rerun()
    with st.expander(('Инженерия · комментарии к статьям' if group=='D' else 'Состав ставки')+' · необязательно'):
        st.caption('Уточнение границ состава ставки')
        comments=old[old['Группа']==group][['Код','Статья затрат',COMMENT]].copy()
        ed=editor(st,comments,key=f'comments3_{mp}_{name}_{group}',hide_index=True,width='stretch',disabled=['Код','Статья затрат'])
        updated=old.copy()
        for _,row in ed.iterrows():updated.loc[updated['Код']==row['Код'],COMMENT]=str(row[COMMENT] or '')
        st.session_state[store][name] = updated
    with st.expander('Копировать ставки A–E в другие блоки'):
        targets=st.multiselect('Блоки назначения',[n for n in names if n!=name],key=f'copy_rates3_{mp}')
        if st.button('Копировать ставки',key=f'copy_rates3_btn_{mp}',disabled=not targets):
            for target in targets:st.session_state[store][target]=st.session_state[store][name].copy()
            reset_editors(st)
            st.rerun()
    return name


def add_review_report(wb,ss,blocks,items,mp,item_costs,qty,factor,helper,colmap,first_row,leaf_formula,g_area,g_length,z_pct,z_fixed):
    """Reader layouts over the actual model. Never creates demonstration inputs."""
    names=list(blocks['Название блока']);res=list(ss["block_rates"])
    factor_ref=f"'{helper}'!$E$4"
    def init(name, headers):
        ws = wb.create_sheet(name)
        for col, label in enumerate(headers, 1):
            ws.cell(5, col, label)
        ws.freeze_panes = 'D6'
        return ws

    def finish(ws, *, totals=(), gutters=()):
        report_table(ws, 5, ws.max_row, 1, ws.max_column, totals=totals, gutters=gutters)
        for i in range(1, ws.max_column + 1):
            ws.column_dimensions[xlcol(i)].width = 22
        ws.column_dimensions['B'].width = 38

    fac=init('Элементы B20.10',['🔒 Блок','🔒 Код','🔒 Элемент','🔒 Тип','🔒 Уточнение','🔒 Ед.','🔒 Объём','🔒 Ставка, руб./ед.','🔒 Статус','🔒 Сумма без сценария, руб','🔒 Со сценарием, руб'])
    facade_parent={};detail_source={}
    for name in names:
        if name not in ss["ep_details"]:continue
        if facade_parent:fac.append(['']);fac.row_dimensions[fac.max_row].height=14
        parent=fac.max_row+1;facade_parent[name]=parent
        details=ss["ep_details"][name];mode=ss["facade_modes"][name]
        rate=number(ss["block_rates"][name].set_index('Код')[RATE].get('B.20.10'))
        area=number(ss["tep_store"][name].get('Площадь фасада, м2'))
        amount,issues=active_facade_cost(details.to_dict('records'),area,rate,mode)
        fac.append([name,'B20.10','Наружные стены и фасадные элементы',mode,'Частичная сумма' if issues else '', 'м² фасада',area,rate,'Нет данных' if issues else 'Рассчитан',None,f'=J{parent}*{factor_ref}'])
        start=parent+1
        for _,r in details.iterrows():
            category=r['Категория'];code,unit=DETAIL_SPEC.get(category,('',''))
            row=fac.max_row+1
            valid=r['Статус']=='Рассчитан' and number(r['Объем'])>0 and number(r[RATE])>0 and (unit!='шт' or number(r['Объем'])%1==0)
            fac.append([name,code,category,r['Тип'],r['Тип / материал / размеры'],unit,number(r['Объем']),number(r[RATE]),r['Статус'],f'=G{row}*H{row}' if valid else 0,f'=J{row}*{factor_ref}' if mode=='Детальный' else 0])
            detail_source[(name,str(r['ID']))]=row
        end=fac.max_row
        fac.cell(parent,10,(f'=SUM(J{start}:J{end})' if end>=start else '=0') if mode=='Детальный' else f'=G{parent}*H{parent}')
        if end>=start:
            fac.row_dimensions.group(start,end,hidden=True);fac.row_dimensions[parent].collapsed=True
    finish(fac, totals=set(facade_parent.values()));fac.column_dimensions['C'].width=38;fac.column_dimensions['E'].width=38
    fac.sheet_properties.outlinePr.summaryBelow=False
    # Keep the existing formula engine linked to the same exact per-element calculation.
    det=wb['Детализация ЭП'];counters={}
    for r in range(2,det.max_row+1):
        name=det.cell(r,1).value
        if name not in ss["ep_details"]:continue
        idx=counters.get(name,0);counters[name]=idx+1
        if idx<len(ss["ep_details"][name].loc[ss["ep_details"][name]['Категория'].isin(DETAIL_SPEC)]):
            key=(name,str(ss["ep_details"][name].loc[ss["ep_details"][name]['Категория'].isin(DETAIL_SPEC)].iloc[idx]['ID']));source=detail_source.get(key)
            if source:det.cell(r,8,f"='Элементы B20.10'!J{source}")
    # Common article list; each block is a separate horizontal section with a gutter.
    headers=['🔒 Код','🔒 Статья затрат','🔒 Группа','']
    for name in names:headers += ['🔒 Ед. / база','🔒 Объём / база','🔒 Ставка / доля','🔒 Состав ставки / уточнение',f'🔒 Сумма · {name}, руб','']
    budget=init('Бюджет по статьям',headers);budget.freeze_panes='E6'
    all_items=[(c,n,g,'AE',u) for c,n,g,u,_,_ in items]
    all_items += [('PARK','СМР паркинга','Паркинг','PARK','компл.')]
    all_items += [(c,n,'G','G_AREA','га участка') for c,n,*_ in g_area]+[(c,n,'G','G_LENGTH','п.м.') for c,n,*_ in g_length]
    all_items += [(c,n,'Z','Z_PCT','доля от базы') for c,n,*_ in z_pct]+[(c,n,'Z','Z_FIXED','прямой ввод') for c,n,*_ in z_fixed]+[('Z.50.10','Инфраструктура блока','Инфраструктура','INFRA','компл.')]
    matrix=blocks.set_index('Название блока');store='block_mp_rates' if mp else 'block_rates'
    for code,label,group,kind,unit in all_items:
        row=budget.max_row+1;budget.append([code,label,group])
        for i,name in enumerate(names):
            start=5+i*6;rate=0.;volume=0.;comment='';amount=0.;u=unit
            if kind=='AE' and name in ss[store]:
                rr=ss[store][name].set_index('Код').loc[code];rate=number(rr[RATE]);comment=str(rr.get(COMMENT,'') or '');volume=float(qty[code][i]);amount=f'={xlcol(start+1)}{row}*{xlcol(start+2)}{row}*{factor_ref}'
                if code=='D.30.20':u=ss["thermal_kind"][name]+', компл.'
                if code=='B.20.10':amount=f"='Элементы B20.10'!K{facade_parent[name]}";comment=(comment+' · '+ss["facade_modes"][name]).strip(' ·')
            elif kind.startswith('G_') or kind.startswith('Z_'):
                amount=leaf_formula((kind,code),name)
                key={'G_AREA':'block_g_area','G_LENGTH':'block_g_length','Z_PCT':'block_z_pct','Z_FIXED':'block_z_fixed'}[kind]
                rr=ss[key][name].set_index('Код').loc[code]
                if kind=='G_AREA':volume=number(matrix.loc[name,'Площадь участка блока, га']);rate=number(rr[RATE])
                elif kind=='G_LENGTH':volume=number(rr['Кол-во']);rate=number(rr[RATE])
                elif kind=='Z_PCT':rate=number(rr['Ставка, доля от СМР+G']);volume='СМР+G+паркинг'
                else:volume=1;rate=number(rr.get('Сумма, руб'))
            elif kind=='INFRA':
                volume=1;rate=float(pd.to_numeric(ss["block_infra"][name]['Сумма, руб'],errors='coerce').fillna(0).sum());amount=f'={xlcol(start+2)}{row}*{factor_ref}';comment='Расшифровка: «Инфраструктура — состав»'
            elif kind=='PARK':
                volume=1
                block=matrix.loc[name]
                rate=(number(block['Прямые затраты'])-number(block['Наружные работы блока (G), руб'])-number(block['Прочие затраты блока (Z), руб'])-number(block['Инфраструктура блока (Z5010), руб'])-number(block['Себестоимость коробки (методика)']))/factor
                amount=f'={xlcol(start+2)}{row}*{factor_ref}'
            for j,value in enumerate([u,volume,rate,comment,amount]):budget.cell(row,start+j,value)
            if kind=='Z_PCT':budget.cell(row,start+2).number_format=PERCENT
    last_article=budget.max_row
    total_row=last_article+1;budget.cell(total_row,2,'ИТОГО прямые затраты')
    for i in range(len(names)):
        c=9+i*6;budget.cell(total_row,c,f'=SUM({xlcol(c)}6:{xlcol(c)}{last_article})')
    finish(budget, totals={total_row}, gutters={4 + i * 6 for i in range(len(names) + 1)})
    for i,name in enumerate(names):
        start=5+i*6;last=start+4;budget.column_dimensions.group(xlcol(start),xlcol(last-1),hidden=True);budget.column_dimensions[xlcol(last)].collapsed=True
        budget.column_dimensions[xlcol(last+1)].width=3
        for r in range(6,budget.max_row+1):budget.cell(r,last+1).fill=PatternFill(fill_type=None)
        budget.column_dimensions[xlcol(start+3)].width=42
    budget.column_dimensions['D'].width=3;budget.sheet_properties.outlinePr.summaryRight=True
    # General block areas only: equipment remains in the article budget.
    bl=init('Блоки',['🔒 Блок','🔒 Тип','🔒 Квартиры, м²','🔒 Коммерция, м²','🔒 Кладовые, м²','🔒 NSA, м²','🔒 Участок, га','🔒 Квартир, шт.'])
    for _,b in blocks.iterrows():bl.append([b['Название блока'],b['Тип блока'],b['S квартир, м2'],b['S коммерции 1 эт., м2'],b['S кладовых, м2'],b['NSA, м2'],b['Площадь участка блока, га'],number(ss["tep_store"].get(b['Название блока'],{}).get('Кол-во квартир, шт'))])
    finish(bl)
    ind=init('Косвенные',['🔒 Раздел','🔒 Статья','🔒 База / способ','🔒 Объём','🔒 Ставка / сумма, руб','🔒 Со сценарием, руб']+[f'🔒 {n}, руб' for n in names]);parents=[]
    indirect_labels={'land':'Земля','landscape':'Благоустройство мест общего пользования','infra':'Магистральные сети','social':'Социальные объекты','soft':'Прочие Soft Costs'}
    for category_key,frame in ss["indirect_items"].items():
        category=indirect_labels.get(category_key,category_key)
        parent=ind.max_row+1;parents.append(parent);ind.append([category,'Итого раздела','',None,None,None]);start=parent+1
        for _,r in frame.iterrows():
            ri=ind.max_row+1;ind.append([category,r.get('Наименование',''),'Прямой ввод',1,number(r.get('Сумма, руб')),f'=D{ri}*E{ri}*{factor_ref}'])
            for j,name in enumerate(names,7):ind.cell(ri,j,f'=F{ri}*{float(matrix.loc[name,"Доля аллокации"])}')
        end=ind.max_row
        if end>=start:
            for c in range(6,7+len(names)):ind.cell(parent,c,f'=SUM({xlcol(c)}{start}:{xlcol(c)}{end})')
            ind.row_dimensions.group(start,end,hidden=True);ind.row_dimensions[parent].collapsed=True
        else:ind.cell(parent,6,0)
    finish(ind, totals=set(parents));ind.sheet_properties.outlinePr.summaryBelow=False
    for name in ['ТЭП A–E','Геометрия МП','Детализация ЭП','Источники ставок A–E','СМР по методике']:
        if name in wb:wb[name].sheet_state='hidden'
