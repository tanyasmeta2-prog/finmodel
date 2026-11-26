"""Version 2 forms, Excel inputs and stage adapter for the existing financial model."""
from copy import deepcopy
from io import BytesIO
import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from methodology_v2 import *

REVISION = 3
SAVE_TABLES = ['block_mp_rates', 'mp_sections', 'ep_details', 'legacy_rates']
SAVE_PLAIN = ['methodology_revision', 'thermal_kind', 'facade_modes', 'block_ids', 'legacy_warnings']
DETAIL_CODES = {v[0] for v in DETAIL_SPEC.values()}
YELLOW = PatternFill('solid', fgColor='FFF2CC')
BLUE = PatternFill('solid', fgColor='DDEBF7')
GRAY = PatternFill('solid', fgColor='F2F2F2')
GREEN = PatternFill('solid', fgColor='E2F0D9')


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
    cfg = {rate_col: st.column_config.NumberColumn('✎ ' + ('доля от базы' if base is not None else 'ставка'), min_value=0, max_value=1 if base is not None else None, format='localized'),
           'Сумма, руб': st.column_config.NumberColumn(format='localized')}
    edited = editor(st, display, key=key, hide_index=True, width='stretch', num_rows='fixed', disabled=disabled, column_config=cfg)
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
    for key in ['thermal_kind','facade_modes','block_ids']:
        for name,value in (snapshot.get(key) or {}).items():
            ws.append([key,name,value])
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
            if row[0] in ['thermal_kind','facade_modes','block_ids'] and row[1] in names:
                ss.setdefault(row[0],{})[row[1]]=row[2]
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
            c.fill = YELLOW
            c.comment = Comment('✎ Редактируемое поле шаблона', 'ТАЛАН')
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
        inp.append([name] + [number(ss["tep_store"].get(name, {}).get(k)) for k in extra] + [0])
        name_row[name] = i+2
    style_input_sheet(inp)
    geom_ws = wb.create_sheet('Геометрия МП')
    geom_ws.append(['Блок'] + SECTION_COLS + ['Застройка, м2','Подземная часть, м2','Надземная площадь, м2','Фасад, м2','Секции, шт'])
    for name in res_names:
        for row in ss["mp_sections"][name].to_dict('records'):
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
        for row in ss["ep_details"][name].to_dict('records'):
            category=row.get('Категория')
            if category not in DETAIL_SPEC: continue
            code,unit=DETAIL_SPEC[category]
            r=det.max_row+1
            det.append([name,code,category,row.get('Тип / материал / размеры'),unit,number(row.get('Объем')),number(row.get(RATE)),f'=F{r}*G{r}' if row.get('Статус')=='Рассчитан' and number(row.get('Объем'))>0 and number(row.get(RATE))>0 and (unit!='шт' or number(row.get('Объем'))%1==0) else 0,row.get('Аналог'),row.get('Дата цен'),row.get('Состав работ / комментарий')])
    style_input_sheet(det,seo=True)
    for r in range(2,det.max_row+1): det.cell(r,8).fill=GRAY
    ws['A4']='МП — ставки второго уровня' if mp else 'ЭП — ставки по видам работ'
    ws.append([])
    rates_header=5
    for c,v in enumerate(['Код','Статья','Группа','Единица']+[f'✎ {n}' for n in res_names],1): ws.cell(rates_header,c,v)
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
            ws.cell(i,j).comment=Comment('Из детализации ЭП' if code in DETAIL_CODES else '✎ Ставка на указанную единицу','ТАЛАН')
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
                    expr=addr
                elif basis=='NSA':expr=f"{h}{colmap['NSA, м2']}{hr}"
                elif basis=='STORAGE_AREA':expr=f"{h}{colmap['S кладовых, м2']}{hr}"
                elif basis=='VOL_TOTAL':expr=f"{h}{colmap['Объем здания ниже 0, м3']}{hr}+{h}{colmap['Объем здания выше 0, м3']}{hr}"
                elif basis=='ELEVATOR_STOPS':
                    # explicit input in the quantity table
                    expr=None;ws.cell(r,j,number(ss["block_elevator_stops"].get(name))).fill=BLUE
                else:
                    mapping={'AREA_FACADE':'Площадь фасада, м2','AREA_FOOTPRINT':'Площадь застройки, м2','VOL_BELOW':'Объем здания ниже 0, м3','VOL_ABOVE':'Объем здания выше 0, м3','AREA_GLAZING_WINDOWS':'Площадь остекления окон, м2','AREA_GLAZING_BALCONY':'Площадь остекления лоджий, м2','APT_COUNT':'Кол-во квартир, шт','ENTRANCE_COUNT':'Кол-во подъездов, шт'}
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
            if code == 'B.20.10' and name in res_names and ss["facade_modes"].get(name)=='Детальный':
                last=max(2,det.max_row)
                formula=f'=SUMIFS(\'Детализация ЭП\'!$H$2:$H${last},\'Детализация ЭП\'!$A$2:$A${last},A{r})'
            ws.cell(r,j,formula).fill=GRAY
        ws.cell(r,total_col,f'=SUM(C{r}:{get_column_letter(total_col-1)}{r})*{factor_ref}').fill=GREEN
    for r in [rates_header,qty_header,calc_header]:
        for c in ws[r]:
            c.font=Font(name='Arial',bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='41AA37')
    ws.column_dimensions['A'].width=22;ws.column_dimensions['B'].width=48
    return calc_header,calc_first,calc_first+len(blocks)-1,total_col,calc_first+len(blocks)-1


# Current UI and calculation adapter agreed in the visual review.
from review_v3 import reset_editors, prepare, legend, render_geometry, render_details, quantities, render_rates, editor, readonly, render_ep_volumes, add_review_report, format_number
