"""Excel entry form and partial updates for the Streamlit model."""
from collections import Counter
from copy import deepcopy
from dataclasses import dataclass
from io import BytesIO
from math import isfinite
from numbers import Real

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from methodology_v2 import SECTION_COLS, DETAIL_COLS, DETAIL_SPEC, RATE
from report_presentation import NUMBER, PERCENT, TITLE_COLOR, TEXT_COLOR, HEADER_COLOR, report_table

SHEETS = ('ТЭП', 'Цены продаж', 'Объёмы', 'Ставки')
MARKER = 'finmodel-sparse-input-v1'
HEADER = 5
FIRST_BLOCK = 5
BLOCK_NAME = 'Название блока'
BLOCK_TYPE = 'Тип блока'
TYPE_RESIDENTIAL = 'Жилой блок'
TYPE_PARKING = 'Наземный/Многоуровневый паркинг'
TYPES = (TYPE_RESIDENTIAL, TYPE_PARKING)


@dataclass(frozen=True)
class Field:
    key: str
    group: str
    label: str
    unit: str
    store: str
    column: str
    code: str = ''
    options: tuple = ()
    applies: tuple = TYPES


def _tables(sheet):
    """Optional repeatable tables share the same four worksheets."""
    if sheet == 'Объёмы':
        geometry = (
            'mp-geometry', 'МП. Геометрия секций и стилобата',
            ('УБ', 'Название секции / стилобата', 'Количество, шт', 'Площадь этажа, м²',
             'Этажей, шт', 'Периметр, м', 'Сблокированные торцы, м²', 'Стилобат'),
        )
        yield geometry
    if sheet in ('Объёмы', 'Ставки'):
        yield (
            'ep-elements', 'ЭП. Детализация наружных конструкций B20.10',
            ('УБ', 'Элемент', 'Уточнение / тип', 'Объём' if sheet == 'Объёмы' else 'Ставка, руб./ед.'),
        )


def _append_repeatable_tables(ws):
    for marker, title, headers in _tables(ws.title):
        title_row = ws.max_row + 3
        ws.cell(title_row, 2, title).font = Font(name='Arial', size=13, bold=True, color=TITLE_COLOR)
        ws.merge_cells(start_row=title_row, start_column=2, end_row=title_row, end_column=len(headers) + 1)
        ws.row_dimensions[title_row].height = 24
        note = ('Укажите название УБ и уникальное название строки геометрии.' if marker == 'mp-geometry'
                else 'Для связи объёма и ставки укажите одинаковые УБ, элемент и уточнение на обоих листах.')
        ws.cell(title_row + 1, 2, note)
        ws.merge_cells(start_row=title_row + 1, start_column=2, end_row=title_row + 1, end_column=len(headers) + 1)
        header = title_row + 2
        for col, label in enumerate(headers, 2):
            ws.cell(header, col, label)
        report_table(ws, header, header, 2, len(headers) + 1)
        for row in range(header + 1, header + 13):
            ws.cell(row, 1, marker)
            ws.row_dimensions[row].height = 21
            for col in range(2, len(headers) + 2):
                cell = ws.cell(row, col)
                cell.fill = PatternFill('solid', fgColor='FFF6D9')
                if col >= 5:
                    cell.number_format = NUMBER
                letter = get_column_letter(col)
                if col > 4 and ws.column_dimensions[letter].width < 18:
                    ws.column_dimensions[letter].width = 22
        choices = ('Да', 'Нет') if marker == 'mp-geometry' else tuple(DETAIL_SPEC)
        if marker == 'mp-geometry':
            formula = '"Да,Нет"'
        else:
            # Inline lists are limited to 255 characters. Use hidden rows below the input area.
            source_row = header + 15
            for row, choice in enumerate(choices, source_row):
                ws.cell(row, 3, choice)
                ws.row_dimensions[row].hidden = True
            formula = f'$C${source_row}:$C${source_row + len(choices) - 1}'
        validation = DataValidation(type='list', formula1=formula,
                                    allow_blank=True, showErrorMessage=True)
        ws.add_data_validation(validation)
        column = 'I' if marker == 'mp-geometry' else 'C'
        validation.add(f'{column}{header + 1}:{column}{header + 12}')


def build_template(fields, names):
    """Build the same empty form offered by the application and distributed separately."""
    wb = Workbook()
    wb.remove(wb.active)
    instructions = (
        'Заполняйте жёлтые ячейки. Пустые ячейки и нули пропускаются при загрузке.',
        'Названия УБ в шапках должны совпадать с модулем. Для нового УБ добавьте столбец с его названием.',
        'Для нового паркинга выберите тип на листе «ТЭП». Доли вводите со знаком %, например 6%.',
    )
    for sheet in SHEETS:
        ws = wb.create_sheet(sheet)
        ws['A1'] = MARKER
        ws['B1'] = sheet
        ws['B1'].font = Font(name='Arial', size=17, bold=True, color=TITLE_COLOR)
        ws.row_dimensions[1].height = 27
        for i, note in enumerate(instructions, 2):
            ws.cell(i, 2, note)
            ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=max(5, 4 + len(names)))
            ws.cell(i, 2).font = Font(name='Arial', size=10, color=TEXT_COLOR)
            ws.cell(i, 2).alignment = Alignment(wrap_text=True, vertical='center')
            ws.row_dimensions[i].height = 20
        ws.append(['Идентификатор', 'Раздел', 'Показатель / статья', 'Ед.', *names])
        for field in fields[sheet]:
            ws.append([field.key, field.group, field.label, field.unit])
        report_table(ws, HEADER, ws.max_row, 2, 4)
        ws.column_dimensions['A'].hidden = True
        for column, width in zip('BCD', (21, 54, 27)):
            ws.column_dimensions[column].width = width
        for col in range(FIRST_BLOCK, FIRST_BLOCK + len(names)):
            ws.column_dimensions[get_column_letter(col)].width = 23
            header = ws.cell(HEADER, col)
            header.fill = PatternFill('solid', fgColor=HEADER_COLOR)
            header.font = Font(name='Arial', bold=True, color=TITLE_COLOR)
            header.alignment = Alignment(wrap_text=True, vertical='center')
            for row, field in enumerate(fields[sheet], HEADER + 1):
                cell = ws.cell(row, col)
                cell.fill = PatternFill('solid', fgColor='FFF6D9')
                cell.font = Font(name='Arial', size=10, color=TEXT_COLOR)
                cell.alignment = Alignment(horizontal='right' if not field.options else 'left', vertical='center')
                cell.number_format = PERCENT if field.unit == '%' else NUMBER
        # Dropdowns are optional guardrails. The importer validates them again.
        last_col = get_column_letter(FIRST_BLOCK + len(names) - 1)
        for row, field in enumerate(fields[sheet], HEADER + 1):
            if field.options:
                validation = DataValidation(type='list', formula1='"' + ','.join(field.options) + '"',
                                            allow_blank=True, showErrorMessage=True,
                                            error='Выберите значение из списка.')
                ws.add_data_validation(validation)
                validation.add(f'E{row}:{last_col}{row}')
        _append_repeatable_tables(ws)
        ws.freeze_panes = 'E6'
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 90
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def load_template(file_bytes):
    """Return this format's workbook, or None so archived projects can use legacy import."""
    try:
        wb = load_workbook(BytesIO(file_bytes), data_only=True)
    except Exception:
        return None
    return wb if any(sheet in wb and wb[sheet]['A1'].value == MARKER for sheet in SHEETS) else None


def _input_value(raw, field):
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        return None
    if field.options:
        text = str(raw).strip()
        if text.replace(',', '.').replace('.', '').strip('0') == '':
            return None
        if text not in field.options:
            raise ValueError('выберите значение из списка')
        return text
    if isinstance(raw, bool):
        raise ValueError('введите число')
    if isinstance(raw, Real):
        value = float(raw)
    else:
        text = ''.join(str(raw).split()).replace(',', '.')
        percent = text.endswith('%')
        try:
            value = float(text.removesuffix('%')) / (100 if percent else 1)
        except ValueError:
            raise ValueError('введите число') from None
        if percent and field.unit != '%':
            raise ValueError('процент для этого показателя не применяется')
    if not isfinite(value) or value < 0:
        raise ValueError('введите неотрицательное число')
    if field.unit == '%' and value > 1:
        raise ValueError('введите долю со знаком %, например 6%')
    if field.unit in ('шт', 'м/м', 'компл.') and not value.is_integer():
        raise ValueError('введите целое количество')
    return value or None


def read_updates(wb, fields):
    updates, warnings = [], []
    block_order = {}
    for name in SHEETS:
        if name not in wb or wb[name]['A1'].value != MARKER:
            warnings.append(f'Лист «{name}» отсутствует или изменён, его данные пропущены.')
            continue
        ws = wb[name]
        catalog = {field.key: field for field in fields[name]}
        columns = {
            cell.column: str(cell.value).strip()
            for cell in ws[HEADER][FIRST_BLOCK - 1:] if cell.value is not None and str(cell.value).strip()
        }
        for block in columns.values():
            block_order.setdefault(block, len(block_order))
        duplicates = {block for block, count in Counter(columns.values()).items() if count > 1}
        if duplicates:
            warnings.append(f'«{name}»: повторяются названия УБ ({", ".join(sorted(duplicates))}). Эти столбцы пропущены.')
        seen = set()
        for row in ws.iter_rows(min_row=HEADER + 1):
            marker = row[0].value
            if marker in ('mp-geometry', 'ep-elements'):
                updates.extend(_repeatable_updates(name, row, warnings))
                continue
            field = catalog.get(marker)
            if field is None:
                continue
            for cell in row[FIRST_BLOCK - 1:]:
                address = f'«{name}», {cell.coordinate}'
                try:
                    value = _input_value(cell.value, field)
                except ValueError as error:
                    warnings.append(f'{address}: {error}.')
                    continue
                if value is None:
                    continue
                block = columns.get(cell.column)
                if not block:
                    warnings.append(f'{address}: у столбца нет названия блока.')
                    continue
                if block in duplicates:
                    continue
                key = (block, field.key)
                if key in seen:
                    warnings.append(f'{address}: показатель уже заполнен для «{block}». Взято первое значение.')
                    continue
                seen.add(key)
                updates.append((address, block, field, value))
    updates.sort(key=lambda update: block_order.get(update[1], len(block_order)))
    return updates, warnings


def _repeatable_updates(sheet, row, warnings):
    marker = row[0].value
    if not any(cell.value not in (None, '') for cell in row[1:9]):
        return
    block = str(row[1].value or '').strip()
    label = str(row[2].value or '').strip()
    address = f'«{sheet}», строка {row[0].row}'
    if not block or not label:
        warnings.append(f'{address}: укажите УБ и название секции или элемента.')
        return
    if marker == 'mp-geometry':
        specs = [
            (cell, Field(column, 'МП', column, unit, 'mp_sections', column,
                         label, options=options, applies=(TYPE_RESIDENTIAL,)))
            for cell, column, unit, options in zip(
                row[3:9], SECTION_COLS[1:], ('шт', 'м²', 'шт', 'м', 'м²', ''),
                ((), (), (), (), (), ('Да', 'Нет')),
            )
        ]
    else:
        if label not in DETAIL_SPEC:
            warnings.append(f'{address}: выберите элемент из списка.')
            return
        description = str(row[3].value or '').strip()
        column = 'Объем' if sheet == 'Объёмы' else RATE
        specs = [(row[4], Field(description, label, label, DETAIL_SPEC[label][1] if column == 'Объем' else '',
                                'ep_details', column, label, applies=(TYPE_RESIDENTIAL,)))]
    for cell, field in specs:
        try:
            value = _input_value(cell.value, field)
        except ValueError as error:
            warnings.append(f'«{sheet}», {cell.coordinate}: {error}.')
            continue
        if value is not None:
            yield f'«{sheet}», {cell.coordinate}', block, field, value


def merge_updates(state, updates, warnings, default_block, catalogs, cascade_defaults):
    """Build changes off-session. Leave existing inputs, project fields and widget state intact."""
    types = {name: value for _, name, field, value in updates if field.column == BLOCK_TYPE}
    blocks = state.get('blocks_df', default_block).copy(deep=True).reset_index(drop=True)
    changed, count = {}, 0
    rows = {name: index for index, name in blocks[BLOCK_NAME].items()}
    block_types = blocks.set_index(BLOCK_NAME)[BLOCK_TYPE].to_dict()

    def store(key):
        if key not in changed:
            changed[key] = deepcopy(state.get(key, {}))
        return changed[key]

    seen_rows = set()
    for address, name, field, value in updates:
        if field.column == BLOCK_TYPE and name in block_types:
            if value != block_types[name]:
                warnings.append(f'{address}: тип существующего блока «{name}» меняется на вкладке «Блоки».')
            continue
        if field.store in ('mp_sections', 'ep_details'):
            identity = (name, field.store, field.code, field.key, field.column)
            if identity in seen_rows:
                warnings.append(f'{address}: это значение уже задано. Взято первое.')
                continue
            seen_rows.add(identity)
        btype = block_types.get(name, types.get(name, TYPE_RESIDENTIAL))
        if btype not in field.applies:
            warnings.append(f'{address}: показатель не применяется для блока «{name}» этого типа.')
            continue
        if name not in rows:
            rows[name] = len(blocks)
            blocks.loc[len(blocks)] = {**default_block.iloc[0].to_dict(), BLOCK_NAME: name, BLOCK_TYPE: btype}
            changed['blocks_df'] = blocks
            for key, catalog in catalogs.items():
                if btype == TYPE_RESIDENTIAL or key not in ("block_rates", "block_mp_rates"):
                    store(key)[name] = catalog.copy(deep=True)
            for key, fields in cascade_defaults.items():
                store(key)[name] = dict.fromkeys(fields, 0.0)
        if field.store == 'blocks_df':
            blocks.at[rows[name], field.column] = value
            changed['blocks_df'] = blocks
        elif field.store == 'mp_sections':
            sections = store('mp_sections')
            frame = sections.setdefault(name, pd.DataFrame(columns=SECTION_COLS))
            match = frame[SECTION_COLS[0]].eq(field.code)
            if match.sum() > 1:
                warnings.append(f'{address}: несколько секций с таким названием. Уточните название в модуле.')
                continue
            if not match.any():
                frame = pd.concat([frame, pd.DataFrame([{SECTION_COLS[0]: field.code}])], ignore_index=True)
                sections[name] = frame
                match = frame[SECTION_COLS[0]].eq(field.code)
            frame.loc[match, field.column] = int(value == 'Да') if field.options else value
        elif field.store == 'ep_details':
            details = store('ep_details')
            frame = details.setdefault(name, pd.DataFrame(columns=DETAIL_COLS))
            match = frame['Категория'].eq(field.code) & frame['Тип / материал / размеры'].fillna('').eq(field.key)
            if match.sum() > 1:
                warnings.append(f'{address}: несколько одноимённых элементов. Уточните тип в модуле.')
                continue
            if not match.any():
                frame = pd.concat([frame, pd.DataFrame([{
                    'Категория': field.code, 'Тип / материал / размеры': field.key,
                }])], ignore_index=True)
                details[name] = frame
                match = frame['Категория'].eq(field.code) & frame['Тип / материал / размеры'].fillna('').eq(field.key)
            frame.loc[match, field.column] = value
            calculated = match & frame['Статус'].fillna('').isin(('', 'Нет данных'))
            calculated &= pd.to_numeric(frame['Объем'], errors='coerce').gt(0)
            calculated &= pd.to_numeric(frame[RATE], errors='coerce').gt(0)
            frame.loc[calculated, 'Статус'] = 'Рассчитан'
        elif field.code:
            frame = store(field.store).setdefault(name, catalogs[field.store].copy(deep=True))
            if field.code not in set(frame['Код']):
                reference = catalogs[field.store]
                frame = pd.concat([frame, reference[reference['Код'] == field.code]], ignore_index=True)
                store(field.store)[name] = frame
            frame.loc[frame['Код'] == field.code, field.column] = value
        elif field.store in ('block_elevator_stops', 'thermal_kind', 'facade_modes'):
            store(field.store)[name] = value
        else:
            store(field.store).setdefault(name, {})[field.column] = value
        count += 1
    return changed, count
