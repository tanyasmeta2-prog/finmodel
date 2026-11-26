"""Presentation and scenario breakdowns for the Excel report."""
from numbers import Real

from openpyxl.cell.cell import MergedCell
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrData, StrRef, StrVal
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NUMBER = '#,##0.00'
PERCENT = '0.00%'
TITLE_COLOR = '2C6331'
TEXT_COLOR = '25332B'
MUTED_COLOR = '59675E'
GREEN = '41AA37'
BORDO = '9B3654'
PALE_GRAY = 'F7F9F7'
HEADER_COLOR = 'EDF2ED'
TOTAL_COLOR = 'E4F0E2'
DETAIL_HEADER = 5

GRID = Border(bottom=Side(style='hair', color='DCE5DC'))
HEADER_LINE = Border(bottom=Side(style='thin', color='CCD9CD'))
TOTAL_LINE = Border(top=Side(style='thin', color='B9CEB7'))


def title(cell, *, primary=True):
    cell.font = Font(name='Arial', size=16 if primary else 11, bold=True, color=TITLE_COLOR)
    cell.alignment = Alignment(vertical='center')


def report_heading(ws, context, headings=None):
    ws['A1'] = ws.title
    title(ws['A1'])
    ws['A2'] = context
    ws['A2'].font = Font(name='Arial', size=10, color=MUTED_COLOR)
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18
    ws.sheet_view.showGridLines = False
    if headings is not None:
        for col, label in enumerate(headings, 1):
            ws.cell(DETAIL_HEADER, col, label)
        ws.freeze_panes = 'C6'


def report_table(ws, header, last, first_col, last_col, *, totals=(), subheaders=(), gutters=()):
    """Apply the same header, data and subtotal styles without moving formula cells."""
    for row in ws.iter_rows(min_row=header, max_row=last, min_col=first_col, max_col=last_col):
        row_no = row[0].row
        if not any(cell.value is not None and cell.value != '' for cell in row):
            continue
        is_header = row_no == header or row_no in subheaders
        is_total = row_no in totals
        color = HEADER_COLOR if is_header else TOTAL_COLOR if is_total else (
            PALE_GRAY if (row_no - header) % 2 == 0 else 'FFFFFF'
        )
        for cell in row:
            if cell.column in gutters or isinstance(cell, MergedCell):
                continue
            is_numeric = (cell.data_type == 'f' and not is_header) or (
                isinstance(cell.value, Real) and not isinstance(cell.value, bool) and cell.number_format != '@'
            )
            if isinstance(cell.value, str) and cell.value.startswith('🔒 '):
                cell.value = cell.value.removeprefix('🔒 ')
            cell.font = Font(name='Arial', size=10, bold=is_header or is_total, color=TEXT_COLOR)
            cell.fill = PatternFill('solid', fgColor=color)
            cell.border = HEADER_LINE if is_header else TOTAL_LINE if is_total else GRID
            cell.alignment = Alignment(
                horizontal='right' if is_numeric and not is_header else 'left',
                vertical='center', wrap_text=not is_numeric or is_header,
            )
            if is_numeric and not is_header:
                fmt = cell.number_format
                if '%' in fmt:
                    cell.number_format = PERCENT
                elif ',,' in fmt:
                    cell.number_format = '#,##0.00,, "млн руб."'
                else:
                    cell.number_format = NUMBER
        height = 32 if is_header else 20
        if ws.row_dimensions[row_no].height is None or ws.row_dimensions[row_no].height < height:
            ws.row_dimensions[row_no].height = height


def negative_values(ws, area):
    ws.conditional_formatting.add(
        area, CellIsRule(operator='lessThan', formula=['0'],
                         font=Font(color=BORDO, bold=True), fill=PatternFill('solid', fgColor='F8EDF1'))
    )


def add_sales(ws, blocks, first_helper_row, columns, helper):
    heading = ws.max_row + 2
    ws.cell(heading, 1, 'Цены продаж и выручка по продуктам')
    title(ws.cell(heading, 1), primary=False)
    ws.merge_cells(start_row=heading, start_column=1, end_row=heading, end_column=6)
    header = heading + 1
    for col, label in enumerate(['Блок', 'Продукт', 'Ед.', 'Объём', 'Цена, руб./ед.', 'Выручка, млн руб.'], 1):
        ws.cell(header, col, label)
    data_row = header + 1
    common = [
        ('Коммерция', 'м²', 'S коммерции 1 эт., м2', 'Цена коммерции, руб/м2'),
        ('Кладовые', 'м²', 'S кладовых, м2', 'Цена кладовых, руб/м2'),
    ]
    residential = ('Квартиры с коэффициентами', 'м²', 'S квартир, м2', 'Цена жилья, руб/м2')
    underground = ('Подземные машино-места', 'м/м', 'Подземный паркинг, м/м', 'Цена подземного м/м, руб')
    above = ('Машино-места', 'м/м', 'Наземный/Многоуровневый паркинг, м/м', 'Цена наземного м/м, руб')
    for i, (_, block) in enumerate(blocks.iterrows()):
        row = first_helper_row + i
        is_residential = block['Тип блока'] == 'Жилой блок'
        products = [residential, *common, underground] if is_residential else [*common, above]
        for product, unit, quantity, price in products:
            ws.append([
                block['Название блока'], product, unit,
                f"='{helper}'!{columns[quantity]}{row}",
                f"='{helper}'!{columns[price]}{row}*'{helper}'!$B$4",
                f'=D{data_row}*E{data_row}/1000000',
            ])
            data_row += 1
    ws.append(['Итого', None, None, None, None, f'=SUM(F{header + 1}:F{data_row - 1})' if data_row > header + 1 else 0])
    report_table(ws, header, data_row, 1, 6, totals={data_row})
    ws.column_dimensions['B'].width = 43


def _column_chart(ws, title_text, anchor, data_columns, names_col, header, count, colors):
    chart = BarChart(type='col', grouping='clustered', overlap=0, style=10)
    chart.title = title_text
    chart.width, chart.height = 15.4, 8.4
    chart.y_axis.numFmt = NUMBER
    chart.x_axis.tickLblPos = 'low'
    chart.visible_cells_only = False
    end = header + count
    cats = Reference(ws, min_col=names_col, min_row=header + 1, max_row=end)
    labels = [StrVal(idx=i, v=str(ws.cell(header + i + 1, names_col).value)) for i in range(count)]
    for column, color in zip(data_columns, colors):
        chart.add_data(Reference(ws, min_col=column, min_row=header, max_row=end), titles_from_data=True)
        series = chart.series[-1]
        # Excel expects a string reference for named blocks and cost categories.
        series.cat = AxDataSource(strRef=StrRef(f=str(cats), strCache=StrData(ptCount=count, pt=labels)))
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.noFill = True
    if len(data_columns) == 1:
        chart.legend = None
    else:
        chart.legend.position = 'b'
    ws.add_chart(chart, anchor)


def dashboard_breakdowns(ws, blocks, columns, helper, first_row, total_row, cost_parts):
    """Visible source ranges keep the charts and tables readable after export."""
    ws.freeze_panes = 'C18'
    ws['B15'] = 'Экономика по блокам'
    title(ws['B15'], primary=False)
    header = 17
    fields = [
        ('Блок', 'Название блока', 1), ('NSA, м²', 'NSA, м2', 1),
        ('Выручка, млн руб.', 'Выручка', 1000000), ('Прямые, млн руб.', 'Прямые затраты', 1000000),
        ('Косвенные, млн руб.', 'Аллоцированные затраты', 1000000),
        ('Полные, млн руб.', 'Полные затраты', 1000000),
        ('Валовая прибыль, млн руб.', 'Валовая прибыль', 1000000),
        ('Валовая рентабельность, %', 'Валовая рентабельность', 1),
    ]
    for column, (label, _, _) in enumerate(fields, 2):
        ws.cell(header, column, label)
    ws.cell(header, 10, 'Полные / NSA, руб./м²')

    for i, name in enumerate([*blocks['Название блока'], 'Итого']):
        row = header + i + 1
        source = min(first_row + i, total_row)
        ws.cell(row, 2, name)
        for column, (_, field, divisor) in enumerate(fields[1:], 3):
            ref = f"='{helper}'!{columns[field]}{source}"
            ws.cell(row, column, ref if divisor == 1 else f'{ref}/{divisor}')
        ws.cell(row, 9).number_format = PERCENT
        if i < len(blocks) and blocks.iloc[i]['Тип блока'] == 'Жилой блок':
            ws.cell(row, 10, f'=IF(C{row}>0,G{row}*1000000/C{row},"—")')
        else:
            ws.cell(row, 10, '—')
    block_total = header + len(blocks) + 1
    report_table(ws, header, block_total, 2, 10, totals={block_total})
    negative_values(ws, f'H{header + 1}:I{block_total}')
    negative_values(ws, 'D12:E12')

    chart_top = block_total + 2
    cost_header = chart_top + 19
    ws.cell(cost_header - 1, 2, 'Структура полных затрат')
    title(ws.cell(cost_header - 1, 2), primary=False)
    ws.cell(cost_header, 2, 'Статья')
    ws.cell(cost_header, 3, 'Сумма, млн руб.')
    for row, (label, expression) in enumerate(cost_parts, cost_header + 1):
        ws.cell(row, 2, label)
        ws.cell(row, 3, f'=({expression})/1000000')
    cost_total = cost_header + len(cost_parts) + 1
    ws.cell(cost_total, 2, 'Итого')
    ws.cell(cost_total, 3, f'=SUM(C{cost_header + 1}:C{cost_total - 1})')
    report_table(ws, cost_header, cost_total, 2, 3, totals={cost_total})
    if len(blocks):
        _column_chart(ws, 'Выручка и полные затраты, млн руб.',
                      f'B{chart_top}', (4, 7), 2, header, len(blocks), (GREEN, 'A4ADA3'))
    _column_chart(ws, 'Структура полных затрат, млн руб.',
                  f'G{chart_top}', (3,), 2, cost_header, len(cost_parts), (GREEN,))
    return cost_total
