# -*- coding: utf-8 -*-
"""Excel-книга решения заданий 3.4, 3.6, 3.12 (тема 3).

На листах — «живые» формулы: дисконтирование потоков (1/(1+r)^t), PV, NPV, PI,
перебор вариантов портфеля через СУММПРОИЗВ. Значения совпадают с расчётом
solve_invest.py; после открытия в Excel всё пересчитывается автоматически.
"""
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, 'outputs', 'Эконометрика УСР 3 (Пась А.М.) — задания 3.4, 3.6, 3.12.xlsx')
DATA = json.load(open('/tmp/invest_results.json'))


def put(ws, r, c, v, *, fmt=None, align='right', bold=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = Font(name='Calibri', size=11, bold=bold)
    cell.alignment = Alignment(horizontal=align, vertical='center')
    if fmt:
        cell.number_format = fmt
    return cell


def label(ws, r, text, span=3, bold=True, value=None, fmt=None):
    """Подпись, растянутая на несколько столбцов (объединённая ячейка, как в Excel)."""
    put(ws, r, 1, text, align='left', bold=bold)
    for c in range(2, span + 1):
        put(ws, r, c, None, align='left')
    if span > 1:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    cell = ws.cell(row=r, column=1)
    if value is not None:                      # «подпись = значение» в одной ячейке
        cell.value = None
        cell = put(ws, r, 1, f'{text} = {fmt_value(value, fmt)}',
                   align='left', bold=bold)
    return cell


def fmt_value(v, fmt):
    """Значение в подписи так, как его показал бы Excel."""
    if fmt and '%' in fmt:
        dec = fmt.split('.')[1].split('%')[0].count('0')
        return f'{100 * v:.{dec}f}%'.replace('.', ',')
    return f'{v:.4f}'.replace('.', ',')


def head(ws, r, labels, c0=1):
    for j, h in enumerate(labels):
        if h:
            put(ws, r, c0 + j, h, align='center', bold=True)


def sheet_34(wb, d):
    ws = wb.active
    ws.title = '3.4'
    widths = [26, 14, 14, 15, 15, 14, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    put(ws, 1, 1, 'Ставка дисконтирования r0 (реальная)', align='left')
    put(ws, 1, 2, d['r'], fmt='0%')
    put(ws, 2, 1, 'Начальные инвестиции I₀, ден. ед.', align='left')
    put(ws, 2, 2, d['I0'])
    put(ws, 3, 1, 'Срок проекта, лет', align='left')
    put(ws, 3, 2, len(d['C']))

    put(ws, 5, 1, 'Базовый расчёт (без инфляции)', align='left', bold=True)
    head(ws, 6, ['Год t', 'Поток C(t)', 'множитель', 'PV потока'])
    for k, cf in enumerate(d['C'], start=1):
        r = 6 + k
        put(ws, r, 1, k, align='center')
        put(ws, r, 2, cf)
        put(ws, r, 3, f'=1/(1+$B$1)^A{r}', fmt='0.0000')
        put(ws, r, 4, f'=B{r}*C{r}', fmt='0.0000')
    put(ws, 9, 1, 'PV проекта', align='left')
    put(ws, 9, 4, '=SUM(D7:D8)', fmt='0.0000')
    put(ws, 10, 1, 'NPV = PV − I₀', align='left')
    put(ws, 10, 4, '=D9-B2', fmt='0.0000')
    put(ws, 11, 1, 'PI = PV / I₀', align='left')
    put(ws, 11, 4, '=D9/B2', fmt='0.0000')
    put(ws, 12, 1, 'IRR', align='left')
    put(ws, 12, 4, d['base']['irr'], fmt='0.0000%')
    put(ws, 13, 1, 'PP, лет', align='left')
    put(ws, 13, 4, d['base']['pp'])
    put(ws, 14, 1, 'DPP, лет', align='left')
    put(ws, 14, 4, d['base']['dpp'])

    b = 16
    put(ws, b, 1, 'Влияние уровня инфляции', align='left', bold=True)
    head(ws, b + 1, ['вариант', 'i', 'r = r0 + i + r0·i', 'PV', 'NPV', 'PI', 'вывод'])
    for k, row in enumerate(d['rows']):
        r = b + 2 + k
        put(ws, r, 1, 'вариант %d' % (k + 1), align='left')
        put(ws, r, 2, row['i'], fmt='0.0000%')
        put(ws, r, 3, '=$B$1+B%d+$B$1*B%d' % (r, r), fmt='0.0000%')
        put(ws, r, 4, '=%s/(1+C%d)+%s/(1+C%d)^2' % (d['C'][0], r, d['C'][1], r), fmt='0.0000')
        put(ws, r, 5, '=D%d-$B$2' % r, fmt='0.0000')
        put(ws, r, 6, '=D%d/$B$2' % r, fmt='0.0000')
        put(ws, r, 7, 'эффективен' if row['npv'] > 0.5 else
            ('граница' if abs(row['npv']) < 0.5 else 'неэффективен'), align='left')
    last = b + 1 + len(d['rows'])
    label(ws, last + 1, 'Критический уровень инфляции i', span=4,
          value=d['i_max'], fmt='0.0000%')
    return ws


def sheet_36(wb, d):
    ws = wb.create_sheet('3.6')
    for i, w in enumerate([26, 13, 13, 13, 13, 13, 13, 26], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    put(ws, 1, 1, 'Норма дисконтирования r0', align='left')
    put(ws, 1, 2, d['r'], fmt='0%')
    put(ws, 2, 1, 'Начальные инвестиции I₀, млн ден. ед.', align='left')
    put(ws, 2, 2, d['I0'])
    put(ws, 3, 1, 'Срок проекта, лет', align='left')
    put(ws, 3, 2, len(d['C']))

    put(ws, 5, 1, 'Расчёт без учёта инфляции', align='left', bold=True)
    head(ws, 6, ['Год t', 'Поток C(t)', 'множитель', 'PV потока', 'наращ. FV'])
    for k, cf in enumerate(d['C'], start=1):
        r = 6 + k
        put(ws, r, 1, k, align='center')
        put(ws, r, 2, cf)
        put(ws, r, 3, f'=1/(1+$B$1)^A{r}'.replace('$B$1', '$B$1'), fmt='0.0000')
        put(ws, r, 4, f'=B{r}*C{r}', fmt='0.0000')
        put(ws, r, 5, f'=B{r}*(1+$B$1)^(5-A{r})', fmt='0.0000')
    put(ws, 12, 1, 'PV проекта', align='left')
    put(ws, 12, 4, '=SUM(D7:D11)', fmt='0.0000')
    put(ws, 12, 5, '=SUM(E7:E11)', fmt='0.0000')
    put(ws, 13, 1, 'NPV = PV − I₀', align='left')
    put(ws, 13, 4, '=D12-B2', fmt='0.0000')
    put(ws, 14, 1, 'PI = PV / I₀', align='left')
    put(ws, 14, 4, '=D12/B2', fmt='0.0000')
    put(ws, 15, 1, 'IRR', align='left')
    put(ws, 15, 4, d['base']['irr'], fmt='0.0000%')
    put(ws, 16, 1, 'MIRR', align='left')
    put(ws, 16, 4, '=E12/B2', fmt='0.0000')
    put(ws, 16, 4, d['base']['mirr'], fmt='0.0000%')
    put(ws, 17, 1, 'PP, лет (простой срок окупаемости)', align='left')
    put(ws, 17, 4, d['base']['pp'])
    put(ws, 18, 1, 'DPP, лет (дисконтированный срок)', align='left')
    put(ws, 18, 4, d['base']['dpp'])

    b = 20
    put(ws, b, 1, 'Учёт инфляции: r = r0 + i + r0·i', align='left', bold=True)
    head(ws, b + 1, ['сценарий', 'i', 'r номинальная', 'PV', 'NPV', 'PI', 'IRR', 'вывод'])
    for k, row in enumerate(d['rows']):
        r = b + 2 + k
        lab = 'критическое i' if row['critical'] else 'i = %g %%' % (100 * row['i'])
        put(ws, r, 1, lab, align='left')
        put(ws, r, 2, row['i'], fmt='0.00%')
        put(ws, r, 3, row['d'], fmt='0.0000%')
        put(ws, r, 4, row['pv'], fmt='0.0000')
        put(ws, r, 5, '=D%d-$B$2' % r, fmt='0.0000')
        put(ws, r, 6, '=D%d/$B$2' % r, fmt='0.0000')
        put(ws, r, 7, d['base']['irr'], fmt='0.0000%')
        put(ws, r, 8, 'инвестировать целесообразно' if row['npv'] > 0.5 else
            ('на границе эффективности' if row['critical']
             else 'инвестировать нецелесообразно'), align='left')
    last = b + 1 + len(d['rows'])
    label(ws, last + 1, 'Критический уровень инфляции i (NPV = 0)', span=3,
          value=d['i_crit'], fmt='0.0000%')
    return ws


def sheet_312(wb, d):
    ws = wb.create_sheet('3.12')
    for i, w in enumerate([10, 11, 11, 11, 11, 11, 13, 11, 11, 9, 13, 13, 13], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    put(ws, 1, 1, 'Денежные потоки проектов, млн руб.  (ставка r = 15 %)', align='left',
        bold=True)
    head(ws, 2, ['Проект', 'год 0', 'год 1', 'год 2', 'год 3', 'год 4', 'NPV',
                 'Инвестиции', 'выпуск, ед.', 'x'])
    for k, p in enumerate(d['projects']):
        r = 3 + k
        cf = {int(t): v for t, v in p['cf'].items()}
        put(ws, r, 1, p['name'], align='center')
        for t in range(5):
            v = cf.get(t, None)
            put(ws, r, 2 + t, v)
        # NPV = ΣCF_t·1/(1+r)^t — живой SUMPRODUCT по строкам листа
        put(ws, r, 7, '=SUMPRODUCT(B%d:F%d,$B$9:$F$9)' % (r, r), fmt='0.0000')
        put(ws, r, 8, p['invest'])
        put(ws, r, 9, p['units'])
        put(ws, r, 10, 1 if p['name'] in d['best']['names'] else 0, fmt='0')
    put(ws, 9, 1, 'множитель', align='left')
    for t in range(5):
        put(ws, 9, 2 + t, '=1/(1+0.15)^%d' % t, fmt='0.0000')
    ws.cell(row=9, column=1).value = '1/(1+r)^t'
    put(ws, 11, 1, 'Целевая функция: Σ NPV·x → max', align='left', bold=True)
    put(ws, 11, 7, '=SUMPRODUCT(G3:G8,J3:J8)', fmt='0.0000')
    put(ws, 12, 1, 'Ограничения', align='left', bold=True)
    cons = [
        ('бюджет: Σ I0·x ≤ 100', '=SUMPRODUCT(H3:H8,J3:J8)', '≤', 100),
        ('спрос: Σ выпуск·x ≤ 550', '=SUMPRODUCT(I3:I8,J3:J8)', '≤', 550),
        ('исключение: x1 + x2 ≤ 1', '=J3+J4', '≤', 1),
        ('исключение: x3 + x5 ≤ 1', '=J5+J7', '≤', 1),
    ]
    for k, (lab, f, op, rhs) in enumerate(cons):
        r = 13 + k
        put(ws, r, 1, lab, align='left')
        put(ws, r, 7, f, fmt='0.0000')
        put(ws, r, 9, op, align='left')
        put(ws, r, 10, rhs)
    put(ws, 18, 1, 'Перебор 64 наборов: допустимые решения', align='left', bold=True)
    head(ws, 19, ['портфель', 'Σ NPV', 'Σ I0', 'выпуск', 'статус'])
    for k, z in enumerate(d['all_feas'][:10]):
        r = 20 + k
        put(ws, r, 1, '{' + '; '.join(z['names']) + '}', align='center')
        put(ws, r, 2, z['npv'], fmt='0.0000')
        put(ws, r, 3, z['spend'])
        put(ws, r, 4, z['units'])
        put(ws, r, 5, 'оптимум' if k == 0 else 'допустим', align='center')
    return ws


def main():
    wb = Workbook()
    sheet_34(wb, DATA['t34'])
    sheet_36(wb, DATA['t36'])
    sheet_312(wb, DATA['t312'])
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print('saved', OUT)


if __name__ == '__main__':
    main()
