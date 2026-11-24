# -*- coding: utf-8 -*-
"""Сборка Excel-решения задачи 2.4 (тема 2 «Модели и методы многоцелевой
оптимизации») — листы 2.4а, 2.4б, 2.4в, 2.4г + лист с исходными данными.

Модель каждого листа приведена к виду, который получается после работы
«Поиска решения» (Excel Solver): значения переменных, формулы СУММПРОИЗВ для
критериев и левых частей ограничений. Числа соответствуют решению, полученному
аналогичным ЛП-расчётом (см. solve_mmolo.py).
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mmolo_models import models, RES, CAP, C1, C2, C3, fmt  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   'outputs', 'Эконометрика УСР 2.4 (Пась А.М.).xlsx')

FONT = Font(name='Calibri', size=11)
THIN = Side(style='thin', color='B7B7B7')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def setup_sheet(ws, nrows, ncols):
    ws.sheet_view.showGridLines = True
    ws.column_dimensions['A'].width = 11
    for c in range(2, 2 + ncols):
        ws.column_dimensions[get_column_letter(c)].width = 12.5
    ws.column_dimensions[get_column_letter(2 + ncols)].width = 3
    ws.column_dimensions[get_column_letter(3 + ncols)].width = 12.5
    ws.column_dimensions[get_column_letter(4 + ncols)].width = 4
    ws.column_dimensions[get_column_letter(5 + ncols)].width = 12.5
    for r in range(1, nrows + 2):
        ws.row_dimensions[r].height = 16.5


def cell(ws, r, c, v, *, align='right', border=True, bold=False):
    cl = ws.cell(row=r, column=c, value=v)
    cl.font = Font(name='Calibri', size=11, bold=bold)
    cl.alignment = Alignment(horizontal=align, vertical='center')
    cl.number_format = 'General'
    if border:
        cl.border = BORDER
    return cl


def main():
    M = models()
    wb = Workbook()

    # --- лист исходных данных -------------------------------------------------
    ws = wb.active
    ws.title = 'Данные'
    ws.sheet_view.showGridLines = True
    for col, w in zip('ABCDEFG', (16, 11, 11, 11, 14, 4, 26)):
        ws.column_dimensions[col].width = w
    cell(ws, 1, 1, 'Таблица 2.3 — расход ресурсов на единицу продукции', align='left', border=False)
    cell(ws, 2, 1, 'Ресурсы', align='left')
    for j, p in enumerate(('1П', '2П', '3П')):
        cell(ws, 2, 2 + j, p)
    cell(ws, 2, 5, 'Наличие ресурсов')
    for i, lab in enumerate(('Ресурс 1', 'Ресурс 2', 'Ресурс 3')):
        cell(ws, 3 + i, 1, lab, align='left')
        for j in range(3):
            cell(ws, 3 + i, 2 + j, RES[i][j])
        cell(ws, 3 + i, 5, CAP[i])
    cell(ws, 7, 1, 'прибыль', align='left')
    for j, v in enumerate(C1):
        cell(ws, 7, 2 + j, v)
    cell(ws, 8, 1, 'оптовая цена', align='left')
    for j, v in enumerate(C2):
        cell(ws, 8, 2 + j, v)
    cell(ws, 9, 1, 'трудоемкость', align='left')
    for j, v in enumerate(C3):
        cell(ws, 9, 2 + j, v)
    cell(ws, 11, 1, 'Ограничения: x3 ≥ 30 (п. а, б, в), x3 ≥ 40 (п. г), x1 ≥ 20 (п. в)',
         align='left', border=False)
    cell(ws, 12, 1, 'Критерии: f1 → max, f2 → max, f3 → min', align='left', border=False)

    # --- листы методов: на каждом листе все состояния задачи, снизу вверх ------
    sheets = {
        '2.4а': ['a1', 'a2', 'a3'],
        '2.4б': ['b1'],
        '2.4в': ['c1', 'c2', 'c3', 'c4'],
        '2.4г': ['d0', 'd1', 'd2', 'd3'],
    }
    for sname, keys in sheets.items():
        ws = wb.create_sheet(sname)
        r = 1
        for k in keys:
            m = M[k]
            cell(ws, r, 1, m['title'] + '  (модель после «Поиска решения»)', align='left', border=False, bold=True)
            r += 1
            start = r
            nxt, _, _ = write_model_rows(ws, m, start, ncol=4)
            r = nxt + 2
        setup_sheet(ws, r, 4)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print('saved', OUT)


def write_model_rows(ws, m, start, ncol=4):
    """Пишет блок модели; столбцы всегда B..E (4 слота), LHS — G, знак — H, RHS — I."""
    nv = max(len(m['vars']), ncol)
    pad = lambda seq: list(seq) + [0] * (nv - len(seq))
    for i, name in enumerate(m['vars']):
        cell(ws, start, 2 + i, name, align='left', border=False)
    r = start + 1
    cell(ws, r, 1, 'значение')
    for i, v in enumerate(m['values']):
        cell(ws, r, 2 + i, round(float(v), 6))
    vrow = r
    r += 1
    for o in m['objs']:
        cell(ws, r, 1, o['label'])
        for j, co in enumerate(pad(o['coef'])):
            cell(ws, r, 2 + j, co if co != 0 or j < len(o['coef']) else None)
        cell(ws, r, 2 + nv + 1,
             f"=СУММПРОИЗВ(B{r}:{get_column_letter(1 + nv)}{r},B${vrow}:{get_column_letter(1 + nv)}${vrow})")
        r += 1
    for i, c in enumerate(m['cons'], start=1):
        cell(ws, r, 1, i)
        for j, co in enumerate(pad(c['coef'])):
            cell(ws, r, 2 + j, co if co != 0 else None)
        cell(ws, r, 2 + nv + 1,
             f"=СУММПРОИЗВ(B{r}:{get_column_letter(1 + nv)}{r},B${vrow}:{get_column_letter(1 + nv)}${vrow})")
        cell(ws, r, 2 + nv + 2, {'<=': '≤', '>=': '≥', '=': '='}.get(c['op'], c['op']), align='left')
        cell(ws, r, 2 + nv + 3, c['rhs'])
        r += 1
    return r, vrow, []


if __name__ == '__main__':
    main()
