# -*- coding: utf-8 -*-
"""Скриншоты листов Excel-решения темы 3 (задания 3.4, 3.6, 3.12).

Кадр «как в Excel»: сетка ячеек выбранного диапазона, значения формул считаются
xlcalc.SheetCalc, форматирование (в т.ч. проценты) берётся из number_format.
Один и тот же рендер даёт и картинки по ходу РЕШЕНИЯ (узкие диапазоны), и
полные листы для ПРИЛОЖЕНИЯ А.
"""
import os
import sys

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
from xlcalc import SheetCalc  # noqa: E402

XLSX = os.path.join(ROOT, 'outputs', 'Эконометрика УСР 3 (Пась А.М.) — задания 3.4, 3.6, 3.12.xlsx')
OUT = os.path.join(ROOT, 'report_assets')
DPI = 300          # плотность «скриншота»; в отчёте картинки вставляются 1:1
OPS = {'<=': '≤', '>=': '≥'}


def cell_text(sc, ws, r, c):
    """Текст ячейки так, как его показывает Excel (с учётом формата)."""
    cell = ws.cell(row=r, column=c)
    v = sc.value(r, c)
    if v is None or v == '':
        return ''
    nf = cell.number_format or ''
    if '%' in nf and isinstance(v, (int, float)):
        seg = nf.split('.')[1].split('%')[0] if '.' in nf else ''
        dec = seg.count('0') + seg.count('#')
        dec = max(dec, 0)
        return f'{100 * float(v):.{dec}f}%'.replace('.', ',')
    if isinstance(v, float):
        return sc.text(r, c, nd=4)
    return OPS.get(str(v).strip(), str(v).strip())


def render(ws, r1, r2, c1, c2, fname, widths_map=None, zoom=1.0):
    sc = SheetCalc(ws)
    ncols, nrow = c2 - c1 + 1, r2 - r1 + 1
    widths_map = widths_map or {}
    widths = []
    for c in range(c1, c2 + 1):
        w = widths_map.get(get_column_letter(c)) or 9.5
        widths.append(max(w, 3.2))
    total = sum(widths)
    xs = [0.0]
    for w in widths:
        xs.append(xs[-1] + w / total * 100)
    fig_w = total * 7.2 / 72.0 * zoom
    fig_h = (nrow * 16.0 + 4) / 72.0 * zoom
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=DPI)
    ax.set_xlim(0, total * 1.0)
    ax.set_ylim(0, nrow + 0.05)
    ax.axis('off')
    fig.subplots_adjust(left=0.004, right=0.996, top=0.996, bottom=0.004)
    xs = [0.0]
    for w in widths:
        xs.append(xs[-1] + w)
    for i, r in enumerate(range(r1, r2 + 1)):
        ytop = nrow - i
        for j, c in enumerate(range(c1, c2 + 1)):
            ax.add_patch(Rectangle((xs[j], ytop - 0.95), xs[j + 1] - xs[j], 0.95,
                                   facecolor='white', edgecolor='#d0d0d0', lw=0.45))
            txt = cell_text(sc, ws, r, c)
            if not txt:
                continue
            x0, x1 = xs[j], xs[j + 1]
            try:
                float(txt.replace(',', '').replace('%', ''))
                is_num = True
            except ValueError:
                is_num = False
            if txt in ('≤', '≥', '='):
                ax.text(x0 + 0.35, ytop - 0.47, txt, ha='left', va='center', fontsize=7.6)
            elif is_num:
                ax.text(x1 - 0.45, ytop - 0.47, txt, ha='right', va='center', fontsize=7.4)
            elif len(txt) <= 5:
                ax.text((x0 + x1) / 2, ytop - 0.47, txt, ha='center', va='center', fontsize=7.2)
            else:
                # текстовая подпись: если не влезает в столбец, Excel льёт её
                # вправо по пустым ячейкам — рисуем так же, затирая сетку
                span = 1
                need = len(txt) * 1.15 + 0.8
                while (xs[j + span] - x0 < need and j + span < ncols
                       and not cell_text(sc, ws, r, c + span)):
                    ax.add_patch(Rectangle((xs[j + span], ytop - 0.95),
                                           xs[j + span + 1] - xs[j + span], 0.95,
                                           facecolor='white', edgecolor='none', lw=0))
                    span += 1
                ax.text(x0 + 0.4, ytop - 0.47, txt, ha='left', va='center',
                        fontsize=7.0, clip_on=True,
                        bbox=dict(facecolor='white', edgecolor='none', pad=0.2))
                ax.set_xlim(xs[0], xs[ncols])
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, fname)
    fig.savefig(path, dpi=DPI, facecolor='white', bbox_inches='tight', pad_inches=0.02)
    plt.close(fig)
    return path


# (лист, r1, r2, c1, c2, имя файла) — по ходу РЕШЕНИЯ
INTEXT = [
    ('3.4', 1, 3, 1, 2, 'inv_34_data.png'),
    ('3.4', 5, 14, 1, 4, 'inv_34_base.png'),
    ('3.4', 16, 26, 1, 7, 'inv_34_infl.png'),
    ('3.6', 1, 3, 1, 2, 'inv_36_data.png'),
    ('3.6', 5, 18, 1, 6, 'inv_36_base.png'),
    ('3.6', 20, 26, 1, 8, 'inv_36_infl.png'),
    ('3.12', 1, 9, 1, 10, 'inv_312_cf.png'),
    ('3.12', 11, 16, 1, 10, 'inv_312_model.png'),
    ('3.12', 18, 29, 1, 5, 'inv_312_portfeli.png'),
]
# полные листы — в ПРИЛОЖЕНИЕ А
FULL = [('3.4', 7, 'inv_full_34.png'), ('3.6', 8, 'inv_full_36.png'),
        ('3.12', 10, 'inv_full_312.png')]


def main():
    wb = load_workbook(XLSX)
    made = []
    for name, r1, r2, c1, c2, fname in INTEXT:
        ws = wb[name]
        made.append(render(ws, r1, r2, c1, c2, fname))
    for name, cols, fname in FULL:
        ws = wb[name]
        widths_map = {k: (d.width or 9.5) for k, d in ws.column_dimensions.items()}
        last = max(cell.row for row in ws.iter_rows() for cell in row if cell.value not in (None, ''))
        made.append(render(ws, 1, last, 1, cols, fname, widths_map))
    for p in made:
        print('saved', os.path.basename(p))


if __name__ == '__main__':
    main()
