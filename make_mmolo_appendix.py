# -*- coding: utf-8 -*-
"""Скриншоты листов Excel-решения задачи 2.4 для ПРИЛОЖЕНИЯ А.

Рендерит реальные диапазоны листа книги
«outputs/Эконометрика УСР 2.4 (Пась А.М.).xlsx»: формулы СУММПРОИЗВ считаются
по значениям переменных, знаки «<=»/«>=» показываются как ≤/≥.
"""
import os
import sys

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mmolo_models import fmt  # noqa: E402

ROOT = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(ROOT, 'outputs', 'Эконометрика УСР 2 (Пась А.М.) — задание 2.4.xlsx')
OUT = os.path.join(ROOT, 'report_assets')
OPS = {'<=': '≤', '>=': '≥', '=': '='}


def _cells(ref):
    ref = ref.strip().replace('$', '')
    if ':' not in ref:
        ref = ref + ':' + ref
    c1, r1, c2, r2 = _bounds(ref)
    return [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def _bounds(ref):
    from openpyxl.utils import range_boundaries
    return range_boundaries(ref)


class Sheet:
    def __init__(self, ws, maxr, maxc):
        self.ws, self.maxr, self.maxc = ws, maxr, maxc
        self.cache = {}

    def raw(self, r, c):
        return self.ws.cell(row=r, column=c).value

    def num(self, r, c):
        key = (r, c)
        if key in self.cache:
            return self.cache[key]
        v = self.raw(r, c)
        if isinstance(v, str) and v.startswith('=СУММПРОИЗВ('):
            body = v[len('=СУММПРОИЗВ('):-1]
            a, b = body.split(',')
            # попарное произведение двух диапазонов
            ca = [self.raw(*k) for k in _cells(a)]
            cb = [self.val(*k) for k in _cells(b)]
            tot = 0.0
            for x, y in zip(ca, cb):
                xv = x if isinstance(x, (int, float)) else 0
                yv = y if isinstance(y, (int, float)) else 0
                tot += xv * yv
            self.cache[key] = tot
            return tot
        self.cache[key] = None
        return None

    def val(self, r, c):
        v = self.raw(r, c)
        if isinstance(v, str) and v.startswith('='):
            return self.num(r, c)
        return v

    def text(self, r, c):
        v = self.val(r, c)
        if v is None:
            return ''
        if isinstance(v, float):
            return fmt(v, 4).replace('.', ',')
        if isinstance(v, int):
            return str(v)
        return OPS.get(str(v), str(v))


def render(sheet, r1, r2, c1, c2, fname, title, widths_map):
    ncol, nrow = c2 - c1 + 1, r2 - r1 + 1
    show_chrome = False
    widths = [max(widths_map.get(get_column_letter(c), 9.5), 3.0) for c in range(c1, c2 + 1)]
    # узкий служебный столбец (F) и столбец знака (H)
    total = sum(widths)
    xs = [0.0]
    for w in widths:
        xs.append(xs[-1] + w / total * 100)
    fig_w = total * 7.2 / 72.0
    fig_h = (nrow * 16.0 + (34 if show_chrome else 4)) / 72.0
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, nrow + (2.0 if show_chrome else 0.05))
    ax.axis('off')
    fig.subplots_adjust(left=0.004, right=0.996, top=0.996, bottom=0.004)

    if show_chrome:
        ax.add_patch(Rectangle((0, nrow + 1.15), 100, 0.85, facecolor='#f0f0f0', edgecolor='#9a9a9a', lw=0.8))
        ax.text(0.8, nrow + 1.57, f'лист «{sheet.ws.title}» — {title}', fontsize=8.0, va='center', fontweight='bold')
        for i, c in enumerate(range(c1, c2 + 1)):
            ax.add_patch(Rectangle((xs[i], nrow + 0.3), xs[i + 1] - xs[i], 0.8,
                                   facecolor='#e4e4e4', edgecolor='#9a9a9a', lw=0.5))
            ax.text((xs[i] + xs[i + 1]) / 2, nrow + 0.68, get_column_letter(c),
                    ha='center', va='center', fontsize=6.2)
        ax.add_patch(Rectangle((0, nrow + 0.3), 3.4, 0.8, facecolor='#e4e4e4', edgecolor='#9a9a9a', lw=0.5))

    y0 = nrow + (0.3 if show_chrome else 0.0)
    for r in range(r1, r2 + 1):
        ytop = y0 - (r - r1) * 1.0
        if show_chrome:
            ax.add_patch(Rectangle((0, ytop - 0.95), 3.4, 0.95, facecolor='#f2f2f2',
                                   edgecolor='#9a9a9a', lw=0.5))
            ax.text(1.7, ytop - 0.47, str(r), ha='center', va='center', fontsize=6.0)
        for i, c in enumerate(range(c1, c2 + 1)):
            ax.add_patch(Rectangle((xs[i], ytop - 0.95), xs[i + 1] - xs[i], 0.95,
                                   facecolor='white', edgecolor='#d0d0d0', lw=0.45))
            txt = sheet.text(r, c)
            if not txt:
                continue
            x0, x1 = xs[i], xs[i + 1]
            try:
                float(txt.replace(',', '.'))
                num = True
            except ValueError:
                num = False
            if txt in ('≤', '≥', '='):
                ax.text(x0 + 0.35, ytop - 0.47, txt, ha='left', va='center', fontsize=7.6)
            elif num:
                ax.text(x1 - 0.45, ytop - 0.47, txt, ha='right', va='center', fontsize=7.4)
            elif len(txt) <= 4:
                ax.text((x0 + x1) / 2, ytop - 0.47, txt, ha='center', va='center', fontsize=7.2)
            else:
                ax.text(x0 + 0.4, ytop - 0.47, txt, ha='left', va='center', fontsize=7.0)
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, fname)
    fig.savefig(path, dpi=200, facecolor='white', bbox_inches='tight', pad_inches=0.02)
    plt.close(fig)
    return path


SHEETS = [('2.4а', 'метод последовательных уступок (шаги 1–3)'),
          ('2.4б', 'метод ведущего критерия'),
          ('2.4в', 'метод равных относительных отклонений (шаги 1–4)'),
          ('2.4г', 'метод минимакса (шаги 1–4)')]


def main():
    wb = load_workbook(XLSX)
    for name, title in SHEETS:
        ws = wb[name]
        r1, r2 = 1, ws.max_row
        widths_map = {}
        for letter, dim in ws.column_dimensions.items():
            if dim.width:
                widths_map[letter] = dim.width
        p = render(Sheet(ws, r2, 9), r1, r2, 1, 9,
                   f'excel_mmolo_{name.replace(".", "")}.png', title, widths_map)
        print('saved', p)


if __name__ == '__main__':
    main()
