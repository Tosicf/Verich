# -*- coding: utf-8 -*-
"""Рендер «скриншотов Excel» для задачи 2.4 (тема 2).

Каждый блок модели из mmolo_models.py рисуется как кадр листа Excel в том же
виде, что и в образцовом отчёте (Эконометрика УСР 3.docx): шапка с именами
переменных, строка «значение», строки критериев со значением целевой функции,
строки ограничений с левыми частями, знаком неравенства и правой частью.
"""
import os
import sys

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mmolo_models import models, fmt  # noqa: E402

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'report_assets')

WIDTHS = {'A': 11.5, 'B': 11.5, 'C': 11.5, 'D': 11.5, 'E': 11.5, 'F': 14.5, 'G': 14.5}
ORDER = list('ABCDEFG')
FS = 9.4
DPI = 200
ROW_PT = 16.0
OPER = {'<=': '≤', '>=': '≥', '=': '='}


def ru(v, nd=4):
    s = fmt(v, nd)
    return s.replace('.', ',') if s else ''


def render_block(m, fname):
    nvars = len(m['vars'])
    rows = [{'kind': 'header', 'vars': m['vars']},
            {'kind': 'values', 'label': 'значение', 'values': m['values']}]
    for o in m['objs']:
        rows.append({'kind': 'obj', 'label': o['label'], 'coef': o['coef'], 'value': o['value']})
    for i, c in enumerate(m['cons'], start=1):
        rows.append({'kind': 'cons', 'num': i, 'coef': c['coef'], 'lhs': c['lhs'],
                     'op': OPER[c['op']], 'rhs': c['rhs']})
    n = len(rows)

    xs, acc = {}, 0.0
    for c in ORDER:
        xs[c] = acc
        acc += WIDTHS[c]
    total_w = acc
    fig_h = n * ROW_PT / 72.0 + 0.03
    fig_w = total_w * 7.2 / 72.0
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=DPI)
    ax.set_xlim(0, total_w)
    ax.set_ylim(0, n)
    ax.axis('off')
    fig.subplots_adjust(left=0.004, right=0.996, top=0.995, bottom=0.005)

    def cbox(c):
        return xs[c], WIDTHS[c]

    for r in range(n):
        y_bot = n - r - 1
        for c in ORDER:
            x0, w = cbox(c)
            ax.add_patch(plt.Rectangle((x0, y_bot), w, 1, facecolor='white',
                                       edgecolor='#cfcfcf', lw=0.5, zorder=1))

    def triangle(c, r):
        x0, w = cbox(c)
        y = n - r - 1
        s = 0.30
        ax.add_patch(Polygon([(x0 + w, y + s * .62), (x0 + w, y), (x0 + w - s * 1.15, y)],
                             closed=True, facecolor='#4c9e34', edgecolor='none', zorder=4))

    def put(c, r, txt, ha, va='center'):
        x0, w = cbox(c)
        y = n - r - 0.5
        x = {'left': x0 + 0.35, 'right': x0 + w - 0.35, 'center': x0 + w / 2}[ha]
        t = ax.text(x, y, txt, ha=ha, va=va, fontsize=FS, zorder=5)
        return t

    pending_ops = []
    for r, row in enumerate(rows):
        k = row['kind']
        if k == 'header':
            for i, name in enumerate(row['vars']):
                put(ORDER[1 + i], r, name, 'left', va='bottom')
            continue
        if k == 'values':
            put('A', r, row['label'], 'right')
            for i, v in enumerate(row['values']):
                put(ORDER[1 + i], r, ru(v), 'right')
                triangle(ORDER[1 + i], r)
            continue
        if k == 'obj':
            put('A', r, row['label'], 'right')
            for j, co in enumerate(row['coef']):
                if co:
                    put(ORDER[1 + j], r, ru(co), 'right')
            put('F', r, ru(row['value']), 'right')
            triangle('F', r)
            continue
        put('A', r, str(row['num']), 'right')
        for j, co in enumerate(row['coef']):
            if co:
                put(ORDER[1 + j], r, ru(co), 'right')
        t = put('F', r, ru(row['lhs']), 'right')
        put('G', r, ru(row['rhs']), 'right')
        pending_ops.append((t, row['op']))
        triangle('F', r)
        triangle('G', r)

    # знак неравенства — сразу правее левой части, как в Excel
    fig.canvas.draw()
    rend = fig.canvas.get_renderer()
    for t, op in pending_ops:
        bb = t.get_window_extent(renderer=rend)
        inv = ax.transData.inverted()
        x_right = inv.transform((bb.x1, 0))[0]
        y = n - int(round(n - (inv.transform((0, bb.y0))[1]))) - 0.5
        ax.text(x_right + 0.45, inv.transform((0, (bb.y0 + bb.y1) / 2))[1], op,
                ha='left', va='center', fontsize=FS, zorder=5)

    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, fname)
    fig.savefig(path, dpi=DPI, facecolor='white', bbox_inches='tight', pad_inches=0.015)
    plt.close(fig)
    return path


NAMES = {
    'a1': 'mmolo_a1.png', 'a2': 'mmolo_a2.png', 'a3': 'mmolo_a3.png',
    'b1': 'mmolo_b1.png',
    'c1': 'mmolo_c1.png', 'c2': 'mmolo_c2.png', 'c3': 'mmolo_c3.png', 'c4': 'mmolo_c4.png',
    'd0': 'mmolo_d1.png', 'd1': 'mmolo_d2.png', 'd2': 'mmolo_d3.png', 'd3': 'mmolo_d4.png',
}


def main():
    M = models()
    for k, f in NAMES.items():
        print('saved', render_block(M[k], f))


if __name__ == '__main__':
    main()
