# -*- coding: utf-8 -*-
"""Модели задачи 2.4 (тема 2 «Модели и методы многоцелевой оптимизации»).

Единый источник описания листов Excel-решения: используется и для сборки
рабочей книги (build_mmolo_excel.py), и для рендера «скриншотов Excel»
(make_mmolo_screenshots.py), и для текстового отчёта (build_mmolo_report.py).

Каждый блок («состояние листа») = то, что видно в Excel после очередного
вызова Поиска решения:

    vars   — имена переменных (столбцы B..D/E)
    values — решения (строка «значение»)
    objs   — строки критериев: (метка, коэффициенты)
    cons   — строки ограничений: (коэффициенты, знак, правая часть, откуда ограничение)
"""
import copy
import json
import os

# ---- исходные данные (табл. 2.3 методички) ----------------------------------
C1 = (22, 12, 14)      # прибыль        -> max
C2 = (24, 16, 30)      # оптовая цена   -> max
C3 = (20, 10, 15)      # трудоёмкость   -> min
RES = ((2, 4, 4), (3, 2, 2), (4, 5, 3))
CAP = (400, 300, 500)
RES_LABELS = ("Ресурс 1", "Ресурс 2", "Ресурс 3")


def res_rows():
    return [(list(RES[i]), "<=", CAP[i], f"расход {RES_LABELS[i].lower()}") for i in range(3)]


RESULTS_PATH = '/tmp/mmolo_results.json'


def load_results():
    with open(RESULTS_PATH) as fh:
        return json.load(fh)


def fmt(v, nd=4):
    """Число так, как его показал бы Excel (плавающий формат, 4 знака макс.)."""
    if v is None:
        return ''
    v = round(float(v), nd)
    if abs(v) < 10 ** (-nd) / 2:
        return '0'
    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))
    s = f"{v:.{nd}f}".rstrip('0').rstrip('.')
    return s


def _mk(block, res):
    """Нормализует строки ограничений и добавляет вычисленные значения LHS.

    Блоки могут делать одни и те же словари строк (общие списки моделей),
    поэтому перед заполнением значений всё копируется.
    """
    block['objs'] = copy.deepcopy(block['objs'])
    block['cons'] = copy.deepcopy(block['cons'])
    block["values"] = [round(float(v), 6) for v in block["values"]]
    for o in block['objs']:
        o["coef"] = [round(float(v), 6) for v in o["coef"]]
    x = block['values']
    norm = []
    for c in block['cons']:
        if isinstance(c, tuple):
            coef, op, rhs, note = c
            c = dict(coef=list(coef), op=op, rhs=rhs, note=note, value=None, lhs=None)
        c["coef"] = [round(float(v), 6) for v in c["coef"]]
        c["rhs"] = round(float(c["rhs"]), 6)
        norm.append(c)
    block['cons'] = norm
    for row in block['objs']:
        row['value'] = sum(c * v for c, v in zip(row['coef'], x))
    for c in block['cons']:
        c['lhs'] = sum(co * v for co, v in zip(c['coef'], x))
    return block


def _row(coef, op, rhs, note, x=None):
    return dict(coef=list(coef), op=op, rhs=rhs, note=note, value=None, lhs=None)


def build_models(res):
    a, b, c, d = res['a'], res['b'], res['c'], res['d']
    M = {}

    # ---------- а) метод последовательных уступок ----------
    x1 = a['step1']['x']; x2 = a['step2']['x']; x3 = a['step3']['x']
    M['a1'] = dict(
        title="а) шаг 1 — максимизация f1",
        sheet="2.4а", vars=["x1", "x2", "x3"], values=list(x1),
        objs=[dict(label="max f1", coef=list(C1)), dict(label="max f2", coef=list(C2)),
              dict(label="min f3", coef=list(C3))],
        cons=res_rows() + [_row([0, 0, 1], ">=", 30, "объём 3П не менее 30 ед.")])
    M['a2'] = dict(
        title="а) шаг 2 — максимизация f2 с уступкой по f1",
        sheet="2.4а", vars=["x1", "x2", "x3"], values=list(x2),
        objs=[dict(label="max f1", coef=list(C1)), dict(label="max f2", coef=list(C2)),
              dict(label="min f3", coef=list(C3))],
        cons=res_rows() + [_row([0, 0, 1], ">=", 30, "объём 3П не менее 30 ед."),
                           _row(list(C1), ">=", a['cons1'], "уступка по критерию f1: f1 ≥ 2180 − 580  (5)")])
    M['a3'] = dict(
        title="а) шаг 3 — минимизация f3 с уступками по f1 и f2 (итог)",
        sheet="2.4а", vars=["x1", "x2", "x3"], values=list(x3),
        objs=[dict(label="max f1", coef=list(C1)), dict(label="max f2", coef=list(C2)),
              dict(label="min f3", coef=list(C3))],
        cons=res_rows() + [_row([0, 0, 1], ">=", 30, "объём 3П не менее 30 ед."),
                           _row(list(C1), ">=", a['cons1'], "уступка по критерию f1  (5)"),
                           _row(list(C2), ">=", a['cons2'], "уступка по критерию f2: f2 ≥ 3450 − 1090  (6)")])

    # ---------- б) метод ведущего критерия ----------
    xb = b['sol']['x']
    M['b1'] = dict(
        title="б) метод ведущего критерия — max f1 при ограничениях по f2 и f3",
        sheet="2.4б", vars=["x1", "x2", "x3"], values=list(xb),
        objs=[dict(label="max f1", coef=list(C1))],
        cons=res_rows() + [_row([0, 0, 1], ">=", 30, "объём 3П не менее 30 ед."),
                           _row(list(C2), ">=", b['bound_f2'], "нижняя граница оптовой цены"),
                           _row(list(C3), "<=", b['bound_f3'], "трудоемкость не более 1600")])

    # ---------- в) метод равных относительных отклонений ----------
    F1, F2, F3 = c['F']
    xw1 = c['f1']['x']; xw2 = c['f2']['x']; xw3 = c['f3']['x']; xw4 = c['sol']['x']
    objs_w = [dict(label="max f1", coef=list(C1)), dict(label="max f2", coef=list(C2)),
              dict(label="min f3", coef=list(C3))]
    cons_w0 = res_rows() + [_row([0, 0, 1], ">=", 30, "объём 3П не менее 30 ед."),
                            _row([1, 0, 0], ">=", 20, "объём 1П не менее 20 ед.")]
    M['c1'] = dict(title="в) шаг 1 — max f1", sheet="2.4в", vars=["x1", "x2", "x3"],
                   values=list(xw1), objs=objs_w, cons=cons_w0)
    M['c2'] = dict(title="в) шаг 2 — max f2", sheet="2.4в", vars=["x1", "x2", "x3"],
                   values=list(xw2), objs=objs_w, cons=cons_w0)
    M['c3'] = dict(title="в) шаг 3 — min f3", sheet="2.4в", vars=["x1", "x2", "x3"],
                   values=list(xw3), objs=objs_w, cons=cons_w0)
    M['c4'] = dict(title="в) шаг 4 — компромисс с ограничениями (6), (7)", sheet="2.4в",
                   vars=["x1", "x2", "x3"], values=list(xw4), objs=objs_w,
                   cons=cons_w0 + [_row(c['row6'], "=", 0, "равенство отклонений f1 и f2  (6)"),
                                    _row(c['row7'], "=", 2, "равенство отклонений f1 и f3  (7)")])

    # ---------- г) метод минимакса ----------
    G1, G2, G3 = d['G']
    xg1 = d['f1']['x']; xg2 = d['f2']['x']; xg3 = d['f3']['x']; xg4 = d['sol']['x']
    objs_g = [dict(label="max f1", coef=list(C1)), dict(label="max f2", coef=list(C2)),
              dict(label="min f3", coef=list(C3)), dict(label="min f4", coef=[0, 0, 0, 1])]
    cons_g0 = [(list(r[0]) + [0], r[1], r[2], r[3]) for r in res_rows()] + \
              [_row([0, 0, 1, 0], ">=", 40, "объём 3П не менее 40 ед.")]
    M['d0'] = dict(title="г) шаг 1 — max f1", sheet="2.4г", vars=["x1", "x2", "x3"],
                   values=list(xg1),
                   objs=[dict(label="max f1", coef=list(C1)), dict(label="max f2", coef=list(C2)),
                         dict(label="min f3", coef=list(C3))],
                   cons=res_rows() + [_row([0, 0, 1], ">=", 40, "объём 3П не менее 40 ед.")])
    M['d1'] = dict(title="г) шаг 2 — max f2", sheet="2.4г", vars=["x1", "x2", "x3"],
                   values=list(xg2),
                   objs=[dict(label="max f1", coef=list(C1)), dict(label="max f2", coef=list(C2)),
                         dict(label="min f3", coef=list(C3))],
                   cons=res_rows() + [_row([0, 0, 1], ">=", 40, "объём 3П не менее 40 ед.")])
    M['d2'] = dict(title="г) шаг 3 — min f3", sheet="2.4г", vars=["x1", "x2", "x3"],
                   values=list(xg3),
                   objs=[dict(label="max f1", coef=list(C1)), dict(label="max f2", coef=list(C2)),
                         dict(label="min f3", coef=list(C3))],
                   cons=res_rows() + [_row([0, 0, 1], ">=", 40, "объём 3П не менее 40 ед.")])
    M['d3'] = dict(title="г) шаг 4 — минимаксная задача (расширенная, с x4 = y)", sheet="2.4г",
                   vars=["x1", "x2", "x3", "x4"], values=list(xg4) + [d['y']],
                   objs=objs_g,
                   cons=cons_g0 + [
                       _row(list(C1) + [G1], ">=", G1, f"f1 ≥ {fmt(G1)}·(1 − x4)   (2.7)"),
                       _row(list(C2) + [G2], ">=", G2, f"f2 ≥ {fmt(G2)}·(1 − x4)   (2.7)"),
                       _row(list(C3) + [-G3], "<=", G3, f"f3 ≤ {fmt(G3)}·(1 + x4)   (2.8)")])
    for k in M:
        _mk(M[k], res)
    return M


def models():
    return build_models(load_results())


if __name__ == '__main__':
    for k, m in models().items():
        print("=" * 88)
        print(k, m['title'], "| лист", m['sheet'])
        print("   x =", [fmt(v) for v in m['values']])
        for o in m['objs']:
            print(f"   {o['label']:<16}", [fmt(c) for c in o['coef']], "->", fmt(o['value']))
        for c in m['cons']:
            print("   cons", [fmt(v) for v in c['coef']], c['op'], fmt(c['rhs']),
                  "| LHS =", fmt(c['lhs']), "|", c['note'])
