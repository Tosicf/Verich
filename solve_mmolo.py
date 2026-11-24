# -*- coding: utf-8 -*-
"""Тема 2, задание 2.4 — численное решение четырьмя методами.

Исходные данные (методичка, табл. 2.3):
    расход ресурсов на ед. продукции 1П..3П:
        Ресурс 1: 2 4 4 | 400
        Ресурс 2: 3 2 2 | 300
        Ресурс 3: 4 5 3 | 500
    прибыль      c1 = (22, 12, 14) -> max   (f1)
    оптовая цена c2 = (24, 16, 30) -> max   (f2)
    трудоёмкость c3 = (20, 10, 15) -> min   (f3)

Результаты сохраняются в /tmp/mmolo_results.json для сборщиков Excel/скриншотов/отчёта.
"""
import json
import numpy as np
from scipy.optimize import linprog

C1 = np.array([22.0, 12.0, 14.0])   # прибыль      -> max
C2 = np.array([24.0, 16.0, 30.0])   # оптовая цена -> max
C3 = np.array([20.0, 10.0, 15.0])   # трудоёмкость -> min
R = np.array([[2.0, 4.0, 4.0],
              [3.0, 2.0, 2.0],
              [4.0, 5.0, 3.0]])
B = np.array([400.0, 300.0, 500.0])


def lp(c, sense, rows, n=3):
    """rows: list of (a, op, b), op in '<=', '>=', '='.  sense: 'max' | 'min'."""
    Aub, bub, Aeq, beq = [], [], [], []
    for a, op, b in rows:
        a = np.asarray(a, dtype=float)
        if op == '<=':
            Aub.append(a); bub.append(b)
        elif op == '>=':
            Aub.append(-a); bub.append(-b)
        else:
            Aeq.append(a); beq.append(b)
    kw = {}
    if Aub:
        kw['A_ub'] = np.array(Aub, dtype=float); kw['b_ub'] = np.array(bub, dtype=float)
    if Aeq:
        kw['A_eq'] = np.array(Aeq, dtype=float); kw['b_eq'] = np.array(beq, dtype=float)
    res = linprog(c=c if sense == 'min' else -c, bounds=[(0, None)] * n, method='highs', **kw)
    if not res.success:
        raise RuntimeError("LP failed: " + res.message)
    return np.round(res.x, 10), res


def report(name, x, extra=None):
    f1 = float(np.dot(C1, x)); f2 = float(np.dot(C2, x)); f3 = float(np.dot(C3, x))
    res = R @ x
    print(f"  {name}: x = ({x[0]:.6g}, {x[1]:.6g}, {x[2]:.6g})  f1={f1:.4f} f2={f2:.4f} f3={f3:.4f}"
          f"  ресурсы=({res[0]:.4f},{res[1]:.4f},{res[2]:.4f})")
    d = dict(x=[float(v) for v in x], f1=f1, f2=f2, f3=f3, res=[float(v) for v in res])
    if extra:
        d.update(extra)
    return d


def base_rows(min_prod=30.0):
    rows = [(R[i], '<=', B[i]) for i in range(3)]
    rows.append(([0, 0, 1], '>=', min_prod))
    return rows


out = {}

print("=== а) метод последовательных уступок (уступки: f1 -> 580, f2 -> 1090) ===")
rows = base_rows()
xa1, _ = lp(C1, 'max', rows); ra = report("шаг 1 (max f1)", xa1)
rows2 = rows + [(C1, '>=', ra['f1'] - 580.0)]
xa2, _ = lp(C2, 'max', rows2); rb = report("шаг 2 (max f2, (5))", xa2)
rows3 = rows2 + [(C2, '>=', rb['f2'] - 1090.0)]
xa3, _ = lp(C3, 'min', rows3); rc = report("шаг 3 (min f3, (6))", xa3)
out['a'] = dict(step1=ra, step2=rb, step3=rc,
                d1=580.0, d2=1090.0,
                cons1=ra['f1'] - 580.0, cons2=rb['f2'] - 1090.0)

print("=== б) метод ведущего критерия (f2>=2400, f3<=1600) ===")
rowsb = base_rows() + [(C2, '>=', 2400.0), (C3, '<=', 1600.0)]
xb, _ = lp(C1, 'max', rowsb); rb2 = report("max f1", xb)
eff = {}
for tag, c, sense in (('f2', C2, 'max'), ('f3', C3, 'min')):
    xs, _ = lp(c, sense, rowsb + [(C1, '>=', rb2['f1'])])
    eff[tag] = dict(x=[float(v) for v in xs], val=float(np.dot(c, xs)))
print("   проверка эффективности: макс f2 =", f"{eff['f2']['val']:.4f}",
      "(в решении", f"{rb2['f2']:.4f})", "; мин f3 =", f"{eff['f3']['val']:.4f}",
      "(в решении", f"{rb2['f3']:.4f})", ")")
out['b'] = dict(sol=rb2, eff=eff, bound_f2=2400.0, bound_f3=1600.0)

print("=== в) метод равных и наименьших относительных отклонений (x1>=20) ===")
rowsv0 = base_rows() + [([1, 0, 0], '>=', 20.0)]
f1v, _ = lp(C1, 'max', rowsv0); f2v, _ = lp(C2, 'max', rowsv0); f3v, _ = lp(C3, 'min', rowsv0)
rv1 = report("max f1", f1v); rv2 = report("max f2", f2v); rv3 = report("min f3", f3v)
F1, F2, F3 = rv1['f1'], rv2['f2'], rv3['f3']
d1, d2, d3 = 1.0 / F1, 1.0 / F2, 1.0 / F3
row6 = [d1 * C1[j] - d2 * C2[j] for j in range(3)]      # одинаковые направления: d1 f1 - d2 f2 = 0
row7 = [d1 * C1[j] + d3 * C3[j] for j in range(3)]      # разные направления:  d1 f1 + d3 f3 = 2
print("   (6):", " ".join(f"{v:+.6f}*x{j+1}" for j, v in enumerate(row6)), "= 0")
print("   (7):", " ".join(f"{v:+.6f}*x{j+1}" for j, v in enumerate(row7)), "= 2")
xv, _ = lp(C1, 'max', rowsv0 + [(row6, '=', 0.0), (row7, '=', 2.0)])
rv4 = report("компромисс", xv)
dev = [abs(rv4['f1'] - F1) / F1, abs(rv4['f2'] - F2) / F2, abs(rv4['f3'] - F3) / F3]
print("   относительные отклонения: " + ", ".join(f"{100*v:.4f}%" for v in dev))
out['c'] = dict(f1=rv1, f2=rv2, f3=rv3, sol=rv4, mode='equal',
                F=[F1, F2, F3], row6=[float(v) for v in row6], row7=[float(v) for v in row7],
                dev=[float(v) for v in dev], y=float(max(dev)))

print("=== г) метод минимакса (x3>=40) ===")
rowsg0 = base_rows(40.0)
g1, _ = lp(C1, 'max', rowsg0); g2, _ = lp(C2, 'max', rowsg0); g3, _ = lp(C3, 'min', rowsg0)
rg1 = report("max f1", g1); rg2 = report("max f2", g2); rg3 = report("min f3", g3)
G1, G2, G3 = rg1['f1'], rg2['f2'], rg3['f3']
# расширенная задача: min x4=y,  f_k + G_k*y >= G_k (k=1,2),  f3 - G3*y <= G3
A = [list(R[i]) + [0.0] for i in range(3)]
b = list(B) + [-40.0]
A.append([0.0, 0.0, -1.0, 0.0])
A.append(list(-C1) + [-G1]); b.append(-G1)
A.append(list(-C2) + [-G2]); b.append(-G2)
A.append(list(C3) + [-G3]); b.append(G3)
rg = linprog(c=[0., 0., 0., 1.], A_ub=np.array(A, dtype=float), b_ub=np.array(b, dtype=float),
             bounds=[(0, None)] * 4, method='highs')
if not rg.success:
    raise RuntimeError("минимакс: " + rg.message)
xg = np.round(rg.x[:3], 10); yg = float(rg.x[3])
rg4 = report("минимакс", xg, extra=dict(x4=yg))
ndev = [abs(rg2['f1'] - G1) / G1, abs(rg2['f2'] - G2) / G2, abs(rg2['f3'] - G3) / G3]
gdev = [abs(rg4['f1'] - G1) / G1, abs(rg4['f2'] - G2) / G2, abs(rg4['f3'] - G3) / G3]
print("   наивное компромиссное (max f2) отклонения:", ", ".join(f"{100*v:.4f}%" for v in ndev),
      "-> max", f"{100*max(ndev):.4f}%")
print("   после минимакса:", ", ".join(f"{100*v:.4f}%" for v in gdev), "-> max", f"{100*max(gdev):.4f}%")
out['d'] = dict(f1=rg1, f2=rg2, f3=rg3, sol=rg4, G=[G1, G2, G3], y=yg,
                naive=rg2, naive_dev=[float(v) for v in ndev], dev=[float(v) for v in gdev])

with open('/tmp/mmolo_results.json', 'w') as fh:
    json.dump(out, fh, ensure_ascii=False, indent=1)
print("\nsaved /tmp/mmolo_results.json")
