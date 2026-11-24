# -*- coding: utf-8 -*-
"""Тема 3 «Модели анализа и оценки инвестиционных проектов».

Задания (нумерация в методичке сбита — ищем по содержанию):

  3.4  (напеч. «1.4», с. 107) Начальные инвестиции 150 ден. ед., срок 2 года,
       денежные потоки 150 и 150, ставка дисконтирования 15 %.
       Найти уровень инфляции, до которого проект эффективен.
  3.6  (напеч. «1.6», с. 107) Инвестиции 800 млн, срок 5 лет, потоки по 300 млн,
       норма дисконтирования 12 %, ожидаемый уровень инфляции 10, 12, 18 %.
       Целесообразно ли инвестировать?
  3.12 (с. 109, табл. 3.5) Шесть проектов, бюджет 100 млн руб. в начале 1-го года,
       взаимоисключающие пары (1,2) и (3,5), выпуск однородной продукции 400/110/150
       ед. при спросе не более 550 ед., ставка дисконтирования 15 %.

Результат — /tmp/invest_results.json (для Excel-книги, скриншотов и отчёта).
"""
import json

from scipy.optimize import brentq

R_RATE = 0.15          # 3.4 — ставка дисконтирования
I0 = 150.0
C = [150.0, 150.0]

R6 = 0.12              # 3.6 — норма дисконтирования
I06 = 800.0
C6 = [300.0] * 5
INFL = [0.10, 0.12, 0.18]

# 3.12 — табл. 3.5 методички: денежные потоки по годам 0, 2, 3, 4 (млн руб.).
# У проектов 1 и 4 графа года 0 пуста — их единственный отток приходится на 2-й год.
YEARS = [0, 1, 2, 3, 4]
PROJ = [
    dict(name='1', cf={2: -90.0, 3: 50.0, 4: 100.0}),
    dict(name='2', cf={0: -50.0, 2: 40.0, 3: 70.0, 4: 110.0}),
    dict(name='3', cf={0: -40.0, 2: 30.0, 3: 30.0, 4: 30.0}),
    dict(name='4', cf={2: -40.0, 3: 20.0, 4: 80.0}),
    dict(name='5', cf={0: -60.0, 2: 20.0, 3: 70.0, 4: 50.0}),
    dict(name='6', cf={0: -100.0, 2: 70.0, 3: 140.0, 4: 100.0}),
]
PROD_UNITS = {'2': 400.0, '3': 110.0, '5': 150.0}
DEMAND = 550.0
BUDGET = 100.0
R12 = 0.15
EXCLUDE = [(0, 1), (2, 4)]      # номера проектов в списке PROJ (1-й и 2-й, 3-й и 5-й)


def pv(cash, r):
    """Приведённая стоимость потока {год: значение} по ставке r."""
    return sum(cf / (1 + r) ** t for t, cf in cash.items())


def nominal(base, i):
    """Номинальная ставка с учётом инфляции: r = r0 + i + r0·i  (3.9)."""
    return base + i + base * i


def crit(cash0, flows, r):
    """Полный набор показателей эффективности проекта."""
    n = max(flows) if flows else 0
    npv = pv(flows, r) + cash0                      # cash0 < 0 — начальные инвестиции
    irr = brentq(lambda x: pv(flows, x) + cash0, -0.9, 5.0)
    pi = pv(flows, r) / -cash0
    cum, pp = 0.0, None
    for t in sorted(flows):
        cum += flows[t]
        if cum >= -cash0:
            pp = t
            break
    cum, dpp = 0.0, None
    for t in sorted(flows):
        cum += flows[t] / (1 + r) ** t
        if cum >= -cash0:
            dpp = t
            break
    fv = sum(cf * (1 + r) ** (n - t) for t, cf in flows.items() if cf > 0)
    pv_out = -cash0
    mirr = (fv / pv_out) ** (1.0 / n) - 1 if pv_out > 0 and fv > 0 else None
    return dict(npv=npv, irr=irr, pi=pi, pp=pp, dpp=dpp, mirr=mirr,
                pv=pv(flows, r), inv=-cash0, n=n)


out = {}

# ================================ 3.4 ========================================
def npv34(i):
    return pv({1: C[0], 2: C[1]}, nominal(R_RATE, i)) - I0


base34 = crit(-I0, {1: C[0], 2: C[1]}, R_RATE)
i_max = brentq(npv34, 0.0, 2.0)
rows34 = []
for i in (0.0, 0.05, 0.10, 0.20, 0.30, 0.40, round(i_max, 6), 0.50):
    d = nominal(R_RATE, i)
    pvs = {t: c / (1 + d) ** t for t, c in enumerate(C, start=1)}
    rows34.append(dict(i=i, d=d, pv1=pvs[1], pv2=pvs[2], pv=sum(pvs.values()),
                       npv=sum(pvs.values()) - I0, pi=sum(pvs.values()) / I0))
out['t34'] = dict(I0=I0, C=C, r=R_RATE, base=base34, i_max=float(i_max), rows=rows34,
                  r_at_imax=float(nominal(R_RATE, i_max)),
                  curve=[[i, float(npv34(i))] for i in [x / 1000 for x in range(0, 701)]])

# ================================ 3.6 ========================================
base36 = crit(-I06, {t: c for t, c in enumerate(C6, start=1)}, R6)
i_crit = brentq(lambda i: pv({t: c for t, c in enumerate(C6, start=1)}, nominal(R6, i)) - I06,
                0.0, 2.0)
rows6 = []
for i in INFL + [i_crit]:
    d = nominal(R6, i)
    k = crit(-I06, {t: c for t, c in enumerate(C6, start=1)}, d)
    rows6.append(dict(i=float(i), d=float(d), **{kk: vv for kk, vv in k.items()},
                      critical=abs(float(i) - i_crit) < 1e-9))
out['t36'] = dict(I0=I06, C=C6, r=R6, infl=[float(x) for x in INFL], base=base36,
                  rows=rows6, i_crit=float(i_crit),
                  curve=[[i, float(pv({t: c for t, c in enumerate(C6, start=1)},
                                      nominal(R6, i)) - I06)]
                         for i in [x / 1000 for x in range(0, 351)]])

# ================================ 3.12 =======================================
disc = {t: 1 / (1 + R12) ** t for t in YEARS}
for p in PROJ:
    p['npv'] = pv(p['cf'], R12)
    p['invest'] = max(-v for v in p['cf'].values() if v < 0)
    p['units'] = PROD_UNITS.get(p['name'], 0.0)
    p['pv_in'] = sum(v * disc[t] for t, v in p['cf'].items() if v < 0)
    p['irr'] = None
    try:
        p['irr'] = brentq(lambda x: pv(p['cf'], x), -0.8, 3.0)
    except ValueError:
        pass

best, all_feas = None, []
for mask in range(1 << len(PROJ)):
    idx = [i for i in range(len(PROJ)) if mask >> i & 1]
    if not idx:
        continue
    if any(a in idx and b in idx for a, b in EXCLUDE):
        continue
    sel = [PROJ[i] for i in idx]
    spend = sum(p['invest'] for p in sel)
    units = sum(p['units'] for p in sel)
    if spend > BUDGET + 1e-9 or units > DEMAND + 1e-9:
        continue
    total = sum(p['npv'] for p in sel)
    all_feas.append(dict(names=[p['name'] for p in sel], spend=spend, units=units,
                         npv=total, n_rejected=0))
all_feas.sort(key=lambda z: -z['npv'])
best = all_feas[0]
out['t312'] = dict(budget=BUDGET, demand=DEMAND, r=R12, disc=disc,
                   projects=[dict(name=p['name'], cf={str(t): v for t, v in p['cf'].items()},
                                  npv=p['npv'], invest=p['invest'], units=p['units'],
                                  irr=p['irr']) for p in PROJ],
                   best=best, all_feas=all_feas,
                   n_all=1 << len(PROJ),
                   n_excl=sum(1 for m in range(1 << len(PROJ))
                              if any(a in [i for i in range(len(PROJ)) if m >> i & 1]
                                     and b in [i for i in range(len(PROJ)) if m >> i & 1]
                                     for a, b in EXCLUDE)),
                   n_violate=sum(1 for m in range(1 << len(PROJ))
                                 if m and (sum(max(-v for v in PROJ[i]['cf'].values() if v < 0)
                                               for i in [j for j in range(len(PROJ)) if m >> j & 1])
                                          > BUDGET + 1e-9
                                          or sum(PROJ[i]['units']
                                                 for i in [j for j in range(len(PROJ))
                                                           if m >> j & 1]) > DEMAND + 1e-9)))

with open('/tmp/invest_results.json', 'w') as fh:
    json.dump(out, fh, ensure_ascii=False, indent=1)

b = out['t34']['base']
print("3.4  NPV = %.4f, IRR = %.4f%%, PI = %.4f, PP = %s, DPP = %s"
      % (b['npv'], 100 * b['irr'], b['pi'], b['pp'], b['dpp']))
print("     критический уровень инфляции i = %.6f (%.4f%%), r_ном = %.4f%%"
      % (i_max, 100 * i_max, 100 * nominal(R_RATE, i_max)))
for row in rows34:
    print("     i=%6.2f%%  r=%6.2f%%  PV=%9.4f  NPV=%9.4f  PI=%6.4f"
          % (100 * row['i'], 100 * row['d'], row['pv'], row['npv'], row['pi']))
b = out['t36']['base']
print("\n3.6  без инфляции: NPV = %.4f, IRR = %.4f%%, PI = %.4f, PP = %s, DPP = %s, MIRR = %.4f%%"
      % (b['npv'], 100 * b['irr'], b['pi'], b['pp'], b['dpp'], 100 * b['mirr']))
for row in rows6:
    print("     i=%5.2f%%  r=%6.2f%%  NPV=%9.4f  PI=%6.4f  IRR=%6.2f%% %s"
          % (100 * row['i'], 100 * row['d'], row['npv'], row['pi'], 100 * row['irr'],
             '<<< критическое' if row['critical'] else ''))
print("\n3.12 NPV проектов:", {p['name']: round(p['npv'], 4) for p in PROJ})
print("     IRR проектов:", {p['name']: (None if p['irr'] is None else round(100 * p['irr'], 2))
                             for p in PROJ})
print("     всего наборов %d, с конфликтом %d, вне ограничений %d, допустимых %d"
      % (out['t312']['n_all'], out['t312']['n_excl'], out['t312']['n_violate'], len(all_feas)))
print("     оптимум:", best['names'], "NPV = %.4f, ΣI0 = %.1f, выпуск = %.0f"
      % (best['npv'], best['spend'], best['units']))
print("     топ-8:", [('[,]'.join(z['names']), round(z['npv'], 3)) for z in all_feas[:8]])
print("\nsaved /tmp/invest_results.json")
