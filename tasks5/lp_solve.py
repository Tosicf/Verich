"""Solve 5.3(в) and 5.5(в) via linear programming (Задача 5.1/5.2 models)."""
import numpy as np
from scipy.optimize import linprog


def solve_5_3_5_5(works, t0, B=None, kind='5.3'):
    """
    works: list of dicts/tuples (i, j, t, d, k)
    kind: '5.3' (min sum x, finish in t_0) or '5.5' (min project time, sum x <= B)
    """
    n_work = len(works)
    # variables: t^h_ij, t^o_ij, x_ij for each work (3 per work) + dummy work (if needed)
    event_final = max(j for i, j, t, d, k in works)
    works_arr = works
    # find works ending at final event
    final_works = [(i, j) for (i, j, t, d, k) in works_arr if j == event_final]
    # build dummy work (event_final, event_final+1), duration 0
    dummy_exists = len(final_works) > 1 or kind == '5.5'

    n_ev = event_final + 1
    # variables
    # t_h[i] = start, t_o[i] = end, x[i] = additional money
    # dummy work index = n_work (if exists)
    n_dummy = 1 if dummy_exists else 0
    n_vars = 3 * n_work + 2 * n_dummy  # dummy: t^h, t^o (x=0)

    def idx_h(m): return m
    def idx_o(m): return n_work + m
    def idx_x(m): return 2 * n_work + m
    def dummy_h(): return 3 * n_work
    def dummy_o(): return 3 * n_work + 1

    # objective coefficients
    c = np.zeros(n_vars)
    if kind == '5.3':
        for m in range(n_work):
            c[idx_x(m)] = 1.0
    else:  # 5.5: minimize project time = dummy end (or max end at final)
        if dummy_exists:
            c[dummy_o()] = 1.0
        else:
            # dummy not needed: minimize max of works ending at final event
            # add variable for max? instead minimize the max directly via constraints
            # Simplest: if single work ends at final event, minimize its end time
            m = [m for m, (i, j, t, d, k) in enumerate(works_arr) if j == event_final][0]
            c[idx_o(m)] = 1.0

    # Equality constraints:
    # t_o - t_h = t - k*x   (for each work)
    # t_h=0 for works starting at event 1
    # t_j^h >= t_ij^o for each (i,j),(j,k) pair

    # Bounds:
    bounds = []
    for m in range(n_work):
        bounds.append((0.0, None))  # t_h
        bounds.append((0.0, None))  # t_o
        bounds.append((0.0, None))  # x
    if dummy_exists:
        bounds.append((0.0, None))  # dummy start
        bounds.append((0.0, None))  # dummy end

    # Build linear constraints A_eq, b_eq
    A_eq = []
    b_eq = []
    # t_o - t_h = t - k*x  ->  t_o - t_h + k*x = t
    for m, (i, j, t, d, k) in enumerate(works_arr):
        row = np.zeros(n_vars)
        row[idx_o(m)] = 1.0
        row[idx_h(m)] = -1.0
        row[idx_x(m)] = k
        A_eq.append(row)
        b_eq.append(t)
    # dummy: t_o = t_h + 0 (duration 0)  ->  t_o - t_h = 0
    if dummy_exists:
        row = np.zeros(n_vars)
        row[dummy_o()] = 1.0
        row[dummy_h()] = -1.0
        A_eq.append(row)
        b_eq.append(0.0)

    A_eq = np.array(A_eq)
    b_eq = np.array(b_eq)

    # Inequality constraints
    A_in = []
    b_in = []
    # t_h = 0 for works starting at event 1 -> t_h <= 0 (since >=0)
    for m, (i, j, t, d, k) in enumerate(works_arr):
        if i == 1:
            row = np.zeros(n_vars)
            row[idx_h(m)] = 1.0
            A_in.append(row)
            b_in.append(0.0)
    # t_j^h >= t_ij^o for works (i,j) and (j,k)
    # i.e., t_o(i,j) - t_h(j,k) <= 0
    for m1, (i, j, t, d, k) in enumerate(works_arr):
        for m2, (j2, k, t2, d2, k2) in enumerate(works_arr):
            if j2 == j:  # work 1 ends at event j; work 2 starts at event j
                row = np.zeros(n_vars)
                row[idx_o(m1)] = 1.0
                row[idx_h(m2)] = -1.0
                A_in.append(row)
                b_in.append(0.0)
    # dummy work: all works ending at final event must finish <= dummy start
    if dummy_exists:
        for m, (i, j, t, d, k) in enumerate(works_arr):
            if j == event_final:
                row = np.zeros(n_vars)
                row[idx_o(m)] = 1.0
                row[dummy_h()] = -1.0
                A_in.append(row)
                b_in.append(0.0)
        # objective: dummy end = project time (no upper bound for now; min objective handles it)
    # x_ij <= (t - d)/k (duration >= d)
    for m, (i, j, t, d, k) in enumerate(works_arr):
        row = np.zeros(n_vars)
        row[idx_x(m)] = 1.0
        A_in.append(row)
        b_in.append((t - d) / k)
    # project constraints
    if kind == '5.3':
        # t^o_{i,n} <= t0
        for m, (i, j, t, d, k) in enumerate(works_arr):
            if j == event_final:
                row = np.zeros(n_vars)
                row[idx_o(m)] = 1.0
                A_in.append(row)
                b_in.append(t0)
    else:  # 5.5
        # sum x <= B
        row = np.zeros(n_vars)
        for m in range(n_work):
            row[idx_x(m)] = 1.0
        A_in.append(row)
        b_in.append(B)

    A_in = np.array(A_in) if A_in else np.zeros((0, n_vars))
    b_in = np.array(b_in) if b_in else np.zeros(0)

    res = linprog(c, A_eq=A_eq, b_eq=b_eq, A_ub=A_in, b_ub=b_in, bounds=bounds, method='highs')
    if not res.success:
        print('LP FAILED:', res.message)
        return None
    x = res.x
    out = []
    for m, (i, j, t, d, k) in enumerate(works_arr):
        out.append({
            'i': i, 'j': j, 't': t, 'd': d, 'k': k,
            't_h': x[idx_h(m)], 't_o': x[idx_o(m)], 'x': x[idx_x(m)],
            'new_duration': x[idx_o(m)] - x[idx_h(m)],
        })
    project_time = (res.x[dummy_o()] if dummy_exists else max(w['t_o'] for w in out))
    total_x = sum(w['x'] for w in out)
    return {
        'works': out, 'project_time': project_time, 'total_x': total_x,
        'success': True,
    }


if __name__ == '__main__':
    works_5_3 = [
        (1, 2, 6, 5, 0.05),
        (1, 3, 13, 10, 0.25),
        (1, 4, 20, 16, 0.30),
        (2, 4, 9, 7, 0.07),
        (2, 5, 14, 11, 0.15),
        (3, 4, 16, 13, 0.10),
        (3, 6, 15, 12, 0.05),
        (4, 5, 10, 7, 0.03),
        (4, 6, 17, 15, 0.14),
        (5, 6, 13, 9, 0.50),
    ]
    res = solve_5_3_5_5(works_5_3, 40, kind='5.3')
    print('5.3(в):')
    for w in res['works']:
        print(f"  ({w['i']},{w['j']}): t={w['t']} d={w['d']} k={w['k']} | start={w['t_h']:.3f} finish={w['t_o']:.3f} x={w['x']:.3f} dur={w['new_duration']:.3f}")
    print(f"  project time: {res['project_time']:.3f}, total x: {res['total_x']:.3f}")

    works_5_5 = [
        (1, 2, 15, 12, 0.25),
        (1, 3, 8, 5, 0.20),
        (1, 4, 7, 4, 0.15),
        (2, 3, 5, 3, 0.10),
        (3, 4, 13, 10, 0.30),
        (3, 5, 11, 8, 0.40),
        (4, 5, 7, 4, 0.20),
    ]
    res = solve_5_3_5_5(works_5_5, None, B=47, kind='5.5')
    print('5.5(в):')
    for w in res['works']:
        print(f"  ({w['i']},{w['j']}): t={w['t']} d={w['d']} k={w['k']} | start={w['t_h']:.3f} finish={w['t_o']:.3f} x={w['x']:.3f} dur={w['new_duration']:.3f}")
    print(f"  project time: {res['project_time']:.3f}, total x: {res['total_x']:.3f}")
