"""Exact resource-scheduling via MILP (scipy milp / HiGHS).
Works (i,j,t,r), R. All starts integer -> exact for integer durations.
"""
import numpy as np
import scipy.optimize

# works: (i, j, t, r)
WORKS = [
    (1, 2, 4, 8),
    (1, 3, 6, 5),
    (1, 4, 8, 9),
    (2, 4, 2, 10),
    (2, 5, 5, 4),
    (3, 4, 2, 1),
    (3, 5, 3, 2),
    (4, 5, 4, 12),
]
R = 18
T_MAX = 30


def solve_mip(verbose=True):
    n = len(WORKS)
    slots = list(range(T_MAX))
    # variables:
    # 0: T (project duration, integer)
    # 1..n: s_w (integer, start time)
    # then u_{w,k} binary (s_w == k)
    # then y_{w,tau} binary (work w active at slot tau)
    # total number of vars
    n_T = 1
    n_s = n
    # U index region
    u_base = n_T + n_s
    # U vars: n * T_MAX (u_{w,k} for k in 0..T_MAX)
    u_count = n * (T_MAX + 1)
    u_region = n * (T_MAX + 1)
    y_base = u_base + u_region
    # y vars: n * T_MAX
    y_count = n * T_MAX
    total = y_base + y_count

    # objective: minimize T
    c = np.zeros(total)
    c[0] = 1.0

    A_rows = []
    b_l = []
    b_u = []
    # helper to build a row
    def add_row(coeffs, lo, hi):
        row = np.zeros(total)
        idx = []
        val = []
        for col, v in coeffs:
            row[col] = v
        A_rows.append(row)
        b_l.append(lo)
        b_u.append(hi)

    lb = []
    ub = []
    lb.append(1); ub.append(T_MAX)  # T
    for w in range(n):
        lb.append(0); ub.append(T_MAX)  # s_w
    for w in range(n):
        for k in range(T_MAX + 1):
            lb.append(0); ub.append(1)  # u binary
    for w in range(n):
        for tau in range(T_MAX):
            lb.append(0); ub.append(1)  # y binary
    bounds = scipy.optimize.Bounds(lb, ub)

    # rows 1) s_w = sum_k k * u_{w,k}
    for w in range(n):
        coeffs = [(1 + w, -1.0)]
        for k in range(T_MAX + 1):
            coeffs.append((u_base + w * (T_MAX + 1) + k, float(k)))
        add_row(coeffs, 0.0, 0.0)
    # 2) sum_k u_{w,k} = 1 for each w
    for w in range(n):
        coeffs = [(u_base + w * (T_MAX + 1) + k, 1.0) for k in range(T_MAX + 1)]
        add_row(coeffs, 1.0, 1.0)
    # 3) y_{w,tau} = sum_{k in [tau - t + 1, tau]} u_{w,k}
    for w, (i, j, t, r) in enumerate(WORKS):
        for tau in range(T_MAX):
            y_idx = y_base + (w * T_MAX + tau)
            coeffs = [(y_idx, -1.0)]
            lo_k = max(0, tau - t + 1)
            hi_k = min(tau, T_MAX)
            for k in range(lo_k, hi_k + 1):
                coeffs.append((u_base + w * (T_MAX + 1) + k, 1.0))
            add_row(coeffs, 0.0, 0.0)
    # 4) resource: sum_w r_w * y_{w,tau} <= R for each tau
    for tau in range(T_MAX):
        coeffs = []
        for w, (i, j, t, r) in enumerate(WORKS):
            coeffs.append((y_base + w * T_MAX + tau, float(r)))
        add_row(coeffs, -np.inf, float(R))
    # 4.5) works from event 1 start at >= 0 (no pinning; project starts with event 1 at 0)
    # 5) dependency: work (i,j) starts at >= max over in-works of event i of (s + t)
    for w, (i, j, t, r) in enumerate(WORKS):
        for w_in, (i_in, j_in, t_in, r_in) in enumerate(WORKS):
            if j_in == i:
                # s_{i,j} >= s_{i_in,i} + t_in
                coeffs = [(1 + w, 1.0), (1 + w_in, -1.0)]
                add_row(coeffs, float(t_in), np.inf)
    # 6) T >= s_w + t_w
    for w, (i, j, t, r) in enumerate(WORKS):
        coeffs = [(0, 1.0), (1 + w, -1.0)]
        add_row(coeffs, float(t), np.inf)

    A = np.array(A_rows)
    b_l = np.array(b_l)
    b_u = np.array(b_u)

    integrality = np.zeros(total)
    integrality[0] = 1
    for w in range(n):
        integrality[1 + w] = 1  # s_w integer
    for w in range(n):
        for k in range(T_MAX + 1):
            integrality[u_base + w * (T_MAX + 1) + k] = 1
    for w in range(n):
        for tau in range(T_MAX):
            integrality[y_base + w * T_MAX + tau] = 1

    constraints = scipy.optimize.LinearConstraint(A, b_l, b_u)
    res = scipy.optimize.milp(
        c=c,
        constraints=constraints,
        integrality=integrality,
        bounds=bounds,
        options={'time_limit': 180}
    )
    return res


def decode(res, n=None):
    n = n or len(WORKS)
    if res.x is None:
        return None
    T = int(round(res.x[0]))
    starts = {w: int(round(res.x[1 + w])) for w in range(n)}
    return T, starts


if __name__ == '__main__':
    res = solve_mip()
    print('status:', res.status, res.message if hasattr(res, 'message') else '')
    dec = decode(res)
    if dec is None:
        print('no solution')
    else:
        T, starts = dec
        print('Optimal minimal project time T =', T)
        print('Works (i,j): dur=t r',)
        for w, (i, j, t, r) in enumerate(WORKS):
            print(f'  ({i},{j}): start={starts[w]} end={starts[w]+t} dur={t} r={r}')
        print('Resource profile per slot:')
        for tau in range(T):
            tot = sum(r for w, (i, j, t, r) in enumerate(WORKS) if starts[w] <= tau < starts[w] + t)
            print(f'  t={tau}: r={tot}')
