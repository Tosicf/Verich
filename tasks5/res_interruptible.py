"""Interruptible resource optimization via MILP (scipy milp / HiGHS).

Works (i,j) with integer duration t and intensity r.
Interruptions allowed: work may be active in non-contiguous slots.
Precedence: work (i,j) starts at or after the end of all works ending at event i.
Works leaving event 1 (no in-works) start at time 0.
"""
import numpy as np
import scipy.optimize

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

IN_WORKS = {}
for wi, (i, j, t, r) in enumerate(WORKS):
    IN_WORKS[i] = [k for k, (a, b, tt, rr) in enumerate(WORKS) if b == i]


def solve(T, time_limit=60):
    n = len(WORKS)
    slots = list(range(T))
    n_slots = len(slots)
    n_x = n * n_slots
    n_y = n * n_slots
    n_z = n * (T + 1)
    total = n_x + n_y + n_z

    def xi(w, s): return w * n_slots + s
    def yj(w, s): return n_x + w * n_slots + s
    def zj(w, k): return n_x + n_y + w * (T + 1) + k

    A_rows = []
    b_l = []
    b_u = []

    def add(row_dict, lo, hi):
        # row_dict: dict col -> value
        row = np.zeros(total)
        for col, v in row_dict.items():
            row[col] += v
        A_rows.append(row)
        b_l.append(lo)
        b_u.append(hi)

    # 1) total work per work
    for w, (i, j, t, r) in enumerate(WORKS):
        add({xi(w, s): 1.0 for s in range(n_slots)}, float(t), float(t))

    # 2) x <= y
    for w in range(n):
        for s in range(n_slots):
            add({xi(w, s): 1.0, yj(w, s): -1.0}, -np.inf, 0.0)

    # 3) resources
    for s in range(n_slots):
        add({yj(w, s): float(WORKS[w][3]) for w in range(n)}, -np.inf, float(R))

    # 4) exactly one z per work
    for w in range(n):
        add({zj(w, k): 1.0 for k in range(T + 1)}, 1.0, 1.0)

    # 5) no activity before start: for k in 0..T:
    #    sum_{tau<k} x_{w,tau} <= min(k,dur)*(1 - z_{w,k})
    for w in range(n):
        dur = WORKS[w][2]
        for k in range(T + 1):
            cap = min(k, dur)
            row = {xi(w, tau): 1.0 for tau in range(k)}
            row[zj(w, k)] = cap
            add(row, -np.inf, float(cap))

    # 6) works leaving event 1 start at 0
    for w, (i, j, t, r) in enumerate(WORKS):
        if IN_WORKS[i] == []:
            for k in range(1, T + 1):
                add({zj(w, k): 1.0}, -np.inf, 0.0)


    # 7) precedence: s_w >= e_{w'} for each in-work w'
    for w, (i, j, t, r) in enumerate(WORKS):
        inwis = IN_WORKS[i]
        for wprime in inwis:
            for tau in range(n_slots):
                row = {zj(w, k): float(k) for k in range(T + 1)}
                row[yj(wprime, tau)] = -(tau + 1.0)
                add(row, 0.0, np.inf)

    A = np.array(A_rows)
    b_l = np.array(b_l)
    b_u = np.array(b_u)

    integrality = np.zeros(total)
    for w in range(n):
        for s in range(n_slots):
            integrality[yj(w, s)] = 1
        for k in range(T + 1):
            integrality[zj(w, k)] = 1

    bounds = scipy.optimize.Bounds(
        np.zeros(total),
        np.array([1.0] * n_x + [1.0] * n_y + [1.0] * n_z),
    )
    constraints = scipy.optimize.LinearConstraint(A, b_l, b_u)
    res = scipy.optimize.milp(
        c=np.zeros(total),
        constraints=constraints,
        integrality=integrality,
        bounds=bounds,
        options={'time_limit': time_limit, 'mip_rel_gap': 1e-4},
    )
    return res


def report(T, res):
    n = len(WORKS)
    if res.status != 0:
        return f'T={T}: infeasible (status {res.status})'
    x = res.x
    slots = list(range(T))
    n_slots = len(slots)
    n_x = n * n_slots
    n_y = n * n_slots
    lines = [f'T={T}: feasible (status {res.status})']
    # schedule check
    profile = [0.0] * n_slots
    for w, (i, j, t, r) in enumerate(WORKS):
        active = []
        total_work = 0.0
        for s in range(n_slots):
            v = x[w * n_slots + s]
            if v > 1e-3:
                active.append(s)
                total_work += v
                profile[s] += v * r
        s_start = min(active) if active else None
        e_end = max(active) + 1 if active else None
        lines.append(f'  ({i},{j}): dur={t} r={r} start={s_start} end={e_end} slots={active} sum={total_work:.2f}')
    ok = all(p <= R + 1e-6 for p in profile)
    lines.append('  resource profile: ' + ' '.join(f't{s}={profile[s]:.0f}' for s in range(n_slots)))
    lines.append(f'  feasible profile: {ok}')
    return '\n'.join(lines)


def is_precedence_ok(res):
    x = res.x
    n = len(WORKS)
    n_slots = None
    # we don't know T here; instead re-derive from x length
    # x = [n_x + n_y + n_z]; n_x = n*n_slots
    total = len(x)
    n_x = None
    for T_try in range(12, 25):
        if n * T_try + n * T_try + n * (T_try + 1) == total:
            slots = list(range(T_try))
            n_slots = len(slots)
            break
    ok = True
    for wi, (i, j, t, r) in enumerate(WORKS):
        inwis = IN_WORKS[i]
        if not inwis:
            continue
        # end times of in-works
        for wprime in inwis:
            a_slots = [s for s in range(n_slots) if x[wprime * n_slots + s] > 1e-3]
            if not a_slots:
                ok = False
                continue
            e_end = max(a_slots) + 1
            # start of this work
            b_slots = [s for s in range(n_slots) if x[wi * n_slots + s] > 1e-3]
            if b_slots and min(b_slots) < e_end:
                ok = False
    return ok


if __name__ == '__main__':
    for T in range(12, 22):
        res = solve(T)
        print(report(T, res))
        if res.status == 0:
            ok = is_precedence_ok(res)
            print(f'  precedence check: {ok}')
            break
