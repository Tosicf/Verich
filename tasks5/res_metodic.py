"""Методичка resource optimization (Задача 5.15-5.16, both a) and б) variants).
Faithful implementation: preliminary step + algorithm with shifts.
"""
import copy

WORKS = [
    # (i, j, t, r)
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
N_EVENTS = 5


def event_early_times(starts, durs, works):
    ev = {1: 0.0}
    for j in range(1, N_EVENTS + 1):
        mx = None
        for w in works:
            i, j2, t, r = w
            if j2 == j:
                val = starts[(i, j2)] + durs[(i, j2)]
                mx = val if mx is None else max(mx, val)
        ev[j] = mx if mx is not None else 0.0
    return ev


def project_time(starts, durs, works):
    return max(starts[(i, j)] + durs[(i, j)] for i, j, t, r in works)


def event_late_times(starts, durs, works):
    T = project_time(starts, durs, works)
    ev = {}
    ev[N_EVENTS] = T
    for i in range(N_EVENTS - 1, 0, -1):
        best = None
        for w in works:
            i2, j2, t, r = w
            if i2 == i:
                val = ev[j2] - durs[(i2, j2)]
                best = val if best is None else min(best, val)
        ev[i] = best
    return ev


def solve_noninterruptible(verbose=True):
    # state: starts / durs dicts keyed by (i,j) work id
    works = [(i, j, t, r) for (i, j, t, r) in WORKS]
    durs = {(i, j): t for (i, j, t, r) in works}
    starts = {}
    # initial chart: event 1's works start at 0; others at event_early
    t0 = 0.0
    starts = {}
    ev_early = {1: 0.0}
    for j in range(1, N_EVENTS + 1):
        ev_early[j] = 0.0
        for w in works:
            i, j2, t, r = w
            if j2 == j:
                ev_early[j] = max(ev_early[j], ev_early[i] + durs[(i, j)])
    for w in works:
        i, j, t, r = w
        starts[(i, j)] = ev_early[i]
    if verbose:
        print('Initial chart:')
        print_chart(starts, durs, works)

    t_k = 0.0
    step = 1
    while True:
        # projections: starts/ends of works whose spans intersect [t_k, inf)
        proj = set()
        for w in works:
            i, j, t, r = w
            if starts[(i, j)] + durs[(i, j)] > t_k + 1e-9:
                proj.add(starts[(i, j)])
                proj.add(starts[(i, j)] + durs[(i, j)])
        proj = sorted(proj)
        candidate = [x for x in proj if x > t_k + 1e-9]
        if not candidate:
            break
        t_next = min(candidate)  # t_{k+1}
        # works located over [t_k, t_next]: span intersects [t_k, t_next]
        loc = [w for w in works if starts[(w[0], w[1])] < t_next - 1e-9 and
               starts[(w[0], w[1])] + durs[(w[0], w[1])] > t_k + 1e-9]
        # recompute events for current chart
        ev_late = event_late_times(starts, durs, works)
        reserves = {}
        for w in loc:
            i, j, t, r = w
            R = ev_late[j] - starts[(i, j)] - durs[(i, j)]
            reserves[(i, j)] = R
        # works started left of t_k
        early_started = [w for w in loc if starts[(w[0], w[1])] < t_k - 1e-9]
        others = [w for w in loc if not (starts[(w[0], w[1])] < t_k - 1e-9)]
        early_sorted = sorted(early_started, key=lambda w: (reserves[(w[0], w[1])] - t_next, -w[3]))
        others_sorted = sorted(others, key=lambda w: (reserves[(w[0], w[1])], -w[3]))
        ordered = early_sorted + others_sorted
        # sum intensities in order, shift works whose sum exceeds R
        total = 0.0
        for w in ordered:
            i, j, t, r = w
            if total + r <= R + 1e-9:
                total += r
            else:
                # shift right: start becomes t_next
                new_start = t_next
                starts[(i, j)] = new_start
                if verbose:
                    print(f'  shift ({i},{j}): start {t_k} -> {new_start} (t={t}, r={r})')
        if verbose:
            print(f'  step {step} [{t_k},{t_next}]: works {[(w[0],w[1]) for w in loc]} -> {[(w[0],w[1], round(reserves[(w[0],w[1])],2)) for w in ordered]}')
        t_k = t_next
        step += 1
        if step > 100:
            raise RuntimeError('too many steps')
    if verbose:
        print('Final chart ( методичка ):')
        print_chart(starts, durs, works)
    return starts, durs, works


def print_chart(starts, durs, works):
    for w in works:
        i, j, t, r = w
        s = starts[(i, j)]
        print(f'  ({i},{j}): dur={t} r={r} | [{s:.1f},{s+durs[(i,j)]:.1f}]')
    T = max(starts[(i, j)] + durs[(i, j)] for i, j, t, r in works)
    print(f'  Project time T = {T}')
    # resource profile
    print('  Resource profile (per integer slot):')
    for tau in range(int(T)):
        tot = sum(r for i, j, t, r in works if starts[(i, j)] <= tau < starts[(i, j)] + durs[(i, j)])
        print(f'    t={tau}: r={tot}')


if __name__ == '__main__':
    solve_noninterruptible(verbose=True)
