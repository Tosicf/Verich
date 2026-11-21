"""5.1(в) network analysis.

Works: 1a –/4; 2a –/6; 3a 2a/2; 4a 2a/6; 5a 1a,3a/3; 6a 1a,3a/3; 7a 4a,6a/5
Network:
  1→2 (1a, 4), 1→3 (2a, 6)
  3→4 (3a, 2), 3→5 (4a, 6)
  2→6 (фиктивная, 0), 4→6 (фиктивная, 0)
  6→7 (5a, 3), 6→8 (6a, 3)
  5→9 (фиктивная, 0), 8→9 (фиктивная, 0)
  9→10 (7a, 5)
"""

WORKS = [
    # (id, label, i, j, dur)
    (0, '1a', 1, 2, 4),
    (1, '2a', 1, 3, 6),
    (2, '3a', 3, 4, 2),
    (3, '4a', 3, 5, 6),
    (4, '5a', 6, 7, 3),
    (5, '6a', 6, 8, 3),
    (6, '7a', 9, 10, 5),
    # фиктивные
    (7, 'Ф(2,6)', 2, 6, 0),
    (8, 'Ф(4,6)', 4, 6, 0),
    (9, 'Ф(5,9)', 5, 9, 0),
    (10, 'Ф(8,9)', 8, 9, 0),
]

N_EVENTS = 10


def early_late():
    # early
    p = [0.0] * (N_EVENTS + 1)
    # topological order by Фалкерson ranks (i<j)
    for j in range(2, N_EVENTS + 1):
        best = 0.0
        for idx, label, i, j_, t in WORKS:
            if j_ == j:
                best = max(best, p[i] + t)
        p[j] = best
    t_final = p[N_EVENTS]
    # late
    tp = [0.0] * (N_EVENTS + 1)
    tp[N_EVENTS] = t_final
    # event 7 (dead end after 5a) treated as terminal -> late = t_final
    tp[7] = max(tp[7], t_final)
    for i in range(N_EVENTS - 1, 0, -1):
        best = None
        for idx, label, j_, j2, t in WORKS:
            if j_ == i:
                val = tp[j2] - t
                best = val if best is None else min(best, val)
        # if no out-works, leave as terminal (event 7)
        tp[i] = best if best is not None else (t_final if i == 7 else tp[i])
    return p, tp, t_final


p, tp, t_final = early_late()
print('Critical time tкр =', t_final)
print()
print('Event early/late:')
for j in range(1, N_EVENTS + 1):
    print(f'  {j}: {p[j]:.0f}/{tp[j]:.0f}  R={tp[j]-p[j]:.0f}')
print()
print('Work params (tп.н./tр.о./tн.н./tн.о./Rп/Rс/Rн):')
for idx, label, i, j, t in WORKS:
    if t == 0:
        continue  # skip фиктивные for the main table, but show them too
    t_nh = p[i]
    t_ro = p[i] + t
    t_nh_late = tp[j] - t
    t_ho = tp[j]
    Rp = tp[j] - p[i] - t
    Rs = p[j] - p[i] - t
    R_indep = max(0.0, p[j] - tp[i] - t)
    print(f'  {label:6} ({i},{j}): {t_nh:.0f}/{t_ro:.0f}/{t_nh_late:.0f}/{t_ho:.0f}  Rп={Rp:.0f} Rс={Rs:.0f} Rн={R_indep:.0f}  (dur={t})')
print()
print('Фиктивные работы:')
for idx, label, i, j, t in WORKS:
    if t == 0:
        Rp = tp[j] - p[i] - t
        Rs = p[j] - p[i] - t
        print(f'  {label:10} ({i},{j}): Rп={Rp:.1f} Rс={Rs:.1f}')

# critical works
print()
print('Critical works (Rп=0 and Rс=0):')
for idx, label, i, j, t in WORKS:
    Rp = tp[j] - p[i] - t
    if abs(Rp) < 1e-9:
        print(f'  {label:6} ({i},{j}): dur={t}')
crit_events = set()
for idx, label, i, j, t in WORKS:
    Rp = tp[j] - p[i] - t
    if abs(Rp) < 1e-9:
        crit_events.add(i); crit_events.add(j)
print('Critical events:', sorted(crit_events))
