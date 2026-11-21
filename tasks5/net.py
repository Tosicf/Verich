"""Network diagram analysis for task 5.1 — standard СПУ calculations."""


class Network:
    def __init__(self, n_events, works):
        # works: list of (i, j, duration, label)
        self.n = n_events
        self.works = works
        self.out_works = {i: [] for i in range(1, n_events + 1)}
        self.in_works = {j: [] for j in range(1, n_events + 1)}
        for idx, (i, j, t, label) in enumerate(works):
            self.out_works[i].append((idx, j, t, label))
            self.in_works[j].append((idx, i, t, label))

    def event_early(self):
        p = [0.0] * (self.n + 1)
        p[1] = 0.0
        for j in range(2, self.n + 1):
            best = 0.0
            for idx, i, t, label in self.in_works[j]:
                best = max(best, p[i] + t)
            p[j] = best
        return p

    def event_late(self, t_final):
        n = self.n
        tp = [0.0] * (n + 1)
        tp[n] = t_final
        for i in range(n - 1, 0, -1):
            best = None
            for idx, j, t, label in self.out_works[i]:
                val = tp[j] - t
                best = val if best is None else min(best, val)
            tp[i] = best if best is not None else 0.0
        return tp

    def work_params(self):
        p = self.event_early()
        t_final = p[self.n]
        tp = self.event_late(t_final)
        res = []
        for idx, (i, j, t, label) in enumerate(self.works):
            early_start = p[i]
            early_finish = p[i] + t
            late_finish = tp[j]
            late_start = late_finish - t
            full_reserve = tp[j] - p[i] - t
            free_reserve = p[j] - p[i] - t
            independent_reserve = max(0.0, p[j] - tp[i] - t)
            res.append({
                'idx': idx, 'i': i, 'j': j, 't': t, 'label': label,
                'early_start': early_start, 'early_finish': early_finish,
                'late_start': late_start, 'late_finish': late_finish,
                'full_reserve': full_reserve, 'free_reserve': free_reserve,
                'independent_reserve': independent_reserve,
            })
        critical_works = [r for r in res if abs(r['full_reserve']) < 1e-9]
        critical_events = set()
        for r in critical_works:
            critical_events.add(r['i'])
            critical_events.add(r['j'])
        return {
            'event_early': p, 'event_late': tp, 'works': res,
            'critical_time': t_final, 'critical_works': critical_works,
            'critical_events': critical_events,
        }


def format_num(x, nd=2):
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f'{x:.{nd}f}'.rstrip('0').rstrip('.').replace('.', ',')
