"""Cost optimization algorithm (Задача 5.3, нефиксированное t_kr) for 5.8(в)."""


class Net:
    def __init__(self, n, works):
        # works: (i, j, D, d, c, C)
        self.n = n
        self.works = list(works)  # (i, j, D, d, c, C)
        self.out_works = {i: [] for i in range(1, n + 1)}
        self.in_works = {j: [] for j in range(1, n + 1)}
        for idx, (i, j, D, d, c, C) in enumerate(self.works):
            self.out_works[i].append((idx, j))
            self.in_works[j].append((idx, i))

    def event_early(self):
        p = [0.0] * (self.n + 1)
        p[1] = 0.0
        for j in range(2, self.n + 1):
            best = 0.0
            for idx, i in self.in_works[j]:
                D = self.works[idx][2]
                best = max(best, p[i] + D)
            p[j] = best
        return p

    def event_late(self, t_final):
        tp = [0.0] * (self.n + 1)
        tp[self.n] = t_final
        for i in range(self.n - 1, 0, -1):
            best = None
            for idx, j in self.out_works[i]:
                D = self.works[idx][2]
                val = tp[j] - D
                best = val if best is None else min(best, val)
            tp[i] = best if best is not None else 0.0
        return tp

    def state(self):
        p = self.event_early()
        t_final = p[self.n]
        tp = self.event_late(t_final)
        res = []
        for idx, (i, j, D, d, c, C) in enumerate(self.works):
            full_reserve = tp[j] - p[i] - D
            res.append({
                'idx': idx, 'i': i, 'j': j, 'D': D, 'd': d, 'c': c, 'C': C,
                'p_i': p[i], 'p_j': p[j], 'tp_i': tp[i], 'tp_j': tp[j],
                'full_reserve': full_reserve,
            })
        crit = [r for r in res if abs(r['full_reserve']) < 1e-9]
        return {
            'p': p, 'tp': tp, 'works': res, 'critical_works': crit,
            't_kr': t_final,
        }

    def crit_paths(self):
        """Return list of critical paths, each = list of work indices (path from event 1 to final)."""
        p = self.event_early()
        t_final = p[self.n]
        paths = []
        def dfs(event, path):
            if event == self.n:
                # path = works from 1 to n; check its duration == t_final
                if abs(sum(self.works[idx][2] for idx in path) - t_final) < 1e-9:
                    paths.append(tuple(path))
                return
            for idx, j in self.out_works[event]:
                D = self.works[idx][2]
                if abs(p[j] - p[event] - D) < 1e-9:  # on early-time path
                    dfs(j, path + [idx])
        dfs(1, [])
        return paths


def cost_increase(total_norm_cost, works, durations_now):
    # h = (C - c)/(D_init - d); cost = c + h*(D_init - current_duration)
    total = 0.0
    for (i, j, D_orig, d, c, C), cur, init in zip(works, durations_now, init_info):
        h = init['h']
        total += c + h * (init['D_init'] - cur)
    return total


init_info = []  # populated by solve_cost_optimization


def solve_cost_optimization(n, works, verbose=True):
    """works: (i, j, D, d, c, C). Returns per-step trace and final state."""
    net = Net(n, works)
    total_norm_cost = sum(w[4] for w in works)
    # pre-compute initial КДЗ and store stats in module-level list
    global init_info
    init_info = []
    for (i, j, D, d, c, C) in works:
        init_info.append({'h': (C - c) / (D - d), 'D_init': D})
    trace = []
    step = 0
    max_iter = 50
    while step < max_iter:
        state = net.state()
        crit = state['critical_works']
        t_kr = state['t_kr']
        # check if all critical works at min duration
        to_shorten = [r for r in crit if abs(r['D'] - r['d']) > 1e-9]
        if not to_shorten:
            break
        # get critical paths
        paths = net.crit_paths()
        # keep paths from crit_paths (exact duration match)
        paths = list(paths)
        # КДЗ for each work
        kdz = {w[0]: None for w in works}  # not used
        kdz_val = {}
        for idx, (i, j, D, d, c, C) in enumerate(net.works):
            kdz_val[idx] = init_info[idx]['h']
        # find works to shorten
        work_indices = [r['idx'] for r in to_shorten]
        # shortable critical works only
        # Check paths structure
        paths_set = set()
        for path in paths:
            paths_set.add(tuple(path))
        if len(paths_set) == 0:
            # fallback: single path, shorten min-КДЗ critical work
            selected = [min(to_shorten, key=lambda r: kdz_val[r['idx']])['idx']]
        elif len(paths_set) == 1:
            # single critical path; shorten min-КДЗ among shortable
            selected = [min(to_shorten, key=lambda r: kdz_val[r['idx']])['idx']]
        else:
            # multiple critical paths
            # works common to all paths?
            all_sets = [set(p) for p in paths_set]
            common = set.intersection(*all_sets)
            # the min-КДЗ work on each critical path
            min_on_each = {}
            for path in paths_set:
                shortable = [r['idx'] for r in to_shorten if r['idx'] in path and abs(net.works[r['idx']][2] - net.works[r['idx']][3]) > 1e-9]
                # careful: to_shorten is a list of work dicts; rebuild
                min_on_each[path] = min(shortable, key=lambda idx: kdz_val[idx]) if shortable else None
            # sum of КДЗ of min-КДЗ works on each path
            sum_kdz_per_path = sum(kdz_val[m] for m in min_on_each.values() if m is not None)
            # also consider common works being shortable
            common_shortable = [idx for idx in common]
            # find min-КДЗ common work
            min_common = min(common_shortable, key=lambda idx: kdz_val[idx]) if common_shortable else None
            if len(paths_set) == 1:
                pass
            if not common_shortable:
                # no common works: shorten min-КДЗ work on each path
                selected = [m for m in min_on_each.values() if m is not None]
            else:
                if min_common is not None:
                    k_common = kdz_val[min_common]
                    if abs(sum_kdz_per_path - k_common) < 1e-9:
                        # equal: shorten common + per-path min works
                        selected = list(set([min_common] + [m for m in min_on_each.values() if m is not None]))
                    elif sum_kdz_per_path < k_common:
                        # sum < common: shorten per-path min works only
                        selected = list(set([m for m in min_on_each.values() if m is not None]))
                    else:
                        # sum > common: shorten common work only
                        selected = [min_common]
                else:
                    selected = list(set([m for m in min_on_each.values() if m is not None]))
        # shortening amount: min over shortable works of (D - d), and min non-critical R_п
        non_crit = [r for r in state['works'] if abs(r['full_reserve']) > 1e-9]
        min_reserve = min((r['full_reserve'] for r in non_crit), default=float('inf'))
        amount = None
        for idx in selected:
            w = net.works[idx]
            amount = min(amount, w[2] - w[3]) if amount is not None else w[2] - w[3]
        if min_reserve != float('inf'):
            amount = min(amount, min_reserve) if selected else 0.0
        # also: don't shorten below 0
        if amount is None:
            break
        # apply shortening
        applied = []
        for idx in selected:
            w = net.works[idx]
            new_D = w[2] - amount
            net.works[idx] = (w[0], w[1], new_D, w[3], w[4], w[5])
            applied.append((idx, w[4], new_D, w[2]))
        # total cost
        durations = [w[2] for w in net.works]
        C = cost_increase(total_norm_cost, net.works, durations)
        if verbose:
            print(f'  Step {step+1}: shortened {applied} by {amount:.3f} -> t_kr={net.state()["t_kr"]:.2f}, C={C:.2f}')
        trace.append({
            'step': step + 1, 'shortened': applied, 'amount': amount,
            't_kr': net.state()['t_kr'], 'cost': C,
            'state': net.state(),
        })
        step += 1
    final = net.state()
    durations = [w[2] for w in net.works]
    C_final = cost_increase(total_norm_cost, net.works, durations)
    return {'net': net, 'state': final, 'trace': trace, 'total_cost': C_final, 'norm_cost': total_norm_cost}


if __name__ == '__main__':
    # 5.8(в)
    works = [
        (1, 2, 6, 2, 6, 14),     # (1,2)
        (1, 3, 11, 6, 2, 17),    # (1,3)
        (2, 3, 4, 1, 4, 13),     # (2,3)
        (2, 4, 7, 3, 3, 19),     # (2,4)
        (2, 5, 9, 4, 2, 27),     # (2,5)
        (3, 5, 3, 1, 5, 7),      # (3,5)
        (4, 5, 8, 3, 1, 31),     # (4,5)
    ]
    res = solve_cost_optimization(5, works)
    state = res['state']
    print()
    print('Final: t_kr =', state['t_kr'])
    print('Total cost =', res['total_cost'], '(norm =', res['norm_cost'], ')')
    print()
    print('--- Final work params ---')
    for r in state['works']:
        print(f"  ({r['i']},{r['j']}): D={r['D']} d={r['d']} es={r['p_i']:.2f} ef={r['p_j']:.2f} full_res={r['full_reserve']:.2f}")
    print('Critical paths:', res['net'].crit_paths())
