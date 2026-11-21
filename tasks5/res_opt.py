"""Resource optimization (task 5.15, non-interruptible works, методичка 7-step algorithm)."""
import copy


class ResNet:
    def __init__(self, n, works):
        # works: list of dicts {i, j, t, r}
        self.n = n
        self.works = works

    def event_early(self, starts, durations):
        p = [0.0] * (self.n + 1)
        for j in range(1, self.n + 1):
            best = 0.0
            for w in self.works:
                if w['j'] == j:
                    # work (i,j): event j = time work ends (or if event j is final)
                    best = max(best, starts[id(w)] + durations[id(w)])
            p[j] = best
        return p

    def linear_chart(self, starts, durations):
        # compute event times from starts/durations via dependency (max over in-works)
        # use effective schedule: each work's start time = specified; compute event times
        p = {}
        for w in self.works:
            i = w['i']
            if i not in p:
                p[i] = 0.0
        # compute event times bottom-up: event = max of work-end for in-works
        # but events also include start times of works -> event = min over out-works of start
        # To be consistent in the Gantt chart, compute event times from the schedule
        starts_i = {}
        ends_i = {}
        for w in self.works:
            starts_i[w['i']] = min(starts_i.get(w['i'], float('inf')), starts[id(w)])
            ends_i[w['j']] = max(ends_i.get(w['j'], 0.0), starts[id(w)] + durations[id(w)])
        p = {}
        for j in range(1, self.n + 1):
            if j in ends_i:
                p[j] = ends_i[j]
            elif j == 1:
                p[j] = 0.0
            else:
                p[j] = 0.0
        return p


def solve_resource_noninterruptible(works_data, R, verbose=True):
    """
    works_data: (i, j, t, r)
    R: available resources
    Returns list of schedules (positions) and final project time.
    """
    works = []
    for idx, (i, j, t, r) in enumerate(works_data):
        works.append({'idx': idx, 'i': i, 'j': j, 't': t, 'r': r})

    n = max(w['j'] for w in works)

    # initial schedule: each work starts at its early time (0 = initial event)
    # compute early event times
    event_early_t = [0.0] * (n + 1)
    for j in range(1, n + 1):
        for w in works:
            if w['j'] == j:
                # event j early time = max of works entering j
                event_early_t[j] = max(event_early_t[j], event_early_t[w['i']] + w['t'])
    # For event 1, no in-works: 0
    final_time = event_early_t[n]

    # initialize starts to early times of their start events
    starts = {}
    for w in works:
        starts[id(w)] = event_early_t[w['i']]
    durations = {id(w): w['t'] for w in works}

    # compute event times and full reserves via the chart
    def compute_event_times(start_map):
        # event = earliest possible time consistent with chart
        p = {}
        p[1] = 0.0
        # topological order: assume events numbered in increasing order of time
        for j in range(1, n + 1):
            in_max = 0.0
            for w in works:
                if w['j'] == j:
                    in_max = max(in_max, start_map[id(w)] + durations[id(w)])
            if j in [w['j'] for w in works]:
                p[j] = in_max
            # for event 1: stays 0
        # for events that have no works ending: they share with predecessors
        for j in range(1, n + 1):
            if j == 1:
                p[j] = 0.0
                continue
            if j not in p:
                # event j started by works starting at it
                out_min = float('inf')
                for w in works:
                    if w['i'] == j:
                        out_min = min(out_min, start_map[id(w)])
                p[j] = out_min if out_min != float('inf') else 0.0
        # ensure p[n] = final time
        p[n] = max(p.get(n, 0.0), max([start_map[id(w)] + durations[id(w)] for w in works if w['j'] == n], default=0.0))
        return p

    def full_reserves(p):
        # late times
        tp = [0.0] * (n + 1)
        tp[n] = p[n]
        for i in range(n - 1, 0, -1):
            best = None
            for w in works:
                if w['i'] == i:
                    val = tp[w['j']] - durations[id(w)]
                    best = val if best is None else min(best, val)
            tp[i] = best if best is not None else 0.0
        reserver = {}
        for w in works:
            fr = tp[w['j']] - p[w['i']] - w['t']
            reserver[id(w)] = fr
        return reserver, tp

    p = compute_event_times(starts)
    fr, tp = full_reserves(p)
    # sort works into a list for iteration
    iteration = 0
    while True:
        iteration += 1
        projection_times = sorted(set(set([starts[id(w)] for w in works]) | set([starts[id(w)] + durations[id(w)] for w in works])))
        # generate all consecutive intervals
        changed_any = False
        # determine t_k / t_{k+1} as we go
        # Process the algorithm on [0, T] where T = project time
        # We'll mimic the методичка's steps by processing intervals sequentially
        # For simplicity: process the projection points in order
        t_points = sorted(set([0.0] + projection_times))
        # Iterate from t=0 through projection points
        cur_t = 0.0
        chart = {}
        for w in works:
            chart[w['idx']] = (starts[id(w)], starts[id(w)] + durations[id(w)])
        # full reserves recomputed per iteration
        # Process intervals
        stop = False
        idx_start = 0
        # iterate over work-start/end points
        # We'll re-evaluate the chart and find the next interval to process
        # For simplicity, process each consecutive interval once per sweep,
        # shifting works that over-burden the resource
        for i in range(len(t_points) - 1):
            t_a, t_b = t_points[i], t_points[i + 1]
            if t_b - t_a < 1e-9:
                continue
            # works covering [t_a, t_b]
            covering = [w for w in works if chart[w['idx']][0] <= t_a + 1e-6 and chart[w['idx']][1] > t_a + 1e-6]
            if not covering:
                continue
            # compute reserves fresh
            p = compute_event_times({id(w): chart[w['idx']][0] for w in works})
            fr, tp = full_reserves(p)
            # for non-interruptible case: works started before t_k (i.e., chart start <= t_a) are numbered
            # in order of (R_п - duration from start to t_b), then other works by R_п
            works_before = [w for w in covering if chart[w['idx']][0] < t_a - 1e-6]
            works_after = [w for w in covering if not (chart[w['idx']][0] < t_a - 1e-6)]
            # make ordering
            works_before_sorted = sorted(works_before, key=lambda w: (fr[id(w)] - (min(t_b, chart[w['idx']][1]) - chart[w['idx']][0]), -w['r']))
            works_after_sorted = sorted(works_after, key=lambda w: (fr[id(w)], -w['r']))
            ordered = works_before_sorted + works_after_sorted
            # Now process with summed intensities
            total_r = 0.0
            for w in ordered:
                if total_r + w['r'] <= R:
                    total_r += w['r']
                else:
                    # shift this work right: non-interruptible case -> start = t_b
                    new_start = t_b
                    # preserve order: start = max(t_b, p(i))
                    new_start = max(new_start, p[w['i']])
                    chart[w['idx']] = (new_start, new_start + durations[id(w)])
                    changed_any = True
                    # note: after shifting, don't add intensity to the sum
            # After processing this interval, continue
            # Recompute chart; proceed
        # After one full pass, check if any changes happened
        if not changed_any:
            # verify no over-demand in any interval
            t_final = max(chart[w['idx']][1] for w in works)
            # iterate: recompute projection points and check resource in each interval
            restart = False
            p = compute_event_times({id(w): chart[w['idx']][0] for w in works})
            all_points = sorted(set([0.0] + [c[0] for c in chart.values()] + [c[1] for c in chart.values()]))
            last_point = 0.0
            for i in range(len(all_points) - 1):
                t_a, t_b = all_points[i], all_points[i+1]
                if t_b - t_a < 1e-9:
                    continue
                total = sum(w['r'] for w in works if chart[w['idx']][0] <= t_a + 1e-6 and chart[w['idx']][1] > t_a + 1e-6)
                if total > R + 1e-9:
                    restart = True
                    break
            if not restart:
                break
        # else loop
        # update starts to the chart's values
        for w in works:
            starts[id(w)] = chart[w['idx']][0]
        # recompute p/tp to check convergence
        p = compute_event_times(starts)
        if iteration > 30:
            break
    # final chart
    t_final = max(starts[id(w)] + durations[id(w)] for w in works)
    return {
        'starts': {w['idx']: starts[id(w)] for w in works},
        'durations': {w['idx']: durations[id(w)] for w in works},
        'project_time': t_final,
    }


if __name__ == '__main__':
    works_data = [
        (1, 2, 6, 4),
        (2, 3, 2, 4),
        (2, 5, 3, 3),
        (1, 3, 4, 5),
        (3, 5, 4, 2),
        (1, 4, 6, 3),
        (4, 5, 3, 4),
    ]
    res = solve_resource_noninterruptible(works_data, 9)
    print('Project time:', res['project_time'])
    for idx, (i, j, t, r) in enumerate(works_data):
        s = res['starts'][idx]
        e = s + res['durations'][idx]
        print(f'  ({i},{j}): dur={t} r={r} | start={s:.2f} end={e:.2f}')
