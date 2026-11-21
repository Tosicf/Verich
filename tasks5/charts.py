"""Generate network diagrams and Gantt charts (black & white)."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, FancyArrowPatch
import matplotlib.font_manager as fm

plt.rcParams['font.family'] = 'DejaVu Serif'
plt.rcParams['font.size'] = 13
plt.rcParams['axes.linewidth'] = 1.0

def draw_network(pos, works, width=7.2, height=4.6, dpi=200, label_offset=0.10):
    fig, ax = plt.subplots(figsize=(width, height))
    ax.set_xlim(-0.5, 11.2)
    ax.set_ylim(0.0, 6.3)
    ax.axis('off')
    # events
    for e, (x, y) in pos.items():
        c = Circle((x, y), 0.26, fill=False, edgecolor='black', linewidth=1.6, zorder=6)
        ax.add_patch(c)
        ax.text(x, y, str(e), ha='center', va='center', fontsize=14, zorder=7, family='DejaVu Serif', fontweight='bold')
    # works
    for w in works:
        i, j, t, r = w[0], w[1], w[2], w[3]
        xi, yi = pos[i]; xj, yj = pos[j]
        is_fict = (t == 0)
        # arrow line from edge to edge
        import math
        dx = xj - xi; dy = yj - yi
        L = math.hypot(dx, dy)
        ux, uy = dx / L, dy / L
        start = (xi + ux * 0.28, yi + uy * 0.28)
        end = (xj - ux * 0.28, yj - uy * 0.28)
        ax.annotate('', xy=end, xytext=start,
                    arrowprops=dict(arrowstyle='-|>', lw=1.4 if not is_fict else 1.2,
                                    color='black', linestyle='--' if is_fict else '-',
                                    mutation_scale=14), zorder=4)
        if is_fict:
            label = w[4] if len(w) > 4 else 'Ф'
            ax.text((xi+xj)/2, (yi+yj)/2 + 0.06*(-uy), label, ha='center', va='center', fontsize=11, style='italic', color='black', zorder=5)
        else:
            label = f'{t},{r}' if r is not None else f'{t}'
            # perpendicular offset
            offx, offy = -uy * 0.22, ux * 0.22
            ax.text((xi+xj)/2 + offx, (yi+yj)/2 + offy, label, ha='center', va='center', fontsize=12, zorder=5,
                    bbox=dict(fc='white', ec='none', pad=1.2))
    fig.tight_layout()
    return fig

def save_net(fig, path, dpi=200):
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white', edgecolor='none')
    plt.close(fig)

# ---------------- 5.1(в) ----------------
def chart_5_1():
    pos = {
        1: (0.4, 4.0),
        2: (2.0, 5.6),
        3: (2.0, 2.6),
        4: (4.1, 1.8),
        5: (4.1, 4.9),
        6: (5.8, 1.6),
        7: (7.7, 2.9),
        8: (7.7, 0.6),
        9: (8.7, 4.9),
        10: (10.3, 4.9),
    }
    works = [
        (1, 2, 4, None, '1а'),   # 1a 1->2 dur4
        (1, 3, 6, None, '2а'),   # 2a
        (3, 4, 2, None, '3а'),   # 3a
        (3, 5, 6, None, '4а'),   # 4a
        (6, 7, 3, None, '5а'),   # 5a
        (6, 8, 3, None, '6а'),   # 6a
        (9, 10, 5, None, '7а'),  # 7a
        (2, 6, 0, None, 'Ф'), # фиктивная 2->6
        (4, 6, 0, None, 'Ф'), # фиктивная 4->6
        (5, 9, 0, None, 'Ф'), # фиктивная 5->9
        (8, 9, 0, None, 'Ф'), # фиктивная 8->9
    ]
    fig = draw_network(pos, works, width=9.6, height=5.4)
    save_net(fig, '/home/user/Verich/report_assets/net_5_1.png')

# ---------------- 5.3(в) ----------------
def chart_5_3():
    pos = {
        1: (0.5, 4.0),
        2: (2.4, 5.7),
        3: (2.4, 2.3),
        4: (4.6, 4.0),
        5: (6.6, 5.5),
        6: (8.6, 4.0),
    }
    works = [
        (1, 2, 6, None, '(1,2) 6'),   # 1a
        (1, 3, 13, None, '(1,3) 13'), # 2a
        (1, 4, 20, None, '(1,4) 20'), # 3a
        (2, 4, 9, None, '(2,4) 9'),   # 4a
        (2, 5, 14, None, '(2,5) 14'), # 5a
        (3, 4, 16, None, '(3,4) 16'), # 6a
        (3, 6, 15, None, '(3,6) 15'), # 7a
        (4, 5, 10, None, '(4,5) 10'), # 8a
        (4, 6, 17, None, '(4,6) 17'), # 9a
        (5, 6, 13, None, '(5,6) 13'), # 10a
    ]
    fig = draw_network(pos, works, width=7.4, height=5.0)
    save_net(fig, '/home/user/Verich/report_assets/net_5_3.png')

# ---------------- 5.5(в) ----------------
def chart_5_5():
    pos = {
        1: (0.6, 3.4),
        2: (2.6, 1.7),
        3: (4.7, 3.6),
        4: (6.6, 5.2),
        5: (8.3, 3.6),
    }
    works = [
        (1, 2, 15, None, '(1,2) 15'),
        (1, 3, 8, None, '(1,3) 8'),
        (1, 4, 7, None, '(1,4) 7'),
        (2, 3, 5, None, '(2,3) 5'),
        (3, 4, 13, None, '(3,4) 13'),
        (3, 5, 11, None, '(3,5) 11'),
        (4, 5, 7, None, '(4,5) 7'),
    ]
    fig = draw_network(pos, works, width=7.4, height=4.8)
    save_net(fig, '/home/user/Verich/report_assets/net_5_5.png')

# ---------------- 5.8(в) ----------------
def chart_5_8():
    pos = {
        1: (0.6, 3.6),
        2: (2.6, 4.9),
        3: (2.6, 2.3),
        4: (5.8, 5.5),
        5: (6.9, 3.2),
    }
    works = [
        (1, 2, 6, None, '(1,2) 6'),
        (1, 3, 11, None, '(1,3) 11'),
        (2, 3, 4, None, '(2,3) 4'),
        (2, 4, 7, None, '(2,4) 7'),
        (2, 5, 9, None, '(2,5) 9'),
        (3, 5, 3, None, '(3,5) 3'),
        (4, 5, 8, None, '(4,5) 8'),
    ]
    fig = draw_network(pos, works, width=7.0, height=4.6)
    save_net(fig, '/home/user/Verich/report_assets/net_5_8.png')

# ---------------- 5.16(в) ----------------
def chart_5_16():
    pos = {
        1: (0.6, 3.8),
        2: (2.6, 5.7),
        3: (2.6, 1.9),
        4: (5.4, 3.7),
        5: (8.2, 3.8),
    }
    works = [
        (1, 2, 4, 8, None),   # 1a
        (1, 3, 6, 5, None),   # 2a
        (1, 4, 8, 9, None),   # 3a
        (2, 4, 2, 10, None),  # 4a
        (2, 5, 5, 4, None),   # 5a
        (3, 4, 2, 1, None),   # 6a
        (3, 5, 3, 2, None),   # 7a
        (4, 5, 4, 12, None),  # 8a
    ]
    fig = draw_network(pos, works, width=7.4, height=4.7)
    save_net(fig, '/home/user/Verich/report_assets/net_5_16.png')

# ---------------- 5.1(в) Gantt ----------------
def chart_gantt_5_1():
    # works: (label, start, dur, is_fict)
    works = [
        ('1а', 0, 4, False),
        ('2а', 0, 6, False),
        ('3а', 6, 2, False),
        ('4а', 6, 6, False),
        ('5а', 8, 3, False),
        ('6а', 8, 3, False),
        ('7а', 12, 5, False),
        ('Ф(2,6)', 4, 4, True),
        ('Ф(4,6)', 8, 0, True),
        ('Ф(5,9)', 12, 0, True),
        ('Ф(8,9)', 11, 1, True),
    ]
    fig, ax = plt.subplots(figsize=(10.5, 6.4))
    ax.set_xlim(-0.6, 18)
    ax.set_ylim(-0.6, len(works) + 0.4)
    # time axis
    ax.set_xticks(range(0, 18, 1))
    ax.set_xticklabels([str(i) for i in range(0, 18, 1)], fontsize=9)
    ax.set_yticks(range(len(works)))
    ax.set_yticklabels([w[0] for w in works][::-1], fontsize=12, family='DejaVu Serif')
    ax.grid(axis='x', color='0.82', linewidth=0.6, linestyle=':')
    for i, (label, start, dur, is_fict) in enumerate(works):
        y = len(works) - 1 - i
        if is_fict:
            if dur > 0:
                ax.plot([start, start + dur], [y, y], color='black', linewidth=1.2, linestyle='--', zorder=5)
                ax.plot(start, y, marker='o', markersize=4, color='black', zorder=6)
                ax.plot(start + dur, y, marker='o', markersize=4, color='black', zorder=6)
            else:
                ax.plot(start, y, marker='o', markersize=6, color='black', zorder=6)
        else:
            ax.barh(y, dur, left=start, height=0.52, color='black', edgecolor='black', zorder=5)
            ax.text(start + dur / 2, y, str(dur), va='center', ha='center', color='white', fontsize=11, zorder=6, family='DejaVu Serif', fontweight='bold')
    ax.set_xlabel('t', fontsize=14, family='DejaVu Serif')
    ax.tick_params(axis='y', length=0)
    for s in ['top', 'right', 'left']:
        ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_linewidth(1.2)
    ax.set_title('Линейный график проекта (рис. 1)', fontsize=15, family='DejaVu Serif', fontweight='bold')
    fig.tight_layout()
    fig.savefig('/home/user/Verich/report_assets/gantt_5_1.png', dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)

# ---------------- 5.16(в) Gantt ----------------
def chart_gantt_5_16(schedule, out_path, title):
    # schedule: list of (label, [intervals], dur, r)
    works = schedule
    fig, ax = plt.subplots(figsize=(11.5, 4.6 + len(works) * 0.34))
    ax.set_xlim(-0.5, 17.2)
    ax.set_ylim(-0.7, len(works) + 0.35)
    ax.set_xticks(range(0, 18, 1))
    ax.set_xticklabels([str(i) for i in range(0, 18, 1)], fontsize=10)
    ax.set_yticks(range(len(works)))
    ax.set_yticklabels([w[0] for w in works][::-1], fontsize=12.5, family='DejaVu Serif')
    ax.grid(axis='x', color='0.82', linewidth=0.6, linestyle=':')
    for i, (label, intervals, dur, r) in enumerate(works):
        y = len(works) - 1 - i
        for a, b in intervals:
            ax.barh(y, b - a, left=a, height=0.52, color='black', edgecolor='black', zorder=5)
        ax.text(-0.25, y, f'{r}', ha='right', va='center', fontsize=11, family='DejaVu Serif', style='italic', color='black')
    ax.set_xlabel('t', fontsize=14, family='DejaVu Serif')
    ax.tick_params(axis='y', length=0)
    for s in ['top', 'right', 'left']:
        ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_linewidth(1.2)
    ax.set_title(title, fontsize=15, family='DejaVu Serif', fontweight='bold')
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)

if __name__ == '__main__':
    chart_5_1()
    chart_5_3()
    chart_5_5()
    chart_5_8()
    chart_5_16()
    chart_gantt_5_1()
    # 5.16(в) non-interruptible schedule (T=16)
    sched_a = [
        ('(1,2)', [(0, 4)], 4, 8),
        ('(1,3)', [(4, 10)], 6, 5),
        ('(1,4)', [(0, 8)], 8, 9),
        ('(2,4)', [(10, 12)], 2, 10),
        ('(2,5)', [(8, 13)], 5, 4),
        ('(3,4)', [(10, 12)], 2, 1),
        ('(3,5)', [(11, 14)], 3, 2),
        ('(4,5)', [(12, 16)], 4, 12),
    ]
    chart_gantt_5_16(sched_a, '/home/user/Verich/report_assets/gantt_5_16a.png', 'Линейный график проекта (а) – работы не допускают перерыва (T=16)')
    # 5.16(в) interruptible schedule (T=15)
    sched_b = [
        ('(1,2)', [(0, 4)], 4, 8),
        ('(1,3)', [(2, 3), (4, 9)], 6, 5),
        ('(1,4)', [(0, 2), (3, 9)], 8, 9),
        ('(2,4)', [(9, 11)], 2, 10),
        ('(2,5)', [(10, 15)], 5, 4),
        ('(3,4)', [(9, 11)], 2, 1),
        ('(3,5)', [(12, 15)], 3, 2),
        ('(4,5)', [(11, 15)], 4, 12),
    ]
    chart_gantt_5_16(sched_b, '/home/user/Verich/report_assets/gantt_5_16b.png', 'Линейный график проекта (б) – работы допускают разрыв (T=15)')
    print('all charts done')
