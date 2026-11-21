"""Generate network diagrams and Gantt charts (color)."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Circle
import math

plt.rcParams['font.family'] = 'DejaVu Serif'
plt.rcParams['font.size'] = 13
plt.rcParams['axes.linewidth'] = 1.0

NAVY = '#1d3557'
BLUE = '#457b9d'
RED = '#d62828'
GRAY = '#6c757d'
LIGHT_GRAY = '#adb5bd'
GOLD = '#e9c46a'

def draw_network(pos, works, width=7.2, height=4.6, dpi=200, critical=(), legend=True):
    """works: list of (i, j, t_or_r, r_or_None, label_or_None, is_critical_or_None)"""
    fig, ax = plt.subplots(figsize=(width, height))
    ax.set_xlim(-0.5, 11.2)
    ax.set_ylim(0.0, 6.3)
    ax.axis('off')
    # events
    for e, (x, y) in pos.items():
        c = Circle((x, y), 0.26, fill=True, facecolor='white', edgecolor=NAVY, linewidth=1.8, zorder=6)
        ax.add_patch(c)
        ax.text(x, y, str(e), ha='center', va='center', fontsize=14, zorder=7, family='DejaVu Serif', fontweight='bold', color=NAVY)
    # works
    for w in works:
        i, j, t, r = w[0], w[1], w[2], w[3]
        xi, yi = pos[i]; xj, yj = pos[j]
        is_fict = (t == 0) and (len(w) > 4 and w[4] == 'Ф')
        # 5.1 works use tuples with (i, j, t, None, label, fict) with r=None for real works
        label = w[4] if len(w) > 4 else None
        is_crit = len(w) > 5 and w[5]
        dx = xj - xi; dy = yj - yi
        L = math.hypot(dx, dy)
        ux, uy = dx / L, dy / L
        start = (xi + ux * 0.28, yi + uy * 0.28)
        end = (xj - ux * 0.28, yj - uy * 0.28)
        color = RED if is_crit else (GRAY if is_fict else BLUE)
        ax.annotate('', xy=end, xytext=start,
                    arrowprops=dict(arrowstyle='-|>', lw=2.2 if is_crit else 1.5,
                                    color=color, linestyle='--' if is_fict else '-',
                                    mutation_scale=16), zorder=4)
        if is_fict:
            ax.text((xi+xj)/2, (yi+yj)/2 + 0.06*(-uy), 'Ф', ha='center', va='center', fontsize=11, style='italic', color=GRAY, zorder=5)
        else:
            disp = f'{label}' if label is not None else (f'{t},{r}' if r is not None else f'{t}')
            offx, offy = -uy * 0.24, ux * 0.24
            ax.text((xi+xj)/2 + offx, (yi+yj)/2 + offy, disp, ha='center', va='center', fontsize=12, zorder=5,
                    bbox=dict(fc='white', ec='none', pad=1.4), color=RED if is_crit else NAVY,
                    fontweight='bold' if is_crit else 'normal')
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    handles = [
        Line2D([0], [0], color=BLUE, linewidth=2, marker='>', markersize=7, markerfacecolor=BLUE, label='Работы'),
        Line2D([0], [0], color=RED, linewidth=2.2, marker='>', markersize=7, markerfacecolor=RED, label='Критические/выделенные работы'),
        Line2D([0], [0], color=GRAY, linestyle='--', linewidth=1.5, label='Фиктивные работы'),
    ]
    if legend:
        ax.legend(handles=handles, loc='upper center', bbox_to_anchor=(0.5, -0.03),
                  frameon=False, fontsize=10, ncol=3, handlelength=2.4)
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
    # works: (i, j, t, r, label, is_critical)
    works = [
        (1, 2, 4, None, '1а  4', False),
        (1, 3, 6, None, '2а  6', True),     # critical
        (3, 4, 2, None, '3а  2', False),
        (3, 5, 6, None, '4а  6', True),     # critical
        (6, 7, 3, None, '5а  3', False),
        (6, 8, 3, None, '6а  3', False),
        (9, 10, 5, None, '7а  5', True),    # critical
        (2, 6, 0, None, 'Ф', False),        # фиктивная
        (4, 6, 0, None, 'Ф', False),
        (5, 9, 0, None, 'Ф', True),         # critical dummy
        (8, 9, 0, None, 'Ф', False),
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
    # works: (i, j, t, r, label, is_critical=True for works with money invested)
    works = [
        (1, 2, 6, None, '(1,2)  6', False),
        (1, 3, 13, None, '(1,3)  13', True),    # invested
        (1, 4, 20, None, '(1,4)  20', False),
        (2, 4, 9, None, '(2,4)  9', False),
        (2, 5, 14, None, '(2,5)  14', False),
        (3, 4, 16, None, '(3,4)  16', True),    # invested
        (3, 6, 15, None, '(3,6)  15', False),
        (4, 5, 10, None, '(4,5)  10', True),    # invested
        (4, 6, 17, None, '(4,6)  17', False),
        (5, 6, 13, None, '(5,6)  13', True),    # invested
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
        (1, 2, 15, None, '(1,2)  15', True),    # invested
        (1, 3, 8, None, '(1,3)  8', False),
        (1, 4, 7, None, '(1,4)  7', False),
        (2, 3, 5, None, '(2,3)  5', True),      # invested
        (3, 4, 13, None, '(3,4)  13', True),    # invested
        (3, 5, 11, None, '(3,5)  11', False),
        (4, 5, 7, None, '(4,5)  7', True),      # invested
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
    # final critical path: (1,2), (2,4), (4,5)
    works = [
        (1, 2, 6, None, '(1,2)  6', True),   # critical
        (1, 3, 11, None, '(1,3)  11', False),
        (2, 3, 4, None, '(2,3)  4', False),
        (2, 4, 7, None, '(2,4)  7', True),   # critical
        (2, 5, 9, None, '(2,5)  9', False),
        (3, 5, 3, None, '(3,5)  3', False),
        (4, 5, 8, None, '(4,5)  8', True),   # critical
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
        (1, 2, 4, 8, None, False),   # 1a
        (1, 3, 6, 5, None, False),
        (1, 4, 8, 9, None, False),
        (2, 4, 2, 10, None, False),
        (2, 5, 5, 4, None, False),
        (3, 4, 2, 1, None, False),
        (3, 5, 3, 2, None, False),
        (4, 5, 4, 12, None, False),
    ]
    fig = draw_network(pos, works, width=7.4, height=4.7, legend=False)
    save_net(fig, '/home/user/Verich/report_assets/net_5_16.png')

# ---------------- 5.1(в) Gantt ----------------
def chart_gantt_5_1():
    # works: (label, start, dur, is_fict, is_critical)
    works = [
        ('1а', 0, 4, False, False),
        ('2а', 0, 6, False, True),    # critical
        ('3а', 6, 2, False, False),
        ('4а', 6, 6, False, True),    # critical
        ('5а', 8, 3, False, False),
        ('6а', 8, 3, False, False),
        ('7а', 12, 5, False, True),   # critical
        ('Ф(2,6)', 4, 4, True, False),
        ('Ф(4,6)', 8, 0, True, False),
        ('Ф(5,9)', 12, 0, True, True),   # critical dummy
        ('Ф(8,9)', 11, 1, True, False),
    ]
    fig, ax = plt.subplots(figsize=(10.5, 6.4))
    ax.set_xlim(-0.6, 18)
    ax.set_ylim(-0.6, len(works) + 0.4)
    ax.set_xticks(range(0, 18, 1))
    ax.set_xticklabels([str(i) for i in range(0, 18, 1)], fontsize=9)
    ax.set_yticks(range(len(works)))
    ax.set_yticklabels([w[0] for w in works][::-1], fontsize=12, family='DejaVu Serif')
    ax.grid(axis='x', color='0.82', linewidth=0.6, linestyle=':')
    for i, (label, start, dur, is_fict, is_crit) in enumerate(works):
        y = len(works) - 1 - i
        color = RED if is_crit else BLUE
        if is_fict:
            if dur > 0:
                ax.plot([start, start + dur], [y, y], color=color, linewidth=1.6, linestyle='--', zorder=5)
                ax.plot(start, y, marker='o', markersize=4, color=color, zorder=6)
                ax.plot(start + dur, y, marker='o', markersize=4, color=color, zorder=6)
            else:
                ax.plot(start, y, marker='o', markersize=6, color=color, zorder=6)
        else:
            ax.barh(y, dur, left=start, height=0.52, color=color, edgecolor='black', linewidth=0.8, zorder=5)
            ax.text(start + dur / 2, y, str(dur), va='center', ha='center', color='white', fontsize=11, zorder=6, family='DejaVu Serif', fontweight='bold')
    ax.set_xlabel('t', fontsize=14, family='DejaVu Serif')
    ax.tick_params(axis='y', length=0)
    for s in ['top', 'right', 'left']:
        ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_linewidth(1.2)
    ax.spines['bottom'].set_color(NAVY)
    from matplotlib.patches import Patch
    from matplotlib.lines import Line2D
    legend_handles = [
        Patch(facecolor=RED, edgecolor='black', linewidth=0.8, label='Критический путь'),
        Patch(facecolor=BLUE, edgecolor='black', linewidth=0.8, label='Работы'),
        Line2D([0], [0], color=GRAY, linestyle='--', linewidth=1.6, label='Фиктивные работы'),
    ]
    ax.legend(handles=legend_handles, loc='upper left', bbox_to_anchor=(0.0, 1.02),
              frameon=False, fontsize=10, ncol=3)
    ax.set_title('Линейный график проекта (рис. 1)', fontsize=15, family='DejaVu Serif', fontweight='bold', color=NAVY, pad=22)
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
        # color by rate r: two-tone tiers (blue -> gold -> red)
        if r <= 3:
            color = BLUE
        elif r <= 8:
            color = GOLD
        else:
            color = RED
        for a, b in intervals:
            ax.barh(y, b - a, left=a, height=0.52, color=color, edgecolor='black', linewidth=0.8, zorder=5)
        ax.text(-0.25, y, f'{r}', ha='right', va='center', fontsize=11, family='DejaVu Serif', style='italic', color=NAVY)
    ax.set_xlabel('t', fontsize=14, family='DejaVu Serif')
    ax.tick_params(axis='y', length=0)
    for s in ['top', 'right', 'left']:
        ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_linewidth(1.2)
    ax.spines['bottom'].set_color(NAVY)
    from matplotlib.patches import Patch
    legend_handles = [
        Patch(facecolor=BLUE, edgecolor='black', linewidth=0.8, label='r ≤ 3'),
        Patch(facecolor=GOLD, edgecolor='black', linewidth=0.8, label='4 ≤ r ≤ 8'),
        Patch(facecolor=RED, edgecolor='black', linewidth=0.8, label='r ≥ 9'),
    ]
    ax.legend(handles=legend_handles, loc='upper left', bbox_to_anchor=(0.0, 1.04),
              frameon=False, fontsize=10, ncol=3)
    ax.set_title(title, fontsize=15, family='DejaVu Serif', fontweight='bold', color=NAVY, pad=22)
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
