"""Render Excel sheets as black-and-white 'screenshot' PNGs for the appendix."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import openpyxl
from openpyxl.utils import range_boundaries

wb = openpyxl.load_workbook('outputs/Эконометрика УСР 5.xlsx')

def col_letter(i):
    s = ''
    while i > 0:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s

for ws in wb.worksheets:
    max_col = ws.max_column
    max_row = ws.max_row

    # collect merged ranges
    merged = []
    for mr in ws.merged_cells.ranges:
        min_col, min_row, max_col_r, max_row_r = range_boundaries(str(mr))
        merged.append((min_row, min_col, max_row_r, max_col_r))

    def covered_by_merge(r, c):
        for (r1, c1, r2, c2) in merged:
            if r1 <= r <= r2 and c1 <= c <= c2:
                return (r1, c1, r2, c2)
        return None

    # column widths
    col_w = []
    for c in range(1, max_col + 1):
        w = 6
        for r in range(1, max_row + 1):
            if covered_by_merge(r, c):
                continue
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() != '':
                w = max(w, min(len(str(v)) * 1.35, 20))
        col_w.append(w)
    total_w = sum(col_w)
    col_w = [w / total_w * 100 for w in col_w]

    num_w = 4.0
    col_x = num_w
    col_x_list = []
    for w in col_w:
        col_x_list.append((col_x, col_x + w))
        col_x += w

    row_h = 0.46
    n_rows = max_row
    header_h = 0.46
    bar_h = 0.78
    top_offset = bar_h + 0.16 + header_h + 0.02
    fig_h = 0.5 + top_offset + n_rows * row_h
    fig_w = 11.5
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, fig_h)
    ax.axis('off')
    ax.set_facecolor('white')
    plt.subplots_adjust(left=0.02, right=0.98, top=0.99, bottom=0.02)

    bar_y_top = fig_h - 0.2
    bar_y = bar_y_top - bar_h
    ax.add_patch(Rectangle((0.3, bar_y), 99.4, bar_h, facecolor='none', edgecolor='black', lw=1.4))
    ax.text(1.2, bar_y + bar_h / 2, f'Эконометрика УСР 5  ·  лист «{ws.title}»',
            fontsize=12, fontweight='bold', va='center')

    hdr_y_top = bar_y - 0.16
    hdr_y = hdr_y_top - header_h
    ax.add_patch(Rectangle((0.3, hdr_y), 99.4, header_h, facecolor='#e8e8e8', edgecolor='black', lw=1.0))
    ax.add_patch(Rectangle((0.3, hdr_y), num_w - 0.3, header_h, facecolor='#e8e8e8', edgecolor='black', lw=1.0))
    for i, (x1, x2) in enumerate(col_x_list, start=1):
        ax.text((x1 + x2) / 2, hdr_y + header_h / 2, col_letter(i), ha='center', va='center', fontsize=8, fontweight='bold')

    # draw rows
    for r in range(1, max_row + 1):
        top_y = hdr_y - 0.02 - (r - 1) * row_h
        y_bot = top_y - row_h + 0.02
        ax.text(num_w / 2, top_y - row_h / 2, str(r), ha='center', va='center', fontsize=7.5)
        ax.add_patch(Rectangle((0.3, y_bot), num_w - 0.3, row_h - 0.02, facecolor='#eaeaea', edgecolor='black', lw=0.7))

        # find merges that start at this row
        starts = [mr for mr in merged if mr[0] == r]
        covered_cells = set()
        for (r1, c1, r2, c2) in merged:
            if r1 == r:  # merge starting here
                # merged rect spanning rows r1..r2 and cols c1..c2
                pass
        # draw non-merged cells and merge rects
        drawn = set()
        for (r1, c1, r2, c2) in merged:
            if r1 == r:
                v = ws.cell(row=r1, column=c1).value
                x1, x2 = col_x_list[c1 - 1][0], col_x_list[c2 - 1][1]
                top_y2 = hdr_y - 0.02 - (r1 - 1) * row_h
                y_bot2 = top_y2 - (row_h * (r2 - r1 + 1)) + 0.02
                ax.add_patch(Rectangle((x1 + 0.06, y_bot2), x2 - x1 - 0.12, top_y2 - y_bot2,
                                          facecolor='white', edgecolor='black', lw=0.7))
                txt = str(v) if v is not None else ''
                if txt.strip():
                    w = x2 - x1
                    fontsize = 6.6 if len(txt) <= w * 0.55 else 6.0
                    ax.text(x1 + 0.5, (top_y2 + y_bot2) / 2, txt, ha='left', va='center', fontsize=fontsize)
                for rr in range(r1, r2 + 1):
                    for cc in range(c1, c2 + 1):
                        drawn.add((rr, cc))
        for c in range(1, max_col + 1):
            if (r, c) in drawn:
                continue
            if covered_by_merge(r, c):
                # part of a merge that started earlier; skip
                if merged:
                    mr = covered_by_merge(r, c)
                    if mr[0] != r:
                        continue
            x1, x2 = col_x_list[c - 1]
            v = ws.cell(row=r, column=c).value
            w = x2 - x1
            txt = str(v) if v is not None else ''
            ax.add_patch(Rectangle((x1 + 0.06, y_bot + 0.02), w - 0.12, row_h - 0.06,
                         facecolor='white', edgecolor='black', lw=0.6))
            is_num = False
            try:
                float(txt.replace(',', '.'))
                is_num = True
            except (ValueError, AttributeError):
                is_num = False
            if is_num:
                ax.text(x2 - 0.5, y_bot + row_h / 2, txt, ha='right', va='center', fontsize=7)
            elif len(txt) <= 2:
                ax.text((x1 + x2) / 2, y_bot + row_h / 2, txt, ha='center', va='center', fontsize=7)
            else:
                maxlen = max(int(w * 0.62), 8)
                disp = txt if len(txt) <= maxlen else txt[:maxlen - 1] + '…'
                ax.text(x1 + 0.45, y_bot + row_h / 2, disp, ha='left', va='center', fontsize=6.5)

    bottom_y = hdr_y - 0.02 - n_rows * row_h
    ax.add_patch(Rectangle((0.3, bottom_y), 99.4, bar_y_top - bottom_y, facecolor='none', edgecolor='black', lw=2.2))

    out = ws.title.replace('/', '_').replace('а', 'a').replace('б', 'b')
    path = f'/home/user/Verich/report_assets/excel_{out}.png'
    fig.savefig(path, dpi=130, facecolor='white', bbox_inches='tight', pad_inches=0.10)
    plt.close(fig)
    print('saved', path)
