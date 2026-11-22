"""Render Лист2 of the УСР Свитина workbook as black-and-white 'screenshot' PNGs."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import openpyxl
from openpyxl.utils import range_boundaries

XLSX = '/home/user/Verich/Пась А.М. 24ИТ-3 УСР Свитина.xlsx'
OUT = '/home/user/Verich/report_assets/'

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb['Лист2']

COL_WIDTHS = dict(wb['Лист2'].column_dimensions)

def get_w(c):
    letter = ''
    i = c
    while i:
        i, r = divmod(i - 1, 26)
        letter = chr(65 + r) + letter
    d = COL_WIDTHS.get(letter)
    return d.width if d is not None and d.customWidth else 9.5

def col_letter(i):
    s = ''
    while i > 0:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s

def render_range(c1, c2, name, filename):
    max_row = ws.max_row

    merged = []
    for mr in ws.merged_cells.ranges:
        min_col, min_row, max_col_r, max_row_r = range_boundaries(str(mr))
        if max_col_r >= c1 and min_col <= c2:
            lo = max(min_col, c1)
            hi = min(max_col_r, c2)
            merged.append((min_row, lo, max_row_r, hi))

    col_w = []
    for c in range(c1, c2 + 1):
        col_w.append(get_w(c))
    total_w = sum(col_w)
    col_w = [w / total_w * 100 for w in col_w]

    num_w = 3.6
    col_x = num_w
    col_x_list = []
    for w in col_w:
        col_x_list.append((col_x, col_x + w))
        col_x += w

    row_h = 0.40
    header_h = 0.42
    bar_h = 0.70
    top_offset = bar_h + 0.12 + header_h + 0.02
    fig_h = 0.4 + top_offset + max_row * row_h
    fig_w = 11.8
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(0, 100)
    ax.set_ylim(0, fig_h)
    ax.axis('off')
    ax.set_facecolor('white')
    plt.subplots_adjust(left=0.015, right=0.985, top=0.99, bottom=0.02)

    bar_y_top = fig_h - 0.14
    bar_y = bar_y_top - bar_h
    ax.add_patch(Rectangle((0.3, bar_y), 99.4, bar_h, facecolor='none', edgecolor='black', lw=1.4))
    ax.text(1.2, bar_y + bar_h / 2, 'Эконометрика УСР 1  ·  лист «Лист2»  ·  ' + name,
            fontsize=11, fontweight='bold', va='center')

    hdr_y_top = bar_y - 0.12
    hdr_y = hdr_y_top - header_h
    ax.add_patch(Rectangle((0.3, hdr_y), 99.4, header_h, facecolor='#e8e8e8', edgecolor='black', lw=1.0))
    ax.add_patch(Rectangle((0.3, hdr_y), num_w - 0.3, header_h, facecolor='#d8d8d8', edgecolor='black', lw=1.0))
    for i, (x1, x2) in enumerate(col_x_list, start=c1):
        ax.text((x1 + x2) / 2, hdr_y + header_h / 2, col_letter(i), ha='center', va='center', fontsize=7, fontweight='bold')

    for (r1, cc1, r2, cc2) in merged:
        v = ws.cell(row=r1, column=cc1).value
        x1 = col_x_list[cc1 - c1][0]
        x2 = col_x_list[cc2 - c1][1]
        top_y2 = hdr_y - 0.02 - (r1 - 1) * row_h
        y_bot2 = top_y2 - (row_h * (r2 - r1 + 1)) + 0.02
        ax.add_patch(Rectangle((x1 + 0.05, y_bot2), x2 - x1 - 0.10, top_y2 - y_bot2,
                               facecolor='#f2f2f2', edgecolor='black', lw=0.9))
        txt = str(v) if v is not None else ''
        if txt.strip():
            w = x2 - x1
            fontsize = min(6.6, max(4.6, w * 0.45))
            ax.text(x1 + (x2 - x1) / 2, (top_y2 + y_bot2) / 2, txt.strip(), ha='center', va='center', fontsize=fontsize)

    for r in range(1, max_row + 1):
        top_y = hdr_y - 0.02 - (r - 1) * row_h
        y_bot = top_y - row_h + 0.02
        ax.text(num_w / 2, top_y - row_h / 2, str(r), ha='center', va='center', fontsize=6.8)
        ax.add_patch(Rectangle((0.3, y_bot), num_w - 0.3, row_h - 0.02, facecolor='#eaeaea', edgecolor='black', lw=0.7))
        for c in range(c1, c2 + 1):
            in_merge = False
            for (r1, cc1, r2, cc2) in merged:
                if r1 <= r <= r2 and cc1 <= c <= cc2:
                    in_merge = True
                    break
            if in_merge:
                continue
            x1, x2 = col_x_list[c - c1]
            v = ws.cell(row=r, column=c).value
            w = x2 - x1
            txt = (str(v) if v is not None else '').strip()
            ax.add_patch(Rectangle((x1 + 0.05, y_bot + 0.02), w - 0.10, row_h - 0.06,
                                  facecolor='white', edgecolor='black', lw=0.6))
            is_num = False
            try:
                float(txt.replace(',', '.'))
                is_num = True
            except (ValueError, AttributeError):
                is_num = False
            max_chars = max(int(w * 0.92), 6)
            if len(txt) > max_chars:
                disp = txt[:max_chars - 1] + '…'
            else:
                disp = txt
            fontsize = 6.4 if max_chars >= 12 else 5.6
            if is_num:
                ax.text(x2 - 0.45, y_bot + row_h / 2, disp, ha='right', va='center', fontsize=fontsize)
            elif len(txt) <= 3:
                ax.text((x1 + x2) / 2, y_bot + row_h / 2, disp, ha='center', va='center', fontsize=fontsize)
            else:
                ax.text(x1 + 0.4, y_bot + row_h / 2, disp, ha='left', va='center', fontsize=fontsize)

    bottom_y = hdr_y - 0.02 - max_row * row_h
    ax.add_patch(Rectangle((0.3, bottom_y), 99.4, bar_y_top - bottom_y, facecolor='none', edgecolor='black', lw=2.0))

    path = OUT + filename
    fig.savefig(path, dpi=150, facecolor='white', bbox_inches='tight', pad_inches=0.07)
    plt.close(fig)
    print('saved', path)

render_range(1, 15, 'схемы МОБ', 'excel_mobr_2a.png')
render_range(18, 23, 'пункты 4 и 5', 'excel_mobr_2b.png')
