"""Build Excel solution (black-and-white, УСР1-style)."""
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Alignment, Side
from openpyxl.utils import get_column_letter

thin = Side(style='thin', color='000000')
BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)

FONT_TITLE = Font(name='Calibri', size=14, bold=True, color='000000')
FONT_SECT = Font(name='Calibri', size=12, bold=True, color='000000')
FONT_HDR = Font(name='Calibri', size=11, bold=True, color='000000')
FONT_BODY = Font(name='Calibri', size=11, color='000000')
FONT_BOLD = Font(name='Calibri', size=11, bold=True, color='000000')
FONT_SMALL = Font(name='Calibri', size=10, color='000000')

CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT_WRAP = Alignment(horizontal='left', vertical='center', wrap_text=True)

def title_row(ws, row, col1, col2, text, height=36):
    ws.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
    c = ws.cell(row=row, column=col1, value=text)
    c.font = FONT_TITLE; c.alignment = CENTER; c.border = BORDER_ALL
    ws.row_dimensions[row].height = height

def section(ws, row, col1, col2, text):
    ws.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
    c = ws.cell(row=row, column=col1, value=text)
    c.font = FONT_SECT; c.alignment = LEFT_WRAP; c.border = BORDER_ALL
    ws.row_dimensions[row].height = 24

def header_row(ws, row, col_values, height=27):
    for col, value in col_values:
        c = ws.cell(row=row, column=col, value=value)
        c.font = FONT_HDR; c.alignment = CENTER; c.border = BORDER_ALL
    ws.row_dimensions[row].height = height

def cell(ws, row, col, value, bold=False, align=None, num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = FONT_BOLD if bold else FONT_BODY
    c.alignment = align if align else CENTER
    c.border = BORDER_ALL
    if num_fmt: c.number_format = num_fmt
    return c

def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def note_row(ws, row, col1, col2, text, height=28):
    ws.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
    c = ws.cell(row=row, column=col1, value=text)
    c.font = FONT_SMALL; c.alignment = LEFT_WRAP
    ws.row_dimensions[row].height = height

def tab(ws, start_row, title_text, col_values, rows, note=None, note_height=28):
    """Insert a section + header + data rows. Returns next free row."""
    section(ws, start_row, 1, len(col_values), title_text)
    header_row(ws, start_row + 1, col_values)
    r = start_row + 2
    for row_vals, row_fmt in rows:
        for col, val in enumerate(row_vals, start=1):
            cell(ws, r, col, val, num_fmt=row_fmt[col-1] if row_fmt else None)
        ws.row_dimensions[r].height = 21
        r += 1
    if note:
        note_row(ws, r + 1, 1, len(col_values), note, note_height)
        r += 2
    return r

wb = Workbook()
wb.remove(wb.active)

# ================= Sheet 5.1 =================
ws = wb.create_sheet('5.1')
ws.sheet_view.showGridLines = False
set_widths(ws, [13, 26, 15, 15, 15, 15, 15, 15])
title_row(ws, 1, 1, 7, '5.1 (в) — Сетевой график проекта')
r = 3
r = tab(ws, r, 'Исходные данные', [(1,'Работа'),(2,'Предшествующие работы'),(3,'t')], [
    (['1а','—',4], None), (['2а','—',6], None), (['3а','2а',2], None),
    (['4а','2а',6], None), (['5а','1а, 3а',3], None), (['6а','1а, 3а',3], None),
    (['7а','4а, 6а',5], None),
])
r += 1
r = tab(ws, r, 'Параметры событий', [(1,'Событие'),(2,'tп'),(3,'tн'),(4,'R')], [
    ([1,0,0],[None,None,None,'0']), ([2,4,9],[None,None,None,'0']), ([3,6,6],[None,None,None,'0']),
    ([4,8,9],[None,None,None,'0']), ([5,12,12],[None,None,None,'0']), ([6,8,9],[None,None,None,'0']),
    ([7,11,17],[None,None,None,'0']), ([8,11,12],[None,None,None,'0']), ([9,12,12],[None,None,None,'0']),
    ([10,17,17],[None,None,None,'0']),
])
# R formulas
for rr in range(r - 11, r - 1):
    pass
# fill R formulas for event table rows (find the events table rows by content)
ev_hdr_row = None
for rr in range(10, 30):
    v = ws.cell(row=rr, column=1).value
    if v == 'Событие':
        ev_hdr_row = rr
        break
if ev_hdr_row:
    for rr in range(ev_hdr_row + 1, ev_hdr_row + 11):
        ws.cell(row=rr, column=4).value = f'=C{rr}-B{rr}'
        ws.cell(row=rr, column=4).number_format = '0'
        ws.cell(row=rr, column=4).border = BORDER_ALL
        ws.cell(row=rr, column=4).font = FONT_BODY
        ws.cell(row=rr, column=4).alignment = CENTER
r += 1
r = tab(ws, r, 'Параметры работ', [(1,'Работа'),(2,'tп.н.'),(3,'tр.о.'),(4,'tн.н.'),(5,'tн.о.'),(6,'Rп'),(7,'Rс')], [
    (['1а',0,4,5,9,5,0], None), (['2а',0,6,0,6,0,0], None),
    (['3а',6,8,7,9,1,0], None), (['4а',6,12,6,12,0,0], None),
    (['5а',8,11,14,17,6,0], None), (['6а',8,11,9,12,1,0], None),
    (['7а',12,17,12,17,0,0], None),
])
r += 1
section(ws, r, 1, 7, 'Результат')
note_row(ws, r + 1, 1, 7, 'tкр = 17. Критический путь (1→3→5→9→10): работы 2а, 4а, Ф(5,9), 7а. Критические события: 1, 3, 5, 9, 10.', 28)
note_row(ws, r + 2, 1, 7, 'Свободный резерв: Rс = tр(j) — tр(i) — ti,j. Независимый резерв Rн = 0 для всех работ.', 24)

# ================= Sheet 5.3 =================
ws = wb.create_sheet('5.3')
ws.sheet_view.showGridLines = False
set_widths(ws, [12, 26, 26, 26, 24, 24])
title_row(ws, 1, 1, 6, '5.3 (в) — Оптимизация проекта по времени, tкр = 40')
r = 3
r = tab(ws, r, 'Исходные данные', [(1,'Работа'),(2,'t₀'),(3,'d'),(4,'k')], [
    (['(1,2)',6,5,0.05],[None,None,None,'0.00']), (['(1,3)',13,10,0.25],[None,None,None,'0.00']),
    (['(1,4)',20,16,0.30],[None,None,None,'0.00']), (['(2,4)',9,7,0.07],[None,None,None,'0.00']),
    (['(2,5)',14,11,0.15],[None,None,None,'0.00']), (['(3,4)',16,13,0.10],[None,None,None,'0.00']),
    (['(3,6)',15,12,0.05],[None,None,None,'0.00']), (['(4,5)',10,7,0.03],[None,None,None,'0.00']),
    (['(4,6)',17,15,0.14],[None,None,None,'0.00']), (['(5,6)',13,9,0.50],[None,None,None,'0.00']),
])
r += 1
hdr_start = r
r = tab(ws, r, 'Решение (tн = 0 для работ, выходящих из события 1)', [(1,'Работа'),(2,'tн'),(3,'то'),(4,'x'),(5,'новая продолжительность')], [
    (['(1,2)',0.0,6.000,0.000,6.000],None), (['(1,3)',0.0,10.000,12.000,10.000],None),
    (['(1,4)',0.0,20.000,0.000,20.000],None), (['(2,4)',6.0,15.000,0.000,9.000],None),
    (['(2,5)',17.0,31.000,0.000,14.000],None), (['(3,4)',10.0,23.000,30.000,13.000],None),
    (['(3,6)',25.0,40.000,0.000,15.000],None), (['(4,5)',23.0,31.000,66.667,8.000],None),
    (['(4,6)',23.0,40.000,0.000,17.000],None), (['(5,6)',31.0,40.000,8.000,9.000],None),
])
first_data = hdr_start + 2
last_data = first_data + 9
for rr in range(first_data, last_data + 1):
    ws.cell(row=rr, column=4).number_format = '0.000'
    ws.cell(row=rr, column=2).number_format = '0.0'
    ws.cell(row=rr, column=3).number_format = '0.0'
    ws.cell(row=rr, column=5).number_format = '0.0'
sum_r = first_data + 10
cell(ws, sum_r, 1, 'Σx', bold=True)
ws.merge_cells(start_row=sum_r, start_column=1, end_row=sum_r, end_column=3)
cell(ws, sum_r, 4, f'=SUM(D{first_data}:D{last_data})', bold=True, num_fmt='0.000')
cell(ws, sum_r, 5, '')
ws.row_dimensions[sum_r].height = 24
note_row(ws, sum_r + 2, 1, 5, 'Проект выполнен за tкр = 40 (директивный срок). Суммарные дополнительные средства: Σx ≈ 116,67 ден. ед.', 28)
note_row(ws, sum_r + 3, 1, 5, 'tн (начало) — ранний срок начала работы; то (конец) — ранний срок окончания; x — дополнительные средства.', 24)

# ================= Sheet 5.5 =================
ws = wb.create_sheet('5.5')
ws.sheet_view.showGridLines = False
set_widths(ws, [12, 26, 26, 26, 24, 24])
title_row(ws, 1, 1, 5, '5.5 (в) — Оптимизация проекта по времени, B = 47')
r = 3
r = tab(ws, r, 'Исходные данные', [(1,'Работа'),(2,'t₀'),(3,'d'),(4,'k')], [
    (['(1,2)',15,12,0.25],[None,None,None,'0.00']), (['(1,3)',8,5,0.20],[None,None,None,'0.00']),
    (['(1,4)',7,4,0.15],[None,None,None,'0.00']), (['(2,3)',5,3,0.10],[None,None,None,'0.00']),
    (['(3,4)',13,10,0.30],[None,None,None,'0.00']), (['(3,5)',11,8,0.40],[None,None,None,'0.00']),
    (['(4,5)',7,4,0.20],[None,None,None,'0.00']),
])
r += 1
hdr_start = r
r = tab(ws, r, 'Решение (минимизация tкр при Σx ≤ B = 47)', [(1,'Работа'),(2,'tн'),(3,'то'),(4,'x'),(5,'новая продолжительность')], [
    (['(1,2)',0.0,12.000,12.000,12.000],None), (['(1,3)',0.0,8.000,0.000,8.000],None),
    (['(1,4)',0.0,7.000,0.000,7.000],None), (['(2,3)',12.0,16.000,10.000,4.000],None),
    (['(3,4)',16.0,26.000,10.000,10.000],None), (['(3,5)',19.0,30.000,0.000,11.000],None),
    (['(4,5)',26.0,30.000,15.000,4.000],None),
])
first_data = hdr_start + 2
last_data = first_data + 6
for rr in range(first_data, last_data + 1):
    ws.cell(row=rr, column=4).number_format = '0.000'
    ws.cell(row=rr, column=2).number_format = '0.0'
    ws.cell(row=rr, column=3).number_format = '0.0'
    ws.cell(row=rr, column=5).number_format = '0.0'
sum_r = first_data + 7
cell(ws, sum_r, 1, 'Σx', bold=True)
ws.merge_cells(start_row=sum_r, start_column=1, end_row=sum_r, end_column=3)
cell(ws, sum_r, 4, f'=SUM(D{first_data}:D{last_data})', bold=True, num_fmt='0.000')
cell(ws, sum_r, 5, '')
ws.row_dimensions[sum_r].height = 24
note_row(ws, sum_r + 2, 1, 5, 'tкр = 30. Суммарные дополнительные средства: Σx = 47 (= B).', 26)

# ================= Sheet 5.8 =================
ws = wb.create_sheet('5.8')
ws.sheet_view.showGridLines = False
set_widths(ws, [12, 26, 16, 16, 16, 16, 18])
title_row(ws, 1, 1, 7, '5.8 (в) — Сокращение стоимости проекта')
r = 3
hdr_start = r
r = tab(ws, r, 'Исходные данные и КДЗ', [(1,'Работа'),(2,'D'),(3,'d'),(4,'c'),(5,'C'),(6,'КДЗ')], [
    (['(1,2)',6,2,6,14,'=(E{rr}-D{rr})/(B{rr}-C{rr})'],[None,None,None,None,None,None]),
    (['(1,3)',11,6,2,17,None],[None,None,None,None,None,None]),
    (['(2,3)',4,1,4,13,None],[None,None,None,None,None,None]),
    (['(2,4)',7,3,3,19,None],[None,None,None,None,None,None]),
    (['(2,5)',9,4,2,27,None],[None,None,None,None,None,None]),
    (['(3,5)',3,1,5,7,None],[None,None,None,None,None,None]),
    (['(4,5)',8,3,1,31,None],[None,None,None,None,None,None]),
])
# КДЗ formulas: (C-c)/(D-d) with columns: B=D, C=d, D=c, E=C
for rr in range(hdr_start + 2, hdr_start + 9):
    ws.cell(row=rr, column=6).value = f'=(E{rr}-D{rr})/(B{rr}-C{rr})'
    ws.cell(row=rr, column=6).number_format = '0.000'
    ws.cell(row=rr, column=6).border = BORDER_ALL
    ws.cell(row=rr, column=6).font = FONT_BODY
    ws.cell(row=rr, column=6).alignment = CENTER
# fixed cell(1,2) value from first row: set it correctly too
ws.cell(row=hdr_start + 2, column=1).value = '(1,2)'
ws.cell(row=hdr_start + 2, column=6).value = f'=(E{hdr_start + 2}-D{hdr_start + 2})/(B{hdr_start + 2}-C{hdr_start + 2})'
r += 1
hdr2 = r
section(ws, r, 1, 7, 'Ход оптимизации (tкр: 17 → 14 → 13 → 12 → 11 → 8 → 8)')
header_row(ws, r + 1, [(1,'Шаг'),(2,'Сокращённые работы'),(3,'Величина сокращения, дни'),(4,'tкр'),(5,'Стоимость, ден. ед.')], 27)
r = r + 2
trace = [
    (1, '(1,2)',                          4, 17, 31),
    (2, '(2,4)',                          3, 14, 43),
    (3, '(2,4), (3,5)',                   1, 13, 48),
    (4, '(4,5), (3,5)',                   1, 12, 55),
    (5, '(4,5), (1,3)',                   1, 11, 64),
    (6, '(4,5), (2,5), (1,3)',            3,  8, 106),
    (7, '(2,5), (1,3)',                   1,  8, 114),
]
first_trace_row = hdr2 + 2
for i, (step, works, amount, t_kr, cost) in enumerate(trace):
    rr = first_trace_row + i
    ws.row_dimensions[rr].height = 21
    cell(ws, rr, 1, step)
    cell(ws, rr, 2, works)
    cell(ws, rr, 3, amount, bold=True)
    cell(ws, rr, 4, t_kr)
    cell(ws, rr, 5, cost)
r_note = first_trace_row + 8
section(ws, r_note, 1, 7, 'Результат (tкр = 8, C = 114, нормальный C = 23)')
note_row(ws, r_note + 1, 1, 7, 'Критический путь: (1,2) → (2,4) → (4,5). Итоговые D/d: (1,2)=6→2, (1,3)=11→6, (2,3)=4→1, (2,4)=7→3, (2,5)=9→4, (3,5)=3→1, (4,5)=8→3.', 30)

# ================= Sheet 5.16а =================
ws = wb.create_sheet('5.16а')
ws.sheet_view.showGridLines = False
set_widths(ws, [13, 16, 16, 22, 22, 22, 22, 22])
title_row(ws, 1, 1, 7, '5.16 (в) — Оптимизация по ресурсам (а), R = 18, не допускают перерыва')
r = 3
r = tab(ws, r, 'Исходные данные', [(1,'Работа'),(2,'t'),(3,'r')], [
    (['(1,2)',4,8],None), (['(1,3)',6,5],None), (['(1,4)',8,9],None), (['(2,4)',2,10],None),
    (['(2,5)',5,4],None), (['(3,4)',2,1],None), (['(3,5)',3,2],None), (['(4,5)',4,12],None),
])
r += 1
hdr_start = r
r = tab(ws, r, 'Оптимальное расписание (tн — ранний срок начала, то — ранний срок окончания)', [(1,'Работа'),(2,'tн'),(3,'то'),(4,'t'),(5,'r')], [
    (['(1,2)',0,4,4,8],None), (['(1,3)',4,10,6,5],None), (['(1,4)',0,8,8,9],None),
    (['(2,4)',10,12,2,10],None), (['(2,5)',8,13,5,4],None), (['(3,4)',10,12,2,1],None),
    (['(3,5)',11,14,3,2],None), (['(4,5)',12,16,4,12],None),
])
r += 1
hdr_start2 = r
r = tab(ws, r, 'Эпюра потребности в ресурсе (t = 0..15)', [(1,'Период'),(2,'r')], [
    ([0,17],None), ([1,17],None), ([2,17],None), ([3,17],None),
    ([4,14],None), ([5,14],None), ([6,14],None), ([7,14],None),
    ([8,9],None), ([9,9],None), ([10,15],None), ([11,17],None),
    ([12,18],None), ([13,14],None), ([14,12],None), ([15,12],None),
], note='max r = 18 (= R). tкр = 16.')
# merge r+1 note
# ================= Sheet 5.16б =================
ws = wb.create_sheet('5.16б')
ws.sheet_view.showGridLines = False
set_widths(ws, [13, 16, 16, 24, 24])
title_row(ws, 1, 1, 5, '5.16 (в) — Оптимизация по ресурсам (б), R = 18, допускают разрыв')
r = 3
r = tab(ws, r, 'Исходные данные', [(1,'Работа'),(2,'t'),(3,'r')], [
    (['(1,2)',4,8],None), (['(1,3)',6,5],None), (['(1,4)',8,9],None), (['(2,4)',2,10],None),
    (['(2,5)',5,4],None), (['(3,4)',2,1],None), (['(3,5)',3,2],None), (['(4,5)',4,12],None),
])
r += 1
r = tab(ws, r, 'Оптимальное расписание (части работы в интервалах)', [(1,'Работа'),(2,'Интервалы'),(3,'t'),(4,'r')], [
    (['(1,2)','[0–4)',4,8],None), (['(1,3)','[2–3), [4–9)',6,5],None),
    (['(1,4)','[0–2), [3–9)',8,9],None), (['(2,4)','[9–11)',2,10],None),
    (['(2,5)','[10–15)',5,4],None), (['(3,4)','[9–11)',2,1],None),
    (['(3,5)','[12–15)',3,2],None), (['(4,5)','[11–15)',4,12],None),
])
r += 1
r = tab(ws, r, 'Эпюра потребности в ресурсе (t = 0..14)', [(1,'Период'),(2,'r')], [
    ([0,17],None), ([1,17],None), ([2,13],None), ([3,17],None),
    ([4,14],None), ([5,14],None), ([6,14],None), ([7,14],None),
    ([8,14],None), ([9,11],None), ([10,15],None), ([11,16],None),
    ([12,18],None), ([13,18],None), ([14,18],None),
], note='max r = 18 (= R). tкр = 15.')

wb.save('/home/user/Verich/outputs/Эконометрика УСР 5.xlsx')
print('Excel saved')
