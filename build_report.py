# -*- coding: utf-8 -*-
"""Build the Word report (Тема 5) — black-and-white, reference-style."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ASSETS = '/home/user/Verich/report_assets/'

doc = Document()

styles = doc.styles
normal = styles['Normal']
normal.font.name = 'Times New Roman'
normal.font.size = Pt(14)
normal.font.color.rgb = RGBColor(0, 0, 0)
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.line_spacing = 1.0
normal.paragraph_format.widow_control = False
normal.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

heading = styles['Heading 1']
heading.font.name = 'Times New Roman'
heading.font.size = Pt(16)
heading.font.bold = True
heading.font.color.rgb = RGBColor(0, 0, 0)
heading.paragraph_format.space_before = Pt(24)
heading.paragraph_format.space_after = Pt(12)
heading.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
heading.paragraph_format.keep_with_next = True

sec = doc.sections[0]
sec.page_width = 7560310
sec.page_height = 10692130
sec.left_margin = 1080135
sec.right_margin = 539750
sec.top_margin = 720090
sec.bottom_margin = 720090

# ---------- helpers ----------
def set_run_font(run, size=None, bold=None, name='Times New Roman', color='000000'):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    rfonts.set(qn('w:ascii'), name)
    rfonts.set(qn('w:hAnsi'), name)
    rfonts.set(qn('w:cs'), name)
    rfonts.set(qn('w:eastAsia'), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)

def para(text='', align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=14, bold=False,
         space_after=None, space_before=None, line=None):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    if space_after is not None:
        pf.space_after = Pt(space_after)
    if space_before is not None:
        pf.space_before = Pt(space_before)
    if line is not None:
        pf.line_spacing = line
    if text:
        r = p.add_run(text)
        set_run_font(r, size=size, bold=bold)
    return p

def heading1(text):
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1']
    r = p.add_run(text)
    set_run_font(r, size=16, bold=True, color='000000')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(28)
    p.paragraph_format.space_after = Pt(16)
    return p

def caption(text):
    p = para(text, align=WD_ALIGN_PARAGRAPH.CENTER, size=12, bold=True, space_before=6, space_after=6)
    return p

def set_cell_bg(cell, hexcolor='FFFFFF'):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)

def set_cell_margins(cell, top=40, bottom=40, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for tag, v in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        el = OxmlElement(f'w:{tag}')
        el.set(qn('w:w'), str(v))
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tcPr.append(mar)

def set_table_borders(table, sz=4):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        borders.append(el)
    tblPr.append(borders)

def add_table(headers, rows, widths_cm, header_size=11, body_size=12):
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    tbl = table._tbl
    tblPr = tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)
    table.autofit = False
    table.allow_autofit = False
    for i, w in enumerate(widths_cm):
        for cell in table.columns[i].cells:
            cell.width = Cm(w)
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:tblHeader'))
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        r = p.add_run(str(h))
        set_run_font(r, size=header_size, bold=True)
        set_cell_margins(cell)
        set_cell_bg(cell)
        cell.vertical_alignment = 1
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            r = p.add_run(str(val))
            set_run_font(r, size=body_size)
            set_cell_margins(cell)
            set_cell_bg(cell)
            cell.vertical_alignment = 1
    # center helper table column sums
    return table

def add_image(path, width_cm, caption_text=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(path, width=Cm(width_cm))
    if caption_text:
        para(caption_text, align=WD_ALIGN_PARAGRAPH.CENTER, size=12, bold=True, space_before=2, space_after=8)

def page_break():
    doc.add_page_break()

# ---------- cover page ----------
def centered(text, size=14, bold=False, after=0, before=0):
    p = para(text, align=WD_ALIGN_PARAGRAPH.CENTER, size=size, bold=bold, space_after=after, space_before=before)
    return p

centered('МИНИСТЕРСТВО ОБРАЗОВАНИЯ РЕСПУБЛИКИ БЕЛАРУСЬ', size=14, after=0)
centered('УО «ПОЛЕССКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ»', size=14, after=0)
centered('', after=20)
centered('Инженерный факультет', size=14, after=0)
centered('', after=20)
centered('Кафедра информационных технологий и интеллектуальных систем', size=14, after=0)
centered('', before=0)
centered('ОТЧЁТ', size=16, bold=True, after=10, before=54)
centered('УСР Тема 5', size=14, bold=True, after=10)
centered('по дисциплине: «Эконометрика и исследование операций в экономике»', size=14, after=0)
centered('', after=8)
centered('на тему: «Модели сетевого планирования и управления»', size=14, after=0)
centered('', after=8)
centered('Вариант в', size=14, after=0)
centered('', after=28)
centered('Выполнил:', size=14, after=0)
centered('студент группы 24ИТ-3                                           Ильич Э.С.', size=14, after=18)
centered('Проверил:', size=14, after=0)
centered('Ассистент                                                          Верич А.И.', size=14, after=120)
centered('ПИНСК 2026', size=14, after=0)

page_break()

# ---------- ЗАДАЧИ ----------
heading1('ЗАДАЧИ')

def problem(num, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(num)
    set_run_font(r, size=14, bold=True)
    r = p.add_run('  ' + text)
    set_run_font(r, size=14)

def problem_instr(text):
    p = para(text, size=13, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6)
    p.paragraph_format.first_line_indent = Cm(0.5)
    return p

def task_table(caption_text, headers, rows, widths_cm):
    caption(caption_text)
    add_table(headers, rows, widths_cm)

problem('5.1(в)', 'Проекты включают в себя следующие работы:')
task_table('Таблица 5.1.1 – Сетевой график проекта (вариант в)',
           ['Работа', 'Предшествующие работы', 'Время выполнения'],
           [['1а', '—', '4'], ['2а', '—', '6'], ['3а', '2а', '2'], ['4а', '2а', '6'],
            ['5а', '1а, 3а', '3'], ['6а', '1а, 3а', '3'], ['7а', '4а, 6а', '5']],
           [5, 5, 5])
problem_instr('Требуется определить ранние tп(i) и поздние tн(j) сроки свершения событий, параметры работ tп.н., tр.о., tн.н., tн.о., Rп, Rс, Rн, построить сетевой график и линейный график проекта, выделить критический путь.')

problem('5.3(в)', 'Проект представлен следующим сетевым графиком:')
task_table('Таблица 5.3.1 – Исходные данные для решения задачи 5.3(в)',
           ['Работа', 'tо (дн.)', 'd (дн.)', 'k'],
           [['(1,2)', '6', '5', '0,05'], ['(1,3)', '13', '10', '0,25'], ['(1,4)', '20', '16', '0,30'],
            ['(2,4)', '9', '7', '0,07'], ['(2,5)', '14', '11', '0,15'], ['(3,4)', '16', '13', '0,10'],
            ['(3,6)', '15', '12', '0,05'], ['(4,5)', '10', '7', '0,03'], ['(4,6)', '17', '15', '0,14'],
            ['(5,6)', '13', '9', '0,50']],
           [6, 4, 4, 4])
problem_instr('Требуется определить tнij, тоij, xij, чтобы проект был выполнен за время tкр = 40 дн., а суммарное количество дополнительных средств было минимальным.')

problem('5.5(в)', 'Проект представлен следующим сетевым графиком:')
task_table('Таблица 5.5.1 – Исходные данные для решения задачи 5.5(в)',
           ['Работа', 'tо (дн.)', 'd (дн.)', 'k'],
           [['(1,2)', '15', '12', '0,25'], ['(1,3)', '8', '5', '0,20'], ['(1,4)', '7', '4', '0,15'],
            ['(2,3)', '5', '3', '0,10'], ['(3,4)', '13', '10', '0,30'], ['(3,5)', '11', '8', '0,40'],
            ['(4,5)', '7', '4', '0,20']],
           [6, 4, 4, 4])
problem_instr('Требуется найти tнij, тоij, xij, чтобы время выполнения проекта было минимальным, а количество используемых дополнительных средств не превышало В = 47 ден. ед.')

problem('5.8(в)', 'Сетевой график проекта изображен на рисунке. Параметры работ:')
task_table('Таблица 5.8.1 – Исходные данные для решения задачи 5.8(в)',
           ['Работа', 'D', 'd', 'c', 'C'],
           [['(1,2)', '6', '2', '6', '14'], ['(1,3)', '11', '6', '2', '17'], ['(2,3)', '4', '1', '4', '13'],
            ['(2,4)', '7', '3', '3', '19'], ['(2,5)', '9', '4', '2', '27'], ['(3,5)', '3', '1', '5', '7'],
            ['(4,5)', '8', '3', '1', '31']],
           [5, 3, 3, 3, 4])
problem_instr('Требуется сократить время выполнения проекта при минимальном возрастании его стоимости.')

problem('5.16(в)', 'Произведена оптимизация проекта по ресурсам, изображенного на рисунке. Числа, приписанные дугам графика, означают времена выполнения работ и необходимое количество ресурса для их выполнения. Количество выделенного ресурса R = 18 ед. Задачу решить в двух вариантах: а) работы не допускают перерыва в выполнении; б) работы допускают разрыв в выполнении.')
task_table('Таблица 5.16.1 – Исходные данные для решения задачи 5.16(в)',
           ['Работа', 't (дн.)', 'r (ед.)'],
           [['(1,2)', '4', '8'], ['(1,3)', '6', '5'], ['(1,4)', '8', '9'], ['(2,4)', '2', '10'],
            ['(2,5)', '5', '4'], ['(3,4)', '2', '1'], ['(3,5)', '3', '2'], ['(4,5)', '4', '12']],
           [6, 4, 4])

page_break()

# ---------- РЕШЕНИЕ ----------
heading1('РЕШЕНИЕ')

# ============ 5.1 ============
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(8)
r = p.add_run('5.1(в)')
set_run_font(r, size=15, bold=True)

para('Сетевой график проекта представлен на рис. 1. Построение сетевого графика и вычисление ранних tп(i) и поздних tн(j) сроков свершения событий позволяют определить полный резерв времени каждого события R(i) = tн(i) − tп(i) (табл. 5.1.2).')
add_image(ASSETS + 'net_5_1.png', 15.4, 'Рисунок 1 – Сетевой график проекта (в)')

caption('Таблица 5.1.2 – Параметры событий')
add_table(['Событие', 'tп (дн.)', 'tн (дн.)', 'R'],
          [['1', '0', '0', '0'], ['2', '4', '9', '5'], ['3', '6', '6', '0'], ['4', '8', '9', '1'],
           ['5', '12', '12', '0'], ['6', '8', '9', '1'], ['7', '11', '17', '6'], ['8', '11', '12', '1'],
           ['9', '12', '12', '0'], ['10', '17', '17', '0']],
          [4, 5, 5, 5])

para('Величины параметров работ определяются соотношениями: tп.н.(ij) = tп(i); tр.о.(ij) = tп.н.(ij) + tij; tн.о.(ij) = tн(j); tн.н.(ij) = tн.о.(ij) − tij; Rп(ij) = tн(j) − tр(i) − tij; Rс(ij) = tр.о.(j) − tр.о.(i) − tij; Rн(ij) = max{0, tр.о.(j) − tн.н.(i) − tij} (табл. 5.1.3).', space_before=12)
caption('Таблица 5.1.3 – Параметры работ')
add_table(['Работа', 'tп.н.', 'tр.о.', 'tн.н.', 'tн.о.', 'Rп', 'Rс'],
          [['1а', '0', '4', '5', '9', '5', '0'], ['2а', '0', '6', '0', '6', '0', '0'],
           ['3а', '6', '8', '7', '9', '1', '0'], ['4а', '6', '12', '6', '12', '0', '0'],
           ['5а', '8', '11', '14', '17', '6', '0'], ['6а', '8', '11', '9', '12', '1', '0'],
           ['7а', '12', '17', '12', '17', '0', '0']],
          [4.5, 2.8, 2.8, 2.8, 2.8, 2.8, 2.8])

para('Критический путь: 1→3→5→9→10; критические работы: 2а, 4а, Ф(5,9), 7а; критические события: 1, 3, 5, 9, 10. tкр = 17 дн.', bold=True)
para('Линейный (график Ганта) график проекта с критическим путём изображён на рис. 2.', space_before=6)
add_image(ASSETS + 'gantt_5_1.png', 15.4, 'Рисунок 2 – Линейный график проекта (в), критический путь выделен')

# ============ 5.3 ============
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(12)
p.paragraph_format.space_after = Pt(8)
r = p.add_run('5.3(в)')
set_run_font(r, size=15, bold=True)

para('Сетевой график проекта представлен на рис. 3. Дополнительные средства xij позволяют сократить продолжительность работ: sij = tij − xij/kij.')
add_image(ASSETS + 'net_5_3.png', 14.6, 'Рисунок 3 – Сетевой график проекта 5.3(в)')

para('Сводные данные (табл. 5.3.2):', space_before=4)
caption('Таблица 5.3.2 – Исходные данные для решения задачи 5.3(в)')
add_table(['Работа', 'tо (дн.)', 'd (дн.)', 'k'],
          [['(1,2)', '6', '5', '0,05'], ['(1,3)', '13', '10', '0,25'], ['(1,4)', '20', '16', '0,30'],
           ['(2,4)', '9', '7', '0,07'], ['(2,5)', '14', '11', '0,15'], ['(3,4)', '16', '13', '0,10'],
           ['(3,6)', '15', '12', '0,05'], ['(4,5)', '10', '7', '0,03'], ['(4,6)', '17', '15', '0,14'],
           ['(5,6)', '13', '9', '0,50']],
          [6, 4, 4, 4])

para('Составим математическую модель задачи. Цель оптимизации – минимизация суммарных дополнительных средств, при ограничении, что проект выполнен за время tкр = 40 дн.:', space_before=10)
para('Σ kij·xij → min', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, space_before=4, space_after=2)
para('Σ tij(новая) = 40 (ограничение на длительность проекта)', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=2)
para('Сводное решение задачи (табл. 5.3.3):', space_before=8)
caption('Таблица 5.3.3 – Сводное решение задачи 5.3(в)')
add_table(['Работа', 'tн (дн.)', 'то (дн.)', 'x (ден. ед.)', 'новая продолжительность (дн.)'],
          [['(1,2)', '0,0', '6,0', '0,000', '6,0'], ['(1,3)', '0,0', '10,0', '12,000', '10,0'],
           ['(1,4)', '0,0', '20,0', '0,000', '20,0'], ['(2,4)', '6,0', '15,0', '0,000', '9,0'],
           ['(2,5)', '17,0', '31,0', '0,000', '14,0'], ['(3,4)', '10,0', '23,0', '30,000', '13,0'],
           ['(3,6)', '25,0', '40,0', '0,000', '15,0'], ['(4,5)', '23,0', '31,0', '66,667', '8,0'],
           ['(4,6)', '23,0', '40,0', '0,000', '17,0'], ['(5,6)', '31,0', '40,0', '8,000', '9,0'],
           ['Σx', '', '', '116,667', '']],
          [4.2, 2.8, 2.8, 4, 5.8])
para('Проект выполнен за tкр = 40 дн. и tн.н. (1,3) = 0 дн. Суммарные дополнительные средства: Σx ≈ 116,67 ден. ед.', bold=True, space_before=10)

# ============ 5.5 ============
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(12)
p.paragraph_format.space_after = Pt(8)
r = p.add_run('5.5(в)')
set_run_font(r, size=15, bold=True)

para('Сетевой график проекта представлен на рис. 4. Дополнительные средства xij позволяют сократить продолжительность работ: sij = tij − xij/kij.')
add_image(ASSETS + 'net_5_5.png', 14.6, 'Рисунок 4 – Сетевой график проекта 5.5(в)')

para('Сводные данные (табл. 5.5.2):', space_before=4)
caption('Таблица 5.5.2 – Исходные данные для решения задачи 5.5(в)')
add_table(['Работа', 'tо (дн.)', 'd (дн.)', 'k'],
          [['(1,2)', '15', '12', '0,25'], ['(1,3)', '8', '5', '0,20'], ['(1,4)', '7', '4', '0,15'],
           ['(2,3)', '5', '3', '0,10'], ['(3,4)', '13', '10', '0,30'], ['(3,5)', '11', '8', '0,40'],
           ['(4,5)', '7', '4', '0,20']],
          [6, 4, 4, 4])

para('Составим математическую модель задачи. Цель – минимизация длительности проекта tкр, при ограничении, что количество дополнительных средств не превышает В = 47 ден. ед.:', space_before=10)
para('t = t5  →  min', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, space_before=4, space_after=2)
para('Σ kij·xij ≤ 47', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, space_after=2)

para('Сводное решение задачи (табл. 5.5.3):', space_before=8)
caption('Таблица 5.5.3 – Сводное решение задачи 5.5(в)')
add_table(['Работа', 'tн (дн.)', 'то (дн.)', 'x (ден. ед.)', 'новая продолжительность (дн.)'],
          [['(1,2)', '0,0', '12,0', '12,000', '12,0'], ['(1,3)', '0,0', '8,0', '0,000', '8,0'],
           ['(1,4)', '0,0', '7,0', '0,000', '7,0'], ['(2,3)', '12,0', '16,0', '10,000', '4,0'],
           ['(3,4)', '16,0', '26,0', '10,000', '10,0'], ['(3,5)', '19,0', '30,0', '0,000', '11,0'],
           ['(4,5)', '26,0', '30,0', '15,000', '4,0'],
           ['Σx', '', '', '47,000', '']],
          [4.2, 2.8, 2.8, 4, 5.8])
para('Проект выполнен за tкр = 30 дн. Суммарные дополнительные средства: Σx = 47 (= B) ден. ед.', bold=True, space_before=10)

# ============ 5.8 ============
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(12)
p.paragraph_format.space_after = Pt(8)
r = p.add_run('5.8(в)')
set_run_font(r, size=15, bold=True)

para('Сетевой график проекта представлен на рис. 5. Для каждой работы определяем коэффициент сокращения длительности (КДЗ), равный отношению разности стоимостей к разности длительностей: КДЗ = (Cij − cij)/(Dij − dij).')
add_image(ASSETS + 'net_5_8.png', 14.6, 'Рисунок 5 – Сетевой график проекта 5.8(в)')

caption('Таблица 5.8.2 – Исходные данные и показатели КДЗ для решения задачи 5.8(в)')
add_table(['Работа', 'D (дн.)', 'd (дн.)', 'c (ден. ед.)', 'C (ден. ед.)', 'КДЗ'],
          [['(1,2)', '6', '2', '6', '14', '2,00'], ['(1,3)', '11', '6', '2', '17', '3,00'],
           ['(2,3)', '4', '1', '4', '13', '3,00'], ['(2,4)', '7', '3', '3', '19', '4,00'],
           ['(2,5)', '9', '4', '2', '27', '5,00'], ['(3,5)', '3', '1', '5', '7', '1,00'],
           ['(4,5)', '8', '3', '1', '31', '6,00']],
          [4.5, 2.6, 2.6, 2.8, 2.8, 2.6])

para('Отправная точка оптимизации – критический путь (1,2)→(2,4)→(4,5) с суммарной длительностью 17 дн. и стоимостью 23 ден. ед. На каждом шаге сокращаются критические работы (наименьший КДЗ). Ход оптимизации приведён в табл. 5.8.3.', space_before=10)
caption('Таблица 5.8.3 – Ход оптимизации проекта 5.8(в)')
add_table(['Шаг', 'Сокращённые работы', 'Величина сокращения (дн.)', 'tкр (дн.)', 'Стоимость (ден. ед.)'],
          [['1', '(1,2)', '4', '17', '31'], ['2', '(2,4)', '3', '14', '43'],
           ['3', '(2,4), (3,5)', '1', '13', '48'], ['4', '(4,5), (3,5)', '1', '12', '55'],
           ['5', '(4,5), (1,3)', '1', '11', '64'], ['6', '(4,5), (2,5), (1,3)', '3', '8', '106'],
           ['7', '(2,5), (1,3)', '1', '8', '114']],
          [2.4, 6.6, 4, 2.6, 3.4])
para('Итоговые продолжительности работ D/d: (1,2) = 6→2, (1,3) = 11→6, (2,3) = 4→1, (2,4) = 7→3, (2,5) = 9→4, (3,5) = 3→1, (4,5) = 8→3.', space_before=8)
para('Проект выполнен за tкр = 8 дн. при стоимости 114 ден. ед. (норма – 23 ден. ед.).', bold=True, space_before=6)

# ============ 5.16 ============
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(12)
p.paragraph_format.space_after = Pt(8)
r = p.add_run('5.16(в)')
set_run_font(r, size=15, bold=True)

para('Сетевой график проекта представлен на рис. 6. Количество выделенного ресурса R = 18 ед. Числа, приписанные дугам графика, означают времена выполнения работ и необходимое количество ресурса для их выполнения (табл. 5.16.2).')
add_image(ASSETS + 'net_5_16.png', 14.6, 'Рисунок 6 – Сетевой график проекта 5.16(в)')

caption('Таблица 5.16.2 – Исходные данные для решения задачи 5.16(в)')
add_table(['Работа', 't (дн.)', 'r (ед.)'],
          [['(1,2)', '4', '8'], ['(1,3)', '6', '5'], ['(1,4)', '8', '9'], ['(2,4)', '2', '10'],
           ['(2,5)', '5', '4'], ['(3,4)', '2', '1'], ['(3,5)', '3', '2'], ['(4,5)', '4', '12']],
          [6, 4, 4])

# ---- 5.16(а) ----
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(12)
r = p.add_run('5.16(в) а) – работы не допускают перерыва')
set_run_font(r, size=13, bold=True)

para('Задача решается мето­дом последовательного перекрытия работ, которые не могут быть начаты в заданный момент из-за нехватки ресурса. Оптимальное расписание приведено в табл. 5.16.3. Максимальное потребление ресурса не превышает R.', space_before=4)
caption('Таблица 5.16.3 – Оптимальное расписание задачи 5.16(в) а)')
add_table(['Работа', 'tн (дн.)', 'то (дн.)', 't (дн.)', 'r (ед.)'],
          [['(1,2)', '0', '4', '4', '8'], ['(1,3)', '4', '10', '6', '5'],
           ['(1,4)', '0', '8', '8', '9'], ['(2,4)', '10', '12', '2', '10'],
           ['(2,5)', '8', '13', '5', '4'], ['(3,4)', '10', '12', '2', '1'],
           ['(3,5)', '11', '14', '3', '2'], ['(4,5)', '12', '16', '4', '12']],
          [6, 3.2, 3.2, 3.2, 3.2])

para('Эпюра потребности в ресурсе (табл. 5.16.4):', space_before=10)
caption('Таблица 5.16.4 – Эпюра потребности в ресурсе (5.16(в) а))')
add_table(['Период (дн.)', 'r (ед.)'],
          [['0', '17'], ['1', '17'], ['2', '17'], ['3', '17'], ['4', '14'], ['5', '14'],
           ['6', '14'], ['7', '14'], ['8', '9'], ['9', '9'], ['10', '15'], ['11', '17'],
           ['12', '18'], ['13', '14'], ['14', '12'], ['15', '12']],
          [6, 6])
para('Линейный график проекта изображён на рис. 7. Максимальное потребление ресурса r = 18 ≤ R = 18 ед. tкр = 16 дн.', space_before=8)
add_image(ASSETS + 'gantt_5_16a.png', 15.4, 'Рисунок 7 – Линейный график проекта 5.16(в) а) (не допускают перерыва)')

# ---- 5.16(б) ----
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(12)
r = p.add_run('5.16(в) б) – работы допускают разрыв')
set_run_font(r, size=13, bold=True)

para('Если работы допускают разрыв во времени, задача решается тем же методом с расщеплением работ на части. Оптимальное расписание приведено в табл. 5.16.5.', space_before=4)
caption('Таблица 5.16.5 – Оптимальное расписание задачи 5.16(в) б)')
add_table(['Работа', 'Интервалы выполнения', 't (дн.)', 'r (ед.)'],
          [['(1,2)', '[0–4)', '4', '8'], ['(1,3)', '[2–3), [4–9)', '6', '5'],
           ['(1,4)', '[0–2), [3–9)', '8', '9'], ['(2,4)', '[9–11)', '2', '10'],
           ['(2,5)', '[10–15)', '5', '4'], ['(3,4)', '[9–11)', '2', '1'],
           ['(3,5)', '[12–15)', '3', '2'], ['(4,5)', '[11–15)', '4', '12']],
          [5.5, 8, 3, 3.5])

para('Эпюра потребности в ресурсе (табл. 5.16.6):', space_before=10)
caption('Таблица 5.16.6 – Эпюра потребности в ресурсе (5.16(в) б))')
add_table(['Период (дн.)', 'r (ед.)'],
          [['0', '17'], ['1', '17'], ['2', '13'], ['3', '17'], ['4', '14'], ['5', '14'],
           ['6', '14'], ['7', '14'], ['8', '14'], ['9', '11'], ['10', '15'], ['11', '16'],
           ['12', '18'], ['13', '18'], ['14', '18']],
          [6, 6])
para('Линейный график проекта изображён на рис. 8. Максимальное потребление ресурса r = 18 = R = 18 ед. tкр = 15 дн.', space_before=8)
add_image(ASSETS + 'gantt_5_16b.png', 15.4, 'Рисунок 8 – Линейный график проекта 5.16(в) б) (допускают разрыв)')

page_break()

# ---------- ПРИЛОЖЕНИЕ А ----------
heading1('ПРИЛОЖЕНИЕ А')
para('Решение задач 5.1(в)–5.16(в) в среде Excel (см. файл «Эконометрика УСР 5.xlsx»).',
     size=12, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6)

caption('Рисунок А.1 – Решение задачи 5.1(в) в Excel')
add_image(ASSETS + 'excel_5.1.png', 16.2, None)
caption('Рисунок А.2 – Решение задачи 5.3(в) в Excel')
add_image(ASSETS + 'excel_5.3.png', 16.2, None)
caption('Рисунок А.3 – Решение задачи 5.5(в) в Excel')
add_image(ASSETS + 'excel_5.5.png', 16.2, None)
caption('Рисунок А.4 – Решение задачи 5.8(в) в Excel')
add_image(ASSETS + 'excel_5.8.png', 16.2, None)
caption('Рисунок А.5 – Решение задачи 5.16(в) а) в Excel')
add_image(ASSETS + 'excel_5_16a.png', 16.2, None)
caption('Рисунок А.6 – Решение задачи 5.16(в) б) в Excel')
add_image(ASSETS + 'excel_5_16b.png', 16.2, None)

sec = doc.sections[0]
footer = sec.footer
footer.is_linked_to_previous = False
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run('Тема 5  ·  УСР  ·  2026')
set_run_font(r, size=9, bold=False)

doc.save('/home/user/Verich/outputs/Эконометрика УСР 5.docx')
print('Word report saved')
