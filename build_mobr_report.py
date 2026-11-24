# -*- coding: utf-8 -*-
"""Build the Word report (УСР №1 — Тема 1 «Модели межотраслевого баланса»).
Black-and-white, reference-style. Data source: «Пась А.М. 24ИТ-3 УСР Свитина.xlsx»."""
import numpy as np
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ASSETS = '/home/user/Verich/report_assets/'
OUT = '/home/user/Verich/outputs/Эконометрика УСР 1 (Пась А.М.).docx'

# ============== data (verified against the workbook) ==============
x_raw = np.array([[30,10,15],[35,50,20],[15,25,30]], dtype=float)
Y_OCH = np.array([90,25,60], dtype=float)
Y_PLAN = np.array([120,30,60], dtype=float)
ZSHARE = 0.3

X_OCH = x_raw.sum(axis=1) + Y_OCH
A = x_raw / X_OCH[np.newaxis, :]
EA = np.eye(3) - A
B = np.linalg.inv(EA)
X_PLAN = B @ Y_PLAN
# пром. затраты по потребителю j: sum_i a_{i,j} * X_j = X_j * sum_i a_{i,j}
PROM_PLAN_CONS = X_PLAN * A.sum(axis=0)
Z_OCH = X_OCH - x_raw.sum(axis=0)
Z_PLAN = X_PLAN - PROM_PLAN_CONS
P_OTCH = ZSHARE * Z_OCH
P_PLAN = ZSHARE * Z_PLAN
PR_OCH = Z_OCH - P_OTCH
PR_PLAN = Z_PLAN - P_PLAN

# Пункт 4 (n1=1, r=2, p1=30 -> elasticity 0.7, Kw = 1 + (r-1)*0.7 = 1.7)
KW = 1 + 1.0 * (1 - 0.30)
W2 = np.array([np.array([1 - A[1,1], -A[2,1]]), [-A[1,2], 1 - A[2,2]]])
b2 = np.array([2.0 * A[0,1] + (P_OTCH[1] * KW + PR_OCH[1]) / X_OCH[1],
               2.0 * A[0,2] + (P_OTCH[2] * KW + PR_OCH[2]) / X_OCH[2]])
p_res = np.linalg.solve(W2, b2)
p2_4, p3_4 = p_res

# Пункт 5 (n2=2, з/п grows 30% -> p_2 = 1.3)
W3 = np.array([[1 - A[0,0], -A[2,0]], [-A[0,2], 1 - A[2,2]]])
b3 = np.array([1.3 * A[1,0] + Z_OCH[0] / X_OCH[0],
               1.3 * A[1,2] + Z_OCH[2] / X_OCH[2]])
p_res3 = np.linalg.solve(W3, b3)
p1_5, p3_5 = p_res3

print('X отч:', X_OCH)
print('A:\n', A)
print('B:\n', B)
print('X план:', X_PLAN)
print('пром затраты план:', PROM_PLAN_CONS)
print('ВДС отч:', Z_OCH, 'з/п отч:', P_OTCH, 'прочие отч:', PR_OCH)
print('ВДС план:', Z_PLAN, 'з/п план:', P_PLAN, 'прочие план:', PR_PLAN)
print('п.4 p2,p3:', p2_4, p3_4)
print('п.5 p1,p3:', p1_5, p3_5)

# ================== document ==================
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

from toc import Toc  # noqa: E402

TOC = Toc(doc, title='ОГЛАВЛЕНИЕ', size=14, levels=2)

sec = doc.sections[0]
sec.page_width = 7560310
sec.page_height = 10692130
sec.left_margin = 1080135
sec.right_margin = 539750
sec.top_margin = 720090
sec.bottom_margin = 720090

def set_run_font(run, size=None, bold=None, name='Times New Roman', color='000000', italic=False):
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
    if italic:
        run.font.italic = True
    run.font.color.rgb = RGBColor.from_string(color)

def para(text='', align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=14, bold=False,
         space_after=None, space_before=None, line=None, italic=False):
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
        set_run_font(r, size=size, bold=bold, italic=italic)
    return p

def heading1(text):
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1']
    r = p.add_run(text)
    set_run_font(r, size=16, bold=True, color='000000')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(28)
    p.paragraph_format.space_after = Pt(16)
    TOC.add(p, 1)
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

def add_table(headers, rows, widths_cm, header_size=11, body_size=12, merge_rows=None):
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    table.autofit = False
    table.allow_autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)
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
    if merge_rows:
        for r0, r1, c0, c1 in merge_rows:
            a = table.rows[r0].cells[c0]
            b_ = table.rows[r1].cells[c1]
            a.merge(b_)
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

def centered(text, size=14, bold=False, after=0, before=0, italic=False):
    p = para(text, align=WD_ALIGN_PARAGRAPH.CENTER, size=size, bold=bold, space_after=after, space_before=before, italic=italic)
    return p

def num_str(values, dec=4):
    return [f'{v:.{dec}f}'.rstrip('0').rstrip('.').replace('.', ',') if dec else str(v) for v in values]

# ================== cover page ==================
centered('МИНИСТЕРСТВО ОБРАЗОВАНИЯ РЕСПУБЛИКИ БЕЛАРУСЬ', size=14, after=0)
centered('УО «ПОЛЕССКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ»', size=14, after=0)
centered('', after=22)
centered('Инженерный факультет', size=14, after=0)
centered('', after=22)
centered('Кафедра информационных технологий и интеллектуальных систем', size=14, after=0)
centered('', before=0)
centered('ОТЧЁТ', size=16, bold=True, after=8, before=54)
centered('УСР №1', size=14, bold=True, after=8)
centered('по дисциплине: «Эконометрика и исследование операций в экономике»', size=14, after=0)
centered('', after=8)
centered('на тему: «Модели межотраслевого баланса»', size=14, after=0)
centered('', after=8)
centered('Вариант б', size=14, after=0)
centered('', after=28)
centered('Выполнил:', size=14, after=0)
centered('студент', size=14, after=0)
centered('группы 24ИТ-3                                   Пась А.М.', size=14, after=4)
centered('Проверил:', size=14, after=0)
centered('Свитина', size=14, after=120)
centered('ПИНСК 2026', size=14, after=0)

page_break()

# ================== ЗАДАЧИ ==================
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
    p.paragraph_format.keep_with_next = True
    TOC.add(p, 2, text='Задача ' + num)
    return p

problem('1.6(б)', 'Для условной экономики, состоящей из трех отраслей, за отчетный период известны межотраслевые потоки x(ij) и вектор объемов конечного использования продукции Yотч. Учитывая, что p% ВДС приходится на з/п, решить следующие задачи:')
para('1) построить схему МОБ за отчетный период с дополнительным выделением промежуточного потребления, промежуточных затрат, з/п, прочих элементов ВДС;', size=13, space_after=4).paragraph_format.first_line_indent = Cm(0.5)
para('2) рассчитать матрицу коэффициентов прямых затрат A;', size=13, space_after=4).paragraph_format.first_line_indent = Cm(0.5)
para('3) определить в плановом периоде валовой выпуск продукции Xпл, если конечное использование определяется вектором Yпл;', size=13, space_after=4).paragraph_format.first_line_indent = Cm(0.5)
para('4) определить, как изменятся цены в отраслях, если цена на продукцию в отрасли n1 увеличится в r раза, при этом учесть, что рост заработной платы во всех отраслях отстает от роста цен на p1%;', size=13, space_after=4).paragraph_format.first_line_indent = Cm(0.5)
para('5) определить, как изменятся индексы цен на продукцию отраслей, если з/п в отрасли n2 увеличиться на p2%, а в других отраслях останется неизменной.', size=13, space_after=6).paragraph_format.first_line_indent = Cm(0.5)

caption('Таблица 1.1 – Исходные данные для задачи 1.6(б)')
add_table(['Отрасли–производители', 'Отрасли–потребители', 'Отрасли–потребители', 'Отрасли–потребители', 'Пром. потребле–ние', 'Кон. использование, Y', 'Валовой выпуск, X'],
          [['1', '30', '10', '15', '55', '90', '145'],
           ['2', '35', '50', '20', '105', '25', '130'],
           ['3', '15', '25', '30', '70', '60', '130'],
           ['Пром. затраты', '80', '85', '65', '', '', ''],
           ['ВДС, Z', '65', '45', '65', '', '175', '175'],
           ['Валовые затраты, X', '145', '130', '130', '', '', '']],
          [2.9, 2.2, 2.2, 2.2, 2.2, 2.6, 2.2])

caption('Таблица 1.2 – Вектор конечного использования Y')
add_table(['Показатель', 'Y отч', 'Y план'],
          [['1-я отрасль', '90', '120'], ['2-я отрасль', '25', '30'], ['3-я отрасль', '60', '60'],
           ['Итого', '175', '210']],
          [6, 4, 4])

caption('Таблица 1.3 – Числовые данные для задачи 1.6(б)')
add_table(['№', 'p', 'n1', 'r', 'p1', 'n2', 'p2'],
          [['б', '30', '1', '2', '30', '2', '30']],
          [4, 4, 4, 4, 4, 4, 4])

page_break()

# ================== РЕШЕНИЕ ==================
heading1('РЕШЕНИЕ')

para('По исходным данным построена схема МОБ за отчетный период (табл. 2.1).')
caption('Таблица 2.1 – Схема МОБ за отчетный период')
add_table(['Отрасли–производители', '1', '2', '3', 'Пром. потребле–ние', 'Кон. использование, Y', 'Валовой выпуск, X'],
          [['1', '30', '10', '15', '55', '90', '145'],
           ['2', '35', '50', '20', '105', '25', '130'],
           ['3', '15', '25', '30', '70', '60', '130'],
           ['Пром. затраты', '80', '85', '65', '', '', ''],
           ['З/п (30% ВДС)', '19,5', '13,5', '19,5', '', '', ''],
           ['Проч. элементы ВДС', '45,5', '31,5', '45,5', '', '', ''],
           ['ВДС, Z', '65', '45', '65', '', '175', '175'],
           ['Валовые затраты, X', '145', '130', '130', '', '', '']],
          [3.4, 2.3, 2.3, 2.3, 2.0, 2.7, 2.2])

para('По формулам (1.4) рассчитываем матрицу A:')
caption('Таблица 2.2 – Матрица коэффициентов прямых затрат A')
add_table(['', '1', '2', '3'],
          [['1', '0,2069', '0,0769', '0,1154'],
           ['2', '0,2414', '0,3846', '0,1538'],
           ['3', '0,1034', '0,1923', '0,2308']],
          [3, 4, 4, 4])

para('Для определения в плановом периоде валового выпуска продукции:', space_before=12)
para('а) строим единичную матрицу E;', space_before=6)
para('б) рассчитываем матрицу E − A;')
caption('Таблица 2.3 – Матрица E − A')
add_table(['', '1', '2', '3'],
          [['1', '0,7931', '−0,0769', '−0,1154'],
           ['2', '−0,2414', '0,6154', '−0,1538'],
           ['3', '−0,1034', '−0,1923', '0,7692']],
          [3, 4, 4, 4])

para('в) с помощью функции МОБР определяем обратную матрицу B:', space_before=6)
caption('Таблица 2.4 – Матрица B = (E − A)⁻¹ (МОБР)')
add_table(['', '1', '2', '3'],
          [['1', '1,3707', '0,2513', '0,2559'],
           ['2', '0,6227', '1,8475', '0,4629'],
           ['3', '0,3400', '0,4957', '1,4501']],
          [3, 4, 4, 4])

para('г) с помощью функции МУМНОЖ рассчитываем в плановом периоде валовой выпуск продукции Xпл = B · Yпл:', space_before=6)
para('Xпл = B · Yпл = (1,3707; 0,6227; 0,3400) · (120; 30; 60) ≈ (187,378; 157,917; 142,678).', italic=True, space_after=10)

para('В таблице 2.5 приведена схема МОБ в плановом периоде.')
caption('Таблица 2.5 – Схема МОБ в плановом периоде')
add_table(['Отрасли–производители', '1', '2', '3', 'Пром. потребле–ние', 'Кон. использование, Y', 'Валовой выпуск, X'],
          [['1', '38,7679', '12,1475', '16,4629', '67,3783', '120', '187,3783'],
           ['2', '45,2292', '60,7374', '21,9505', '127,9171', '30', '157,9171'],
           ['3', '19,3840', '30,3687', '32,9258', '82,6784', '60', '142,6784'],
           ['Пром. затраты', '103,3811', '103,2535', '71,3392', '', '', ''],
           ['З/п (30% ВДС)', '25,1991', '16,3991', '21,4018', '', '', ''],
           ['Проч. элементы ВДС', '58,7980', '38,2645', '49,9375', '', '', ''],
           ['ВДС, Z', '83,9972', '54,6636', '71,3392', '', '210', '210'],
           ['Валовые затраты, X', '187,3783', '157,9171', '142,6784', '', '', '']],
          [3.4, 2.3, 2.3, 2.3, 2.0, 2.7, 2.2])

para('4) Увеличение цены на продукцию первой отрасли в 2 раза приведёт к увеличению цен на продукцию во второй и третьей отраслях в p2 и p3 раз соответственно. Т.к. рост з/п отстаёт от роста цен на 30%, то коэффициент эластичности заработной платы равен 0,7. Тогда схема МОБ будет иметь вид (табл. 2.6):', space_before=12)
caption('Таблица 2.6 – Схема МОБ с учётом роста цены в первой отрасли в 2 раза (n1 = 1, r = 2)')
add_table(['Отрасли–производители', '1', '2', '3'],
          [['1', '30 · 2', '10 · 2', '15 · 2'],
           ['2', '35 · p2', '50 · p2', '20 · p2'],
           ['3', '15 · p3', '25 · p3', '30 · p3'],
           ['З/п', '19,5 · 1,7', '13,5 · 1,7', '19,5 · 1,7'],
           ['Прочие элементы ВДС', '45,5', '31,5', '45,5'],
           ['Валовые затраты, X', '145 · 2', '130 · p2', '130 · p3']],
          [4.5, 3.6, 3.6, 3.6])

para('Поскольку величина затрат на продукцию первой отрасли не влияет на формирование цены в этой отрасли (p1 = 2 задана), то с учётом (1.7) система будет состоять из двух балансовых уравнений для второй и третьей отраслей:', space_before=10)
para('130 · p2 = 10 · 2 + 50 · p2 + 25 · p3 + 13,5 · 1,7 + 31,5', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_before=6, italic=True)
para('130 · p3 = 15 · 2 + 20 · p2 + 30 · p3 + 19,5 · 1,7 + 45,5', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True)
para('После преобразования получим:', space_before=8)
para('80 · p2 = 74,45 + 25 · p3', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, bold=True, space_before=2, italic=True)
para('100 · p3 = 108,65 + 20 · p2', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, bold=True, italic=True)
para('Решим данную систему с помощью функций МОБР и МУМНОЖ:', space_before=10)
para('p2 ≈ 1,355;   p3 ≈ 1,357.', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, bold=True, space_before=6)
para('Таким образом, повышение цен на продукцию первой отрасли в 2 раза приведет к повышению цен на продукцию второй отрасли в 1,355 раза, а на продукцию третьей отрасли в 1,357 раза.', space_before=12)

para('5) Источником инфляции является увеличение з/п во второй отрасли на 30% (1,3 раза). Т.к. заработная плата в других отраслях остается неизменной (инвариантность реальной величины), то схема МОБ будет иметь вид:', space_before=14)
para('з/п второй отрасли = 13,5 · 1,3 = 17,55;   з/п первой и третьей отраслей = 19,5 (без изменений).', size=13, space_after=8)
para('Следовательно, индексы цен на продукцию отраслей определяются из решения системы двух балансовых уравнений для первой и третьей отраслей (вторая отрасль уже содержит в себе заложенное изменение з/п, p2 = 1,3 задан):')
para('145 · p1 = 30 · p1 + 35 · 1,3 + 15 · p3 + 65', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_before=6, italic=True)
para('130 · p3 = 15 · p1 + 20 · 1,3 + 30 · p3 + 65', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True)
para('После преобразования получим:', space_before=8)
para('115 · p1 = 110,5 + 15 · p3', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, bold=True, space_before=2, italic=True)
para('100 · p3 = 91 + 15 · p1', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, bold=True, italic=True)
para('Решим данную систему с помощью функций МОБР и МУМНОЖ:', space_before=10)
para('p1 ≈ 1,1011;   p2 = 1,3;   p3 ≈ 1,0752', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, bold=True, space_before=6)
para('Таким образом, увеличение з/п во второй отрасли на 30% вызовет повышение цен на продукцию первой отрасли в 1,1011 раза, второй – в 1,3 раза, третьей – в 1,0752 раза.', space_before=12)

page_break()

# ================== ПРИЛОЖЕНИЕ А ==================
heading1('ПРИЛОЖЕНИЕ А')
para('Решение задачи 1.6(б) в среде Excel (см. файл «Пась А.М. 24ИТ-3 УСР Свитина.xlsx», лист «Лист2»).', size=12, space_after=6)
caption('Рисунок А.1 – Решение задачи 1.6(б) в Excel: схемы МОБ за отчетный и плановый периоды, матрица коэффициентов прямых затрат, матрица E − A, обратная матрица B')
add_image(ASSETS + 'excel_mobr_2a.png', 17.0)
page_break()
caption('Рисунок А.2 – Решение задачи 1.6(б) в Excel: пункты 4 и 5')
add_image(ASSETS + 'excel_mobr_2b.png', 17.0)

page_break()

# ================== ПРИЛОЖЕНИЕ Б ==================
heading1('ПРИЛОЖЕНИЕ Б')

para('Формула (1.1) описывает систему уравнений, которые называются уравнениями распределения продукции отраслей по направлениям использования:')
para('Σj x(ij) + Y(i) = X(i),   i = 1, …, n', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=8)

para('Формула (1.2) описывает систему уравнений, которые отражают стоимостной состав продукции всех отраслей:')
para('Σi x(ij) + Z(j) = X(j),   j = 1, …, n', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=8)

para('Просуммировав уравнения (1.1) и (1.2) по всем отраслям, получим:')
para('Σi X(i) = Σi Z(i) + Σj Y(j)', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=6)
para('Вычтя из первого уравнения второе, получим:')
para('Σi X(i) = Σi Z(i) + Σj Y(j)  ⇒  Σi Z(i) = Σj Y(j)', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=6)
para('Последнее уравнение показывает, что в межотраслевом балансе соблюдается важнейший принцип единства материального и стоимостного состава ВВП.', space_after=10)

para('Основу ЭММ МОБ составляет технологическая матрица A, содержащая коэффициенты прямых материальных затрат aij на производство единицы продукции  j-й отрасли, которые рассчитываются по формуле:')
para('aij = x(ij) / X(j),   i, j = 1, …, n', align=WD_ALIGN_PARAGRAPH.CENTER, size=13, italic=True, space_after=8)
para('Коэффициенты прямых материальных затрат aij показывают, какое количество продукции i-й отрасли необходимо, если учитывать только прямые затраты, для производства единицы продукции j-й отрасли. Коэффициенты прямых материальных затрат не зависят от объема производства в отрасли и являются довольно стабильной величиной во времени.')

para('С учётом формулы (1.4) систему уравнений баланса (1.1) можно переписать в виде:')
para('Σj aij X(j) + Y(i) = X(i),   i = 1, …, n', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=6)
para('В векторно-матричной форме система уравнений (1.5) примет вид:')
para('X = A·X + Y,   X = (E − A)⁻¹ · Y', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=8)
para('Система уравнений (1.5), или в матричной форме (1.6), называется ЭММ МОБ (модель Леонтьева, модель «затраты – выпуск»).')

para('С помощью модели (1.6) можно выполнять следующие расчеты.')
para('Задав в модели величины валовой продукции каждой отрасли, можно определить объемы конечной продукции каждой отрасли по формуле:')
para('Y = (E − A) · X = (E − A) · X', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=6)
para('где E – единичная матрица n-го порядка.')
para('Задав величины конечной продукции всех отраслей, можно определить величины валовой продукции каждой отрасли:')
para('X = (E − A)⁻¹ · Y', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=6)
para('где (E − A)⁻¹ – матрица полных материальных затрат (обратная матрица к (E − A)); элементы матрицы (E − A)⁻¹ показывают, какое количество продукции i-й отрасли нужно произвести, чтобы с учетом прямых и косвенных затрат этой продукции получить единицу конечной продукции j-й отрасли.')
para('Формула (1.7) является базовой балансовой моделью прогнозирования цен в отраслях экономики:')
para('p(j) · X(j) = Σi p(i) · x(ij) + p(j) · Z(j),   j = 1, …, n', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=6)

# footer: слева подпись, справа номер страницы
sec = doc.sections[0]
sec.different_first_page_header_footer = True
footer = sec.footer
footer.is_linked_to_previous = False
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
tabs = OxmlElement('w:tabs')
tab = OxmlElement('w:tab')
tab.set(qn('w:val'), 'right')
tab.set(qn('w:pos'), '9350')
tabs.append(tab)
fp._p.get_or_add_pPr().append(tabs)
r = fp.add_run('УСР №1  ·  Тема 1  ·  2026')
set_run_font(r, size=9, bold=False)
r = fp.add_run('\t')
set_run_font(r, size=9, bold=False)
fld = OxmlElement('w:fldSimple')
fld.set(qn('w:instr'), 'PAGE')
fr = OxmlElement('w:r')
ft = OxmlElement('w:t')
ft.text = '2'
fr.append(ft)
fld.append(fr)
fp._p.append(fld)

TOC.finish()          # ОГЛАВЛЕНИЕ + номера страниц (Word пересчитает поля при открытии)
doc.save(OUT)
print('оглавление:', TOC.entries())
print('saved', OUT)
