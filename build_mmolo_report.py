# -*- coding: utf-8 -*-
"""Сборка отчёта УСР по теме 2 «Модели и методы многоцелевой оптимизации»,
задание 2.4 (вариант — данные методички, табл. 2.3).

Структура и оформление — по образцу «Эконометрика УСР 3.docx»:
ЗАДАЧА → РЕШЕНИЕ (4 метода, внутри каждого — скриншоты Excel-решения) →
ПРИЛОЖЕНИЕ А (листы Excel-книги) → ПРИЛОЖЕНИЕ Б (теория методов).

Числа берутся из mmolo_models.py / solve_mmolo.py (ЛП-расчёт, сверен с Excel).
"""
import copy
import os
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
from mmolo_models import models, fmt, RES, CAP, C1, C2, C3  # noqa: E402

ASSETS = os.path.join(ROOT, 'report_assets')
OUT = os.path.join(ROOT, 'outputs', 'Эконометрика УСР 2 (Пась А.М.) — задание 2.4.docx')
M = models()
RES_JSON = __import__('json').load(open('/tmp/mmolo_results.json'))
a, b, c, d = (RES_JSON['a'], RES_JSON['b'], RES_JSON['c'], RES_JSON['d'])

MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def ru(v, nd=4):
    s = fmt(v, nd)
    return s.replace('.', ',') if s else ''


# ============================ документ =======================================
doc = Document()
normal = doc.styles['Normal']
normal.font.name = 'Times New Roman'
normal.font.size = Pt(14)
normal.font.color.rgb = RGBColor(0, 0, 0)
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.line_spacing = 1.0
normal.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

sec = doc.sections[0]
sec.page_width, sec.page_height = Emu(7560310), Emu(10692130)
sec.left_margin, sec.right_margin = Emu(1080135), Emu(539750)
sec.top_margin, sec.bottom_margin = Emu(720090), Emu(720090)


def set_run_font(run, size=14, bold=False, name='Times New Roman', italic=False, sup=False, sub=False):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    if sup:
        run.font.superscript = True
    if sub:
        run.font.subscript = True


def para(text='', align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=14, bold=False, italic=False,
         before=None, after=None, line=None, first_indent=None):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line
    if first_indent is not None:
        pf.first_line_indent = Cm(first_indent)
    if text:
        r = p.add_run(text)
        set_run_font(r, size=size, bold=bold, italic=italic)
    return p


def heading(text, page_break=False):
    if page_break:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    p = para(text, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, bold=True, before=16, after=12)
    p.paragraph_format.keep_with_next = True
    return p


def formula(text, size=14):
    """Центрированная формула (текстом) с меткой справа."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run_font(r, size=size, italic=False)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    return p


# ------------------------------ OMML helpers ---------------------------------
def _e(tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn('m:' + k), v)
    return el


def mrun(text, italic=True, bold=False):
    r = _e('m:r')
    rpr = _e('m:rPr') if not italic else None
    run = OxmlElement('w:r')
    rpr_w = OxmlElement('w:rPr')
    rfonts = OxmlElement('w:rFonts')
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rfonts.set(qn(attr), 'Cambria Math')
    rpr_w.append(rfonts)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '28'); rpr_w.append(sz)
    if italic:
        i = OxmlElement('w:i'); rpr_w.append(i)
    if bold:
        OxmlElement('w:b')
    run.append(rpr_w)
    t = OxmlElement('w:t')
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    run.append(t)
    r.append(run)
    return r


def msub(base, sb):
    el = _e('m:sSub')
    el.append(_e('m:e')); el[-1].append(mrun(base) if isinstance(base, str) else base)
    el.append(_e('m:sub')); el[-1].append(mrun(sb) if isinstance(sb, str) else sb)
    return el


def msup(base, sp):
    el = _e('m:sSup')
    el.append(_e('m:e')); el[-1].append(mrun(base) if isinstance(base, str) else base)
    el.append(_e('m:sup')); el[-1].append(mrun(sp) if isinstance(sp, str) else sp)
    return el


def mfrac(num, den):
    el = _e('m:f')
    n = _e('m:num')
    n.append(mrun(num) if isinstance(num, str) else num)
    d_ = _e('m:den')
    d_.append(mrun(den) if isinstance(den, str) else den)
    el.append(n); el.append(d_)
    return el


def mobj(*children):
    """Собирает <m:oMath> из перечисленных элементов/строк."""
    om = OxmlElement('m:oMath')
    for ch in children:
        om.append(mrun(ch) if isinstance(ch, str) else ch)
    return om


def para_math(*children, align=WD_ALIGN_PARAGRAPH.CENTER, before=4, after=4, label=None):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p._p.append(mobj(*children))
    if label:
        r = p.add_run('\t\t' + label)
        set_run_font(r, size=14)
    return p


def criterion_block():
    """ЭММ задачи: три критерия и система ограничений."""
    para('f1 = 22·x1 + 12·x2 + 14·x3 → max;', align=WD_ALIGN_PARAGRAPH.CENTER)
    para('f2 = 24·x1 + 16·x2 + 30·x3 → max;', align=WD_ALIGN_PARAGRAPH.CENTER)
    para('f3 = 20·x1 + 10·x2 + 15·x3 → min;', align=WD_ALIGN_PARAGRAPH.CENTER)


def res_constraints(extra=(), min3=30):
    para('2·x1 + 4·x2 + 4·x3 ≤ 400;', align=WD_ALIGN_PARAGRAPH.CENTER)
    para('3·x1 + 2·x2 + 2·x3 ≤ 300;', align=WD_ALIGN_PARAGRAPH.CENTER)
    para('4·x1 + 5·x2 + 3·x3 ≤ 500;', align=WD_ALIGN_PARAGRAPH.CENTER)
    if min3:
        para(f'x3 ≥ {min3};', align=WD_ALIGN_PARAGRAPH.CENTER)
    for line in extra:
        para(line, align=WD_ALIGN_PARAGRAPH.CENTER)
    para('x1 ≥ 0, x2 ≥ 0, x3 ≥ 0.', align=WD_ALIGN_PARAGRAPH.CENTER)


# ------------------------------- таблицы -------------------------------------
def table(headers, rows, widths_cm, font_size=12, header_size=12):
    t = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    t.allow_autofit = False
    tblPr = t._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement('w:' + edge)
        el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '0'); el.set(qn('w:color'), '000000')
        borders.append(el)
    tblPr.append(borders)
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed'); tblPr.append(layout)
    for i, w in enumerate(widths_cm):
        for cell in t.columns[i].cells:
            cell.width = Cm(w)
    for ci, h in enumerate(headers):
        cell = t.rows[0].cells[ci]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(str(h))
        set_run_font(r, size=header_size, bold=True)
        cell.vertical_alignment = 1
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = t.rows[ri].cells[ci]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(str(val))
            set_run_font(r, size=font_size)
            cell.vertical_alignment = 1
    return t


def caption(text):
    return para(text, align=WD_ALIGN_PARAGRAPH.CENTER, size=12, before=4, after=10)


def table_caption(text):
    return para(text, align=WD_ALIGN_PARAGRAPH.LEFT, size=12, before=8, after=4)


FIG = [0]


def figure(fname, text, width_cm=16.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    FIG[0] += 1
    n = FIG[0]
    run = p.add_run()
    run.add_picture(os.path.join(ASSETS, fname), width=Cm(width_cm))
    para(f'Рисунок 2.{n} – {text}', align=WD_ALIGN_PARAGRAPH.CENTER, size=12, before=2, after=10)
    return n


# ================================ ТИТУЛ =======================================
for line, size, bold, before in (
        ('МИНИСТЕРСТВО ОБРАЗОВАНИЯ РЕСПУБЛИКИ БЕЛАРУСЬ', 13, False, 0),
        ('УО «ПОЛЕССКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ»', 13, False, 2),
        ('', 13, False, 0),
        ('Инженерный факультет', 13, False, 100),
        ('Кафедра информационных технологий и интеллектуальных систем', 13, False, 2),
):
    para(line, align=WD_ALIGN_PARAGRAPH.CENTER, size=size, bold=bold, before=before, after=0)

para('', after=0)
para('ОТЧЁТ', align=WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True, before=40, after=6)
para('УСР №2', align=WD_ALIGN_PARAGRAPH.CENTER, size=14, bold=True, after=10)
para('по дисциплине: «Эконометрика и исследование операций в экономике»',
     align=WD_ALIGN_PARAGRAPH.CENTER, size=14, after=10)
para('на тему: «Модели и методы многоцелевой оптимизации»',
     align=WD_ALIGN_PARAGRAPH.CENTER, size=14, after=6)
para('Задание 2.4', align=WD_ALIGN_PARAGRAPH.CENTER, size=14, after=140)
SIGN_LEFT = Emu(-180340)


def sign_line(text, after=0):
    p = para(text, align=WD_ALIGN_PARAGRAPH.LEFT, size=14, after=after)
    p.paragraph_format.left_indent = SIGN_LEFT
    return p


sign_line('Выполнил:')
sign_line('студент')
sign_line('группы 24ИТ-3' + ' ' * 55 + 'Пась А.М.', after=6)
sign_line('Проверил:')
sign_line('' + ' ' * 70 + 'Свитина', after=140)
para('ПИНСК 2026', align=WD_ALIGN_PARAGRAPH.CENTER, size=14, after=0)

# ================================ ЗАДАЧА ======================================
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
heading('ЗАДАЧА')
para('Задание 2.4. Для производства трех видов продукции предприятие использует три '
     'вида ресурсов. Данные представлены в табл. 1:', first_indent=1.25, after=8)
table_caption('Таблица 1 – Исходные данные задачи 2.4')
t1 = table(['Ресурсы', 'Расход ресурсов на единицу продукции', '', '', 'Наличие ресурсов'],
           [['', '1П', '2П', '3П', ''],
            ['Ресурс 1', '2', '4', '4', '400'],
            ['Ресурс 2', '3', '2', '2', '300'],
            ['Ресурс 3', '4', '5', '3', '500']],
           [4.2, 2.6, 2.6, 2.6, 4.0], font_size=13)
t1.rows[0].cells[1].merge(t1.rows[0].cells[3])          # «Расход ресурсов…» на 3 столбца
t1.rows[0].cells[4].merge(t1.rows[1].cells[4])          # «Наличие ресурсов» на 2 строки
t1.rows[0].cells[0].merge(t1.rows[1].cells[0])          # «Ресурсы» на 2 строки
para('', after=6)
para('Единица продукции каждого вида характеризуется следующими показателями: '
     'прибылью – (22, 12, 14), оптовой ценой – (24, 16, 30), трудоемкостью – '
     '(20, 10, 15). Третьего вида продукции должно быть произведено не менее 30 ед. '
     'Требуется найти решение задачи:', after=6)
para('а) методом последовательных уступок при следующей важности критериев: прибыль – '
     'уступка 580, оптовая цена – уступка 1090, трудоемкость;', first_indent=1.25)
para('б) методом ведущего критерия: в качестве самого важного критерия считать прибыль, '
     'нижняя граница оптовой цены составляет 2400, трудоемкость не должна превышать 1600;',
     first_indent=1.25)
para('в) методом равных и наименьших относительных отклонений при объеме производства '
     'продукции первого вида не меньше 20 единиц;', first_indent=1.25)
para('г) методом минимакса при объеме производства продукции третьего вида не меньше 40 '
     'единиц.', first_indent=1.25, after=8)

# ================================ РЕШЕНИЕ =====================================
heading('РЕШЕНИЕ')
para('Обозначим через x1, x2, x3 объемы производства продукции видов 1П, 2П и 3П '
     'соответственно. Тогда для всех пунктов задачи экономико-математическая модель '
     '(ЭММ) имеет вид: три критерия',
     first_indent=1.25, after=6)
criterion_block()
para('при ограничениях по ресурсам и по объёму производства продукции 3П:',
     first_indent=1.25, after=6)
res_constraints()
para('Во всех пунктах f1 (прибыль) и f2 (оптовая цена) максимизируются, f3 '
     '(трудоемкость) минимизируется. Решение выполнено в Excel средствами надстройки '
     '«Поиск решения» (Simplex LP); ниже приведены кадры листов с полученными '
     'оптимальными планами.', first_indent=1.25, after=10)

# ------------------------------- а) ------------------------------------------
para('а) Решение методом последовательных уступок.', bold=True, after=6)
para('Критерии упорядочены по важности: f1 (прибыль) → f2 (оптовая цена) → f3 '
     '(трудоемкость). По первым двум критериям заданы уступки: Δf1 = 580, '
     'Δf2 = 1090.', first_indent=1.25, after=6)
para('Шаг 1. Решаем задачу по наиболее важному критерию f1 → max (рисунок 2.1).',
     first_indent=1.25, after=4)
figure('mmolo_a1.png', 'Поиск оптимального решения по критерию I (f1 → max)')
para(f'Получим X1* = ({ru(a["step1"]["x"][0])}; {ru(a["step1"]["x"][1])}; '
     f'{ru(a["step1"]["x"][2])}), при этом f1* = {ru(a["step1"]["f1"])}, '
     f'f2 = {ru(a["step1"]["f2"])}, f3 = {ru(a["step1"]["f3"])}.',
     first_indent=1.25, after=6)
para('Шаг 2. Так как f1 → max, вводим ограничение по уступке f1 ≥ f1* − Δf1 = '
     f'{ru(a["step1"]["f1"])} − 580 = {ru(a["cons1"])} и решаем задачу по критерию '
     'f2 → max (рисунок 2.2):', first_indent=1.25, after=4)
formula('22·x1 + 12·x2 + 14·x3 ≥ 1600        (1)')
figure('mmolo_a2.png', 'Поиск оптимального решения по критерию II (f2 → max) с ограничением (1)')
para(f'Получим X2* = ({ru(a["step2"]["x"][0])}; {ru(a["step2"]["x"][1])}; '
     f'{ru(a["step2"]["x"][2])}), при этом f1 = {ru(a["step2"]["f1"])}, '
     f'f2* = {ru(a["step2"]["f2"])}, f3 = {ru(a["step2"]["f3"])}.',
     first_indent=1.25, after=6)
para('Шаг 3. По критерию f2 назначаем уступку: f2 ≥ f2* − Δf2 = '
     f'{ru(a["step2"]["f2"])} − 1090 = {ru(a["cons2"])}. Решаем задачу по наименее '
     'важному критерию f3 → min с ограничениями (1) и (2) (рисунок 2.3):',
     first_indent=1.25, after=4)
formula('24·x1 + 16·x2 + 30·x3 ≥ 2360        (2)')
figure('mmolo_a3.png', 'Поиск оптимального решения по критерию III (f3 → min) с ограничениями (1), (2)')
x3a = a['step3']['x']
para(f'Компромиссное решение: X* = ({ru(x3a[0])}; {ru(x3a[1])}; {ru(x3a[2])}); '
     f'f1 = {ru(a["step3"]["f1"])}, f2 = {ru(a["step3"]["f2"])}, '
     f'f3* = {ru(a["step3"]["f3"])}.', first_indent=1.25, after=6)
table_caption('Таблица 2 – Реализация метода последовательных уступок')
table(['Шаг', 'Целевая функция', 'Доп. ограничения', 'x1', 'x2', 'x3', 'f1', 'f2', 'f3'],
      [['1', 'f1 → max', '—', ru(a['step1']['x'][0]), ru(a['step1']['x'][1]), ru(a['step1']['x'][2]),
        ru(a['step1']['f1']), ru(a['step1']['f2']), ru(a['step1']['f3'])],
       ['2', 'f2 → max', '(1)', ru(a['step2']['x'][0]), ru(a['step2']['x'][1]), ru(a['step2']['x'][2]),
        ru(a['step2']['f1']), ru(a['step2']['f2']), ru(a['step2']['f3'])],
       ['3', 'f3 → min', '(1), (2)', ru(x3a[0]), ru(x3a[1]), ru(x3a[2]),
        ru(a['step3']['f1']), ru(a['step3']['f2']), ru(a['step3']['f3'])]],
      [1.3, 2.3, 2.2, 1.7, 1.8, 1.6, 1.9, 1.9, 2.1], font_size=11, header_size=11)
para('', after=6)
para('Таким образом, чтобы условия задачи были выполнены, предприятие должно '
     f'произвести {ru(x3a[0])} ед. продукции 1П, {ru(x3a[1])} ед. 2П и {ru(x3a[2])} ед. 3П; '
     f'прибыль составит {ru(a["step3"]["f1"])} ден. ед., оптовая цена – {ru(a["step3"]["f2"])} '
     f'ден. ед., трудоемкость – {ru(a["step3"]["f3"])} чел.-ч. Значения приоритетных '
     'критериев находятся в пределах назначенных уступок, задача решена.',
     first_indent=1.25, after=10)

# ------------------------------- б) ------------------------------------------
para('б) Решение методом ведущего критерия.', bold=True, after=6)
para('По условию самым важным является критерий f1 (прибыль), поэтому остальные '
     'критерии переводятся в разряд ограничений: f2 ≥ 2400 и f3 ≤ 1600 (рисунок 2.4).',
     first_indent=1.25, after=4)
figure('mmolo_b1.png', 'Решение задачи методом ведущего критерия (f1 → max)')
xb = b['sol']['x']
para(f'Результат: X* = ({ru(xb[0])}; {ru(xb[1])}; {ru(xb[2])}), f1* = {ru(b["sol"]["f1"])} '
     f'ден. ед.; при этом f2 = {ru(b["sol"]["f2"])} ден. ед. и f3 = {ru(b["sol"]["f3"])} '
     'чел.-ч, т.е. обе границы связаны (2500 ≥ 2400; 1600 ≤ 1600).', first_indent=1.25, after=6)
para('Проверим принадлежность решения области компромиссов: при закрепленном '
     f'f1 = {ru(b["sol"]["f1"])} максимально возможное значение f2 равно '
     f'{ru(b["eff"]["f2"]["val"])}, а минимально возможное значение f3 равно '
     f'{ru(b["eff"]["f3"]["val"])} — оба совпадают с полученными, следовательно план '
     'эффективен (оптимален по Парето) и является компромиссным.',
     first_indent=1.25, after=10)

# ------------------------------- в) ------------------------------------------
para('в) Решение методом равных (и наименьших) относительных отклонений.', bold=True, after=6)
para('Дополнительное условие пункта: объем производства продукции первого вида не '
     'менее 20 единиц, т.е. x1 ≥ 20. Сначала решаем задачу по каждому критерию '
     'отдельно.', first_indent=1.25, after=4)
figure('mmolo_c1.png', 'Оптимальное решение по критерию I (f1 → max)')
para(f'X1* = ({ru(c["f1"]["x"][0])}; {ru(c["f1"]["x"][1])}; {ru(c["f1"]["x"][2])}), '
     f'f1* = {ru(c["F"][0])}.', first_indent=1.25, after=4)
figure('mmolo_c2.png', 'Оптимальное решение по критерию II (f2 → max)')
para(f'X2* = ({ru(c["f2"]["x"][0])}; {ru(c["f2"]["x"][1])}; {ru(c["f2"]["x"][2])}), '
     f'f2* = {ru(c["F"][1])}.', first_indent=1.25, after=4)
figure('mmolo_c3.png', 'Оптимальное решение по критерию III (f3 → min)')
para(f'X3* = ({ru(c["f3"]["x"][0])}; {ru(c["f3"]["x"][1])}; {ru(c["f3"]["x"][2])}), '
     f'f3* = {ru(c["F"][2])}.', first_indent=1.25, after=6)
para('Запишем условие равенства относительных отклонений критериев от их экстремальных '
     'значений:', first_indent=1.25, after=4)
formula('(f1 − f1*) / f1* = (f2 − f2*) / f2* = (f3 − f3*) / f3*')
para('Критерии f1 и f2 имеют одинаковое направление оптимизации, поэтому для них '
     'справедливо d1·f1 − d2·f2 = 0, а для пары f1 (max) и f3 (min) направления '
     'противоположны, откуда d1·f1 + d3·f3 = 2, где dк = 1/fк*. С подстановкой '
     'экстремальных значений получаем два дополнительных ограничения:',
     first_indent=1.25, after=4)
para_math(mfrac(mobj(msub('f', '1')), mobj('2180')),
          '−',
          mfrac(mobj(msub('f', '2')), mobj('3450')), '=', mobj('0'), label='(2)')
para_math(mfrac(mobj(msub('f', '1')), mobj('2180')),
          '+',
          mfrac(mobj(msub('f', '3')), mobj('850')), '=', mobj('2'), label='(3)')
para('После подстановки коэффициентов ограничения (2) и (3) принимают вид:',
     first_indent=1.25, after=4)
para('0,0031·x1 + 0,0009·x2 − 0,0023·x3 = 0        (2)\u2032', align=WD_ALIGN_PARAGRAPH.CENTER)
para('0,0336·x1 + 0,0173·x2 + 0,0241·x3 = 2        (3)\u2032', align=WD_ALIGN_PARAGRAPH.CENTER)
para('В качестве целевой функции берем f1 → max с ограничениями (2)′, (3)′ и исходной '
     'системой (рисунок 2.7):', first_indent=1.25, after=4)
figure('mmolo_c4.png', 'Компромиссное решение с дополнительными ограничениями (2)′, (3)′')
xw = c['sol']['x']
para(f'X* = ({ru(xw[0])}; {ru(xw[1])}; {ru(xw[2])}); f1 = {ru(c["sol"]["f1"])}, '
     f'f2 = {ru(c["sol"]["f2"])}, f3 = {ru(c["sol"]["f3"])}.', first_indent=1.25, after=6)
para('Проверка равенства относительных отклонений:', first_indent=1.25, after=4)
devs = c['dev']
for k, (val, ext) in enumerate(zip([c['sol']['f1'], c['sol']['f2'], c['sol']['f3']], c['F']), start=1):
    para(f'({ru(val)} − {ru(ext)}) / {ru(ext)} = −{ru(devs[k - 1] * 100, 2)} %'
         if k < 3 else f'({ru(val)} − {ru(ext)}) / {ru(ext)} = +{ru(devs[k - 1] * 100, 2)} %',
         align=WD_ALIGN_PARAGRAPH.CENTER)
para('Отклонения всех трёх критериев равны по модулю (42,14 %), значит найденное '
     f'решение компромиссное. Предприятие должно выпустить {ru(xw[0])} ед. продукции 1П, '
     f'{ru(xw[1])} ед. 2П и {ru(xw[2])} ед. 3П при прибыли {ru(c["sol"]["f1"])} ден. ед., '
     f'оптовой цене {ru(c["sol"]["f2"])} ден. ед. и трудоемкости {ru(c["sol"]["f3"])} чел.-ч.',
     first_indent=1.25, after=10)

# ------------------------------- г) ------------------------------------------
para('г) Решение методом минимакса.', bold=True, after=6)
para('По условию пункта третьего вида продукции должно быть произведено не менее 40 '
     'ед., т.е. x3 ≥ 40. Решаем задачу по каждому критерию отдельно (рисунки 2.8–2.10).',
     first_indent=1.25, after=4)
figure('mmolo_d1.png', 'Оптимальное решение по критерию I (f1 → max) при x3 ≥ 40')
para(f'X1* = ({ru(d["f1"]["x"][0])}; {ru(d["f1"]["x"][1])}; {ru(d["f1"]["x"][2])}), '
     f'f1* = {ru(d["G"][0])}.', first_indent=1.25, after=4)
figure('mmolo_d2.png', 'Оптимальное решение по критерию II (f2 → max) при x3 ≥ 40')
para(f'X2* = ({ru(d["f2"]["x"][0])}; {ru(d["f2"]["x"][1])}; {ru(d["f2"]["x"][2])}), '
     f'f2* = {ru(d["G"][1])}.', first_indent=1.25, after=4)
figure('mmolo_d3.png', 'Оптимальное решение по критерию III (f3 → min) при x3 ≥ 40')
para(f'X3* = ({ru(d["f3"]["x"][0])}; {ru(d["f3"]["x"][1])}; {ru(d["f3"]["x"][2])}), '
     f'f3* = {ru(d["G"][2])}.', first_indent=1.25, after=6)
para('За начальное компромиссное решение примем план, полученный по критерию f2 '
     '(наибольшие отклонения остальных критериев). Относительные отклонения от '
     'экстремальных значений:', first_indent=1.25, after=4)
for lab, val, ext in (('f1', d['naive']['f1'], d['G'][0]), ('f2', d['naive']['f2'], d['G'][1]),
                      ('f3', d['naive']['f3'], d['G'][2])):
    dv = abs(val - ext) / ext
    para(f'({lab}): |{ru(val)} − {ru(ext)}| / {ru(ext)} = {ru(dv * 100, 2)} %',
         align=WD_ALIGN_PARAGRAPH.CENTER)
para('Наибольшее отклонение — по критерию f3 (254,17 %). Минимизируем максимальное '
     'относительное отклонение: вводим переменную x4 = y, для максимизируемых '
     'критериев получаем ограничения вида fк ≥ fк*·(1 − x4), для минимизируемого — '
     'f3 ≤ f3*·(1 + x4); целевая функция x4 → min (рисунок 2.11).', first_indent=1.25, after=4)
figure('mmolo_d4.png', 'Решение расширенной минимаксной задачи (x4 → min)', width_cm=16.0)
xg = d['sol']['x']
para(f'X* = ({ru(xg[0])}; {ru(xg[1])}; {ru(xg[2])}), x4 = y = {ru(d["y"], 4)}. '
     f'Значения критериев: f1 = {ru(d["sol"]["f1"])}, f2 = {ru(d["sol"]["f2"])}, '
     f'f3 = {ru(d["sol"]["f3"])}.', first_indent=1.25, after=6)
para('Относительные отклонения в найденном компромиссном решении:', first_indent=1.25, after=4)
for lab, val, ext in (('f1', d['sol']['f1'], d['G'][0]), ('f2', d['sol']['f2'], d['G'][1]),
                      ('f3', d['sol']['f3'], d['G'][2])):
    para(f'({lab}): {ru(abs(val - ext) / ext * 100, 2)} %', align=WD_ALIGN_PARAGRAPH.CENTER)
para('Максимальное относительное отклонение уменьшилось с 254,17 % до 55,76 %, а '
     'отклонения по критериям f1 и f3 выровнены — решение компромиссное. '
     f'Следовательно, предприятию следует произвести {ru(xg[0])} ед. продукции 1П, '
     f'{ru(xg[1])} ед. 2П и {ru(xg[2])} ед. 3П; прибыль составит {ru(d["sol"]["f1"])} '
     f'ден. ед., оптовая цена – {ru(d["sol"]["f2"])} ден. ед., трудоемкость – '
     f'{ru(d["sol"]["f3"])} чел.-ч.', first_indent=1.25, after=8)

table_caption('Таблица 3 – Сравнение компромиссных решений по методам')
table(['Метод', 'x1', 'x2', 'x3', 'f1 (прибыль)', 'f2 (цена)', 'f3 (трудоемкость)'],
      [['а) последовательных уступок', ru(x3a[0]), ru(x3a[1]), ru(x3a[2]),
        ru(a['step3']['f1']), ru(a['step3']['f2']), ru(a['step3']['f3'])],
       ['б) ведущего критерия', ru(xb[0]), ru(xb[1]), ru(xb[2]),
        ru(b['sol']['f1']), ru(b['sol']['f2']), ru(b['sol']['f3'])],
       ['в) равных относительных отклонений', ru(xw[0]), ru(xw[1]), ru(xw[2]),
        ru(c['sol']['f1']), ru(c['sol']['f2']), ru(c['sol']['f3'])],
       ['г) минимакса', ru(xg[0]), ru(xg[1]), ru(xg[2]),
        ru(d['sol']['f1']), ru(d['sol']['f2']), ru(d['sol']['f3'])]],
      [4.6, 1.6, 1.8, 1.8, 2.2, 2.0, 2.4], font_size=11, header_size=11)

# ============================ ПРИЛОЖЕНИЕ А =====================================
heading('ПРИЛОЖЕНИЕ А', page_break=True)
para('Excel-модели и результаты «Поиска решения» (файл «Эконометрика УСР 2 (Пась А.М.) '
     '— задание 2.4.xlsx»). На каждом листе последовательно приведены все состояния '
     'модели, соответствующие шагам решения.', first_indent=1.25, after=8)


def appendix_figure(fname, text, width_cm):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    run = p.add_run()
    run.add_picture(os.path.join(ASSETS, fname), width=Cm(width_cm))
    para(text, align=WD_ALIGN_PARAGRAPH.CENTER, size=12, before=2, after=10)


appendix_figure('excel_mmolo_24а.png', 'Рисунок А.1 – Лист 2.4а: метод последовательных уступок', 16.5)
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
appendix_figure('excel_mmolo_24в.png', 'Рисунок А.2 – Лист 2.4в: метод равных относительных отклонений', 16.5)
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
appendix_figure('excel_mmolo_24г.png', 'Рисунок А.3 – Лист 2.4г: метод минимакса', 16.5)
appendix_figure('excel_mmolo_24б.png', 'Рисунок А.4 – Лист 2.4б: метод ведущего критерия', 16.5)

# ============================ ПРИЛОЖЕНИЕ Б =====================================
heading('ПРИЛОЖЕНИЕ Б', page_break=True)
para('Конспект: модели и методы многоцелевой оптимизации (тема 2)', bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
para('2.1  Постановка задачи многоцелевой оптимизации. Многие экономико-управленческие '
     'задачи являются многоцелевыми: оптимальное решение по одному критерию может оказаться '
     'не лучшим по значениям других критериев. Множество критериев представляют векторной '
     'целевой функцией F(X) = (f1(X), f2(X), …, fK(X)); чтобы минимизировать частный '
     'критерий fк(X), достаточно максимизать −fк(X).', first_indent=1.25)
para('Задача многоцелевой оптимизации записывается как многоцелевая (векторная) задача '
     'математического программирования:', first_indent=1.25, after=4)
formula('max F(X) = (f1(X), f2(X), …, fK(X))        (2.1)')
formula('φi(X) ≥ bi,  i = 1, …, m        (2.2)')
formula('X ≥ 0        (2.3)')
para('План X2 не хуже плана X1, если fк(X2) ≥ fк(X1) для всех к = 1, …, K (2.4). Если хотя '
     'бы одно неравенство строгое, план X2 предпочтительнее X1. План оптимален по Парето '
     '(эффективен, компромиссен, субоптимален), если он допустим и не существует другого '
     'допустимого плана, для которого выполняются неравенства (2.4) и хотя бы для одного '
     'критерия — строгое неравенство.', first_indent=1.25)
para('При решении векторных задач решают три проблемы: нормализацию критериев '
     '(приведение к единому безразмерному масштабу: fк/fк* или (fк − fк*)/fк*), учет '
     'приоритета критериев и определение области компромисса.', first_indent=1.25)
para('2.2  Метод линейной комбинации частных критериев. Задан вектор весов '
     'α = (α1, …, αK), αк > 0, Σαк = 1; критерии в свертке обязательно нормированы, а '
     'многоцелевая задача становится одноцелевой: max F(X) = Σ αк·fк(X).',
     first_indent=1.25)
para('2.3  Метод последовательных уступок. Сначала устанавливается предпочтительность '
     'всех критериев, на первое место ставится самый важный. Находится оптимальное '
     'решение f1* и по нему назначается уступка Δf1. Если f1 → max, далее решается задача '
     'по второму критерию с дополнительным ограничением:', first_indent=1.25, after=4)
formula('f1 ≥ f1* − Δf1        (2.5)')
para('После нахождения f2* назначается уступка Δf2 и решается задача по третьему '
     'критерию с двумя дополнительными ограничениями; если f2 → min, ограничение имеет '
     'вид f2 ≤ f2* + Δf2. Решение принимают, если значения всех критериев устраивают; '
     'иначе величины уступок изменяют и задачу решают заново.', first_indent=1.25)
para('2.4  Метод ведущего критерия. Все критерии, кроме самого важного, переводятся в '
     'разряд ограничений:', first_indent=1.25, after=4)
formula('max f1(X)  при  fк(X) ≥ βк,  к = 2, …, K        (2.6)')
para('где βк — нижние границы (критерии минимизации предварительно умножают на −1). '
     'Полученное решение может не быть эффективным, поэтому проверяют его принадлежность '
     'области компромиссов.', first_indent=1.25)
para('2.5  Метод равных и наименьших относительных отклонений. Записывают условие '
     'равенства относительных отклонений критериев от их экстремальных значений:',
     first_indent=1.25, after=4)
para_math(mfrac(mobj(msub('f', '1'), '−', mobj(msub('f', '*'))), mobj(msub('f', '*'))),
          '=',
          mfrac(mobj(msub('f', '2'), '−', mobj(msub('f', '*'))), mobj(msub('f', '*'))),
          '=',
          mfrac(mobj(msub('f', '3'), '−', mobj(msub('f', '*'))), mobj(msub('f', '*'))),
          )
para('Обозначив dк = 1/fк*, для критериев с одинаковым направлением оптимизации '
     'получают линейное ограничение d1·f1 − d2·f2 = 0, а для критериев с '
     'противоположными направлениями — d1·f1 + d3·f3 = 2. Введённые ограничения '
     'добавляют в исходную систему, целевую функцию можно взять любую из fк.',
     first_indent=1.25)
para('2.6  Метод минимакса. Сначала задачу решают по каждому критерию отдельно и находят '
     'экстремальные значения fк*; предполагая, что X0 — компромиссное решение, записывают '
     'относительные отклонения:', first_indent=1.25, after=4)
formula('yk = (Σj ckj·x0j − fк*) / fк*,   k = 1, …, K        (2.7)')
para('Среди найденных значений берут наибольшее и требуют, чтобы в искомом компромиссном '
     'решении оно было минимальным: min max(yk). Заменяя отклонения наибольшей величиной '
     'x(n+1) = y, получают ограничения', first_indent=1.25, after=4)
formula('Σj ckj·xj + fк*·x(n+1) ≥ fк*   (для максимизируемых критериев)        (2.8)')
formula('Σj hкj·xj − fк*·x(n+1) ≤ fк*   (для минимизируемых критериев)        (2.9)')
para('Целевая функция расширенной задачи: f = x(n+1) → min. Метод повторяют, пока '
     'наибольшее относительное отклонение уменьшается.', first_indent=1.25, after=10)

# ------------------------------ колонтитул -----------------------------------
def add_field(paragraph, instr):
    r = paragraph.add_run()
    set_run_font(r, size=12)
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), instr)
    t = OxmlElement('w:r')
    tt = OxmlElement('w:t')
    tt.text = '1'
    t.append(tt)
    fld.append(t)
    paragraph._p.append(fld)


sec.different_first_page_header_footer = True
sec.footer.is_linked_to_previous = False
fp = sec.footer.paragraphs[0]
for r0 in list(fp.runs):
    r0._r.getparent().remove(r0._r)
for extra in list(sec.footer.paragraphs[1:]):
    extra._p.getparent().remove(extra._p)
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_field(fp, 'PAGE')

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print('saved', OUT)
