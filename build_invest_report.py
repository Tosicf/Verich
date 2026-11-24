# -*- coding: utf-8 -*-
"""Сборка отчёта УСР по теме 3 «Модели анализа и оценки инвестиционных проектов»,
задания 3.4, 3.6, 3.12 (в методичке напечатаны как 1.4, 1.6 и 3.12 — с. 107, 109).

Структура и оформление — по образцу «Эконометрика УСР 3.docx»:
ЗАДАЧА → РЕШЕНИЕ (по заданиям, внутри каждого — кадры листов Excel-решения) →
ПРИЛОЖЕНИЕ А (листы Excel-книги) → ПРИЛОЖЕНИЕ Б (методика темы).

Формулы — настоящие формулы Word (OMML, модуль omml.py): дроби, Σ с пределами,
индексы, системы с фигурной скобкой. Числа — из /tmp/invest_results.json
(расчёт solve_invest.py, сверен с формулами книги build_invest_excel.py).
"""
import json
import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
from omml import Math  # noqa: E402

ASSETS = os.path.join(ROOT, 'report_assets')
OUT = os.path.join(ROOT, 'outputs',
                   'Эконометрика УСР 3 (Пась А.М.) — задания 3.4, 3.6, 3.12.docx')
D = json.load(open('/tmp/invest_results.json'))
D34, D36, D312 = D['t34'], D['t36'], D['t312']


def ru(v, nd=4):
    """Число в русской записи: разделитель — запятая, лишние нули убираем."""
    if v is None:
        return '—'
    v = round(float(v), nd)
    if abs(v) < 0.5 * 10 ** (-nd):
        v = 0.0
    if abs(v - round(v)) < 1e-9:
        s = str(int(round(v)))
    else:
        s = f'{v:.{nd}f}'.rstrip('0').rstrip('.')
    return s.replace('.', ',').replace('-', '−')


def pc(v, nd=2):
    return ru(100 * v, nd) + ' %'


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

mm = Math(doc, size=14)


def set_run_font(run, size=14, bold=False, name='Times New Roman', italic=False):
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


def para(text='', align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=14, bold=False, italic=False,
         before=None, after=None, first_indent=None):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
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


def subhead(text):
    p = para(text, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, bold=True, before=12, after=8)
    p.paragraph_format.keep_with_next = True
    return p


# ------------------------------ формулы OMML ---------------------------------
def F(num, den):
    return mm.frac(num, den)


def s1(t, k):
    """t_k (k — число, строка или готовый OMML-элемент)."""
    return mm.sub(t, k)


def sp(t, k):
    """t^k."""
    return mm.sup(t, k)


def mpara(*items, label=None, before=6, after=6):
    mm.para(*items, label=label, align='center', before=before, after=after)


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


TAB = [0]


def table_caption(text):
    return para(text, align=WD_ALIGN_PARAGRAPH.LEFT, size=12, before=8, after=4)


def tbl(headers, rows, widths_cm, caption, font_size=12, header_size=12):
    TAB[0] += 1
    table_caption(f'Таблица {TAB[0]} – {caption}')
    table(headers, rows, widths_cm, font_size=font_size, header_size=header_size)
    para('', after=6)


FIG = [0]


def figure(fname, text, width_cm=None):
    """Картинка в естественном размере: 300 dpi, но не шире поля страницы."""
    from PIL import Image
    path = os.path.join(ASSETS, fname)
    if width_cm is None:
        px = Image.open(path).size[0]
        width_cm = min(16.5, max(6.0, px / 300 * 2.54))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    FIG[0] += 1
    n = FIG[0]
    run = p.add_run()
    run.add_picture(path, width=Cm(width_cm))
    para(f'Рисунок 3.{n} – {text}', align=WD_ALIGN_PARAGRAPH.CENTER, size=12,
         before=2, after=10)
    return n


# ================================ ТИТУЛ ======================================
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
para('УСР №3', align=WD_ALIGN_PARAGRAPH.CENTER, size=14, bold=True, after=10)
para('по дисциплине: «Эконометрика и исследование операций в экономике»',
     align=WD_ALIGN_PARAGRAPH.CENTER, size=14, after=10)
para('на тему: «Модели анализа и оценки инвестиционных проектов»',
     align=WD_ALIGN_PARAGRAPH.CENTER, size=14, after=6)
para('Задания 3.4, 3.6, 3.12', align=WD_ALIGN_PARAGRAPH.CENTER, size=14, after=140)
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

# ================================ ЗАДАЧА =====================================
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
heading('ЗАДАЧА')
para('Задания темы 3 «Модели анализа и оценки инвестиционных проектов» по сборнику '
     'задания ЭиЭММиМ (с. 107 и 109). Нумерация в сборнике сбита: задание 3.4 '
     'напечатано как 1.4, задание 3.6 — как 1.6, а задание 3.12 идёт под своим '
     'номером, поэтому дальше используется сквозная нумерация темы.',
     first_indent=1.25, after=10)

para('Задание 3.4 (в сборнике — 1.4).', bold=True, after=4)
para('Инвестор рассматривает проект, начальные инвестиции в который составляют '
     '150 ден. ед. Срок реализации проекта 2 года. Последовательность денежных '
     'потоков по годам – 150, 150 ден. ед. Ставка дисконтирования – 15 %.',
     first_indent=1.25)
para('Найти уровень инфляции, до которого проект эффективен.', first_indent=1.25,
     after=10)

para('Задание 3.6 (в сборнике — 1.6).', bold=True, after=4)
para('Целесообразно ли инвестировать в следующий проект: начальные инвестиции – '
     '800 млн ден. ед.; срок реализации – 5 лет; годовые денежные потоки – 300 млн '
     'ден. ед.; норма дисконтирования – 12 %; ожидаемый уровень инфляции – 10, 12, '
     '18 % годовых.', first_indent=1.25, after=10)

para('Задание 3.12.', bold=True, after=4)
para('Инвестор рассматривает шесть инвестиционных проектов, информация о которых '
     'приведена в таблице 1.', first_indent=1.25, after=4)
table_caption('Таблица 1 – Денежные потоки инвестиционных проектов, млн руб.')
rows_cf = []
for p in D312['projects']:
    cf = {int(t): v for t, v in p['cf'].items()}
    rows_cf.append([p['name']] + [ru(cf[t]) if t in cf else '—' for t in range(5)])
t1 = table(['№ проекта', 'год 0', 'год 1', 'год 2', 'год 3', 'год 4'], rows_cf,
           [2.6, 2.5, 2.5, 2.5, 2.5, 2.5], font_size=13, header_size=13)
TAB[0] = 1                      # «Таблица 1» уже описана вручную выше
para('В начале первого года инвестор располагает суммой в 100 млн руб. '
     'Взаимоисключающими являются первый и второй проекты, а также третий и пятый. '
     'В результате осуществления второго, третьего и пятого проектов выпускается '
     'однородная продукция в количествах 400, 110 и 150 единиц соответственно. Спрос '
     'на продукцию не превышает 550 единиц.', first_indent=1.25, before=8)
para('Определить оптимальный портфель инвестиций, если ставка дисконтирования – 15 %.',
     first_indent=1.25, after=8)

# =============================== РЕШЕНИЕ =====================================
heading('РЕШЕНИЕ')
para('Эффективность проекта оценивается по его свободному денежному потоку '
     '(поступления за вычетом затрат) и норме дисконтирования r. Воспользуемся '
     'показателями темы 3; нумерация формул — как в конспекте.', first_indent=1.25,
     after=6)
para('Текущая стоимость проекта — сумма дисконтированных потоков, чистая текущая '
     'стоимость — разность между ней и начальными инвестициями:', first_indent=1.25,
     after=4)
mpara('PV = ', mm.sum('k=1', 'n', F(s1('C', 'k'), sp('1 + r', 'k'))), label='(3.3)')
mpara('NPV = PV − ', s1('I', '0'), label='(3.4)')
para('Внутренняя норма прибыли IRR — ставка, при которой NPV = 0; индекс '
     'прибыльности PI; простой и дисконтированный сроки окупаемости PP и DPP:',
     first_indent=1.25, after=4)
mpara(s1('I', '0'), ' = ', mm.sum('k=1', 'n', F(s1('C', 'k'), sp('1 + IRR', 'k'))),
      label='(3.5)')
mpara('PI = ', F('PV', s1('I', '0')), label='(3.6)')
mpara('PP = ', mm.op('min n'), '  ', mm.op('при котором'), '  ',
      mm.sum('k=1', 'n', s1('C', 'k')), ' ≥ ', s1('I', '0'),
      ',    DPP = ', mm.op('min n'), '  ', mm.op('при котором'), '  ',
      mm.sum('k=1', 'n', F(s1('C', 'k'), sp('1 + r', 'k'))), ' ≥ ', s1('I', '0'),
      label='(3.7), (3.8)')
para('Модифицированная внутренняя норма доходности считается по наращённым '
     'положительным потокам:', first_indent=1.25, after=4)
mpara('MIRR = ', sp(F('FV', s1('I', '0')), F('1', 'n')), ' − 1,   ',
      'FV = ', mm.sum('k=1', 'n', [s1('C', 'k'), sp('1 + r', 'n − k')]))
para('Проект принимается, если NPV > 0 (эквивалентно PI > 1 и IRR > r). Инфляция '
     'учитывается переводом реальной ставки в номинальную:', first_indent=1.25,
     after=4)
mpara('r = ', s1('r', '0'), ' + i + ', s1('r', '0'), '·i', label='(3.9)')
para('Ниже по каждому заданию приведены кадры листов Excel-книги: в строке формул '
     'и в ячейках видно, какой расчёт выполняет книга, справа — результат. Полный '
     'снимок листа вынесен в ПРИЛОЖЕНИЕ А.', first_indent=1.25, after=10)

# ---------------------------- задание 3.4 ------------------------------------
subhead('Задание 3.4. Уровень инфляции, до которого проект эффективен')
para('Исходные данные задания:', first_indent=1.25, after=4)
mpara(s1('I', '0'), ' = 150,   n = 2,   ', s1('C', '1'), ' = 150,   ',
      s1('C', '2'), ' = 150,   ', s1('r', '0'), ' = 0,15')
figure('inv_34_data.png', 'Начальные данные проекта на листе 3.4 книги Excel')
para('Шаг 1. Базовая оценка проекта без учёта инфляции (r = r0 = 0,15). По '
     'формулам (3.3)–(3.4):', first_indent=1.25, after=4)
mpara('PV = ', F('150', '1,15'), ' + ', F('150', sp('1,15', '2')), ' = ',
      ru(D34['base']['pv']), ';    NPV = ', ru(D34['base']['pv']), ' − 150 = ',
      ru(D34['base']['npv']), label='(1)')
b = D34['base']
para(f'На листе 3.4 (рисунок 3.2) кроме NPV получены PI = {ru(b["pi"])}, '
     f'IRR = {pc(b["irr"])}, PP = {b["pp"]}, DPP = {b["dpp"]}, '
     f'MIRR = {pc(b["mirr"])}. Все критерии согласованны: NPV > 0, PI > 1, '
     'IRR > r, значит при нулевой инфляции проект эффективен.',
     first_indent=1.25, after=4)
figure('inv_34_base.png', 'Расчёт показателей эффективности проекта (формулы в ячейках)')
para('Шаг 2. Введём инфляцию. Номинальная ставка дисконтирования зависит от уровня '
     'инфляции i по формуле (3.9), поэтому чистая текущая стоимость проекта — '
     'функция i:', first_indent=1.25, after=4)
mpara('NPV(i) = ', F('150', '1 + r'), ' + ', F('150', sp('1 + r', '2')), ' − 150,   ',
      'r = 0,15 + 1,15·i', label='(2)')
para('Граница эффективности — корень уравнения NPV(i) = 0. Обозначив x = 1 + r и '
     'умножив обе части на x², получаем квадратное уравнение:', first_indent=1.25,
     after=4)
mpara('150x + 150 = 150', sp('x', '2'), '   ⇔   ', sp('x', '2'), ' − x − 1 = 0',
      label='(3)')
para('Его положительный корень x = (1 + √5)/2 = 1,618034 (золотое сечение), откуда '
     'r = 0,618034 — это в точности IRR проекта, что и ожидалось: NPV обращается в '
     'нуль при ставке, равной внутренней норме прибыли. Тогда из (3.9):',
     first_indent=1.25, after=4)
mpara('i = ', F(['r − ', s1('r', '0')], ['1 + ', s1('r', '0')]), ' = ',
      F('0,618034 − 0,15', '1,15'), ' = ', ru(D34['i_max'], 6), ' = ',
      pc(D34['i_max'], 4), label='(4)')
para('Шаг 3. Проверка перебором: на листе 3.4 ниже базового расчёта построена '
     'таблица вариантов i = 0; 5; 10; 20; 30; 40; 50 % и критическое i = 40,6986 % '
     '(формула (2) продублирована '
     'в ячейках C…G) и построен график NPV(i) — рисунки 3.3 и 3.4.', first_indent=1.25,
     after=4)
figure('inv_34_infl.png', 'Варианты расчёта NPV при различных уровнях инфляции')
figure('inv_34_curve.png', 'Зависимость NPV проекта от уровня инфляции')
rows34 = [[f'вариант {k + 1}', ru(100 * z['i'], 4), ru(100 * z['d'], 4),
           ru(z['pv1']), ru(z['pv2']),
           ru(z['pv']), ru(z['npv']), ru(z['pi']),
           'эффективен' if z['npv'] > 0.5 else ('граница' if abs(z['npv']) < 0.5
                                                else 'неэффективен')]
          for k, z in enumerate(D34['rows'])]
tbl(['Вариант', 'i, %', 'r номинальная, %', 'PV потока 1', 'PV потока 2',
     'PV проекта', 'NPV', 'PI', 'Вывод'], rows34,
    [2.0, 1.5, 2.1, 1.9, 1.9, 1.9, 1.7, 1.3, 2.4],
    'Расчёт задания 3.4 по уровням инфляции', font_size=11)
para('Ответ: проект эффективен при уровне инфляции i ≤ 40,70 %. Критическое '
     'значение i = 0,4070 (40,70 %): при нём номинальная ставка дисконтирования '
     f'вырастет до {pc(D34["r_at_imax"])} и NPV = 0; при большей инфляции проект '
     'неэффективен, например при i = 50 % NPV = '
     f'{ru(D34["rows"][-1]["npv"])}.', first_indent=1.25, after=10)

# ---------------------------- задание 3.6 ------------------------------------
subhead('Задание 3.6. Целесообразность инвестиций при разных уровнях инфляции')
para('Данные задания:', first_indent=1.25, after=4)
mpara(s1('I', '0'), ' = 800,   n = 5,   ', s1('C', 'k'), ' = 300,  k = 1, …, 5,   ',
      s1('r', '0'), ' = 0,12,   i ∈ {0,10; 0,12; 0,18}')
figure('inv_36_data.png', 'Начальные данные проекта на листе 3.6 книги Excel')
para('Шаг 1. Расчёт без учёта инфляции. Все годовые потоки одинаковы — это '
     'финансовая рента, поэтому из (3.3) и (3.4):', first_indent=1.25, after=4)
mpara('PV = ', mm.sum('k=1', '5', F('300', sp('1,12', 'k'))), ' = ',
      ru(D36['base']['pv']), ';   NPV = ', ru(D36['base']['pv']), ' − 800 = ',
      ru(D36['base']['npv']), label='(5)')
b = D36['base']
para(f'На листе 3.6 (рисунок 3.6): PI = {ru(b["pi"])}, IRR = {pc(b["irr"])}, '
     f'MIRR = {pc(b["mirr"])}, PP = {b["pp"]}, DPP = {b["dpp"]}. '
     'Без инфляции проект явно привлекателен: NPV > 0, PI > 1, IRR > r, а '
     'дисконтированный срок окупаемости на год больше простого — следствие '
     'дисконтирования потоков.', first_indent=1.25, after=4)
figure('inv_36_base.png', 'Показатели эффективности проекта без учёта инфляции')
para('Шаг 2. Оценка с поправкой на инфляцию. По (3.9) номинальная ставка равна '
     'r = 0,12 + i + 0,12·i, и для каждого ожидаемого уровня i чистая текущая '
     'стоимость пересчитывается:', first_indent=1.25, after=4)
mpara('NPV(i) = ', mm.sum('k=1', '5', F('300', sp('1 + r', 'k'))), ' − 800,   ',
      'r = 0,12 + 1,12·i', label='(6)')
r0 = D36['rows'][0]
para(f'Для i = 0,10 номинальная ставка r = 0,12 + 0,10 + 0,12·0,10 = '
     f'{ru(r0["d"], 4)}, откуда PV = {ru(r0["pv"])} и NPV = {ru(r0["npv"])} > 0 — '
     'инвестировать целесообразно. Остальные сценарии рассчитаны на листе так же '
     '(рисунок 3.7), а результат сведён в таблицу 3.', first_indent=1.25, after=4)
rows6 = [[('критический (NPV = 0)' if z['critical'] else f'i = {ru(100 * z["i"], 0)} %'),
          ru(100 * z['d'], 2), ru(z['pv']), ru(z['npv']), ru(z['pi']),
          f'{int(z["pp"])} / ' + ('не достигается' if z['dpp'] is None
                                  else str(int(z['dpp']))),
          'инвестировать можно' if z['npv'] > 0.5 else
          ('проект на границе' if z['critical'] else 'инвестировать нельзя')]
         for z in D36['rows']]
tbl(['Сценарий', 'r номинальная, %', 'PV', 'NPV', 'PI', 'PP / DPP, лет', 'Вывод'],
    rows6, [3.0, 2.4, 2.0, 1.9, 1.4, 2.4, 3.9],
    'Оценка проекта задания 3.6 при различных уровнях инфляции', font_size=12)
figure('inv_36_infl.png', 'Расчёт сценариев по уровням инфляции и критическое i')
figure('inv_36_curve.png', 'Зависимость NPV проекта от уровня инфляции (задание 3.6)')
para(f'Критический уровень инфляции, при котором NPV = 0, найден на листе '
     f'подбором: i = {ru(D36["i_crit"], 6)} ({pc(D36["i_crit"])}). Он согласуется с '
     f'IRR проекта: 1,12·(1 + i) = 1 + IRR = {ru(1 + D36["base"]["irr"])}, '
     'то есть инфляция «съедает» всю доходность проекта при 11,98 % годовых.',
     first_indent=1.25, after=4)
para('Ответ: инвестировать целесообразно только при инфляции 10 % (NPV = '
     f'{ru(r0["npv"])} млн ден. ед.). При инфляции 12 % проект '
     f'находится на границе и даёт NPV = {ru(D36["rows"][1]["npv"])} < 0, а при '
     f'инфляции 18 % NPV = {ru(D36["rows"][2]["npv"])} < 0 — в этих двух случаях '
     'вложения в проект нецелесообразны (проще положить средства в банк).',
     first_indent=1.25, after=10)

# ---------------------------- задание 3.12 -----------------------------------
subhead('Задание 3.12. Оптимальный портфель инвестиционных проектов')
para('Шаг 1. Считаем NPV каждого проекта по (3.3)–(3.4) при ставке r = 0,15. '
     'Множители дисконтирования 1/(1,15)^t вынесены в строку 9 листа 3.12 '
     '(рисунок 3.9): 1,0000; 0,8696; 0,7561; 0,6575; 0,5718.',
     first_indent=1.25, after=4)
for p in D312['projects']:
    cf = {int(t): v for t, v in p['cf'].items()}
    items = []
    for j, (t, v) in enumerate(sorted(cf.items())):
        if j:
            items.append(' + ' if v > 0 else ' − ')
        elif v < 0:
            items.append('− ')
        items.append(ru(abs(v)))
        if t:                                  # множитель года 0 равен единице
            items.append('·')
            items.append(F('1', sp('1,15', t)))
    mpara(s1('NPV', p['name']), ' = ', *items, ' = ', ru(p['npv']))
para('У проектов 1 и 4 графа года 0 в таблице условия пуста: их единственный отток '
     '(90 и 40 млн руб.) приходится на второй год. Для ограничения по бюджету '
     'используем величину этого оттока, так как решение об участии проекта в '
     'портфеле инвестор принимает в начале первого года.', first_indent=1.25,
     after=4)
figure('inv_312_cf.png',
       'Денежные потоки, NPV и характеристики проектов на листе 3.12')
figure('inv_312_npv.png', 'Сравнение NPV проектов при ставке дисконтирования 15 %')
para('Все шесть проектов по отдельности эффективны (NPV > 0), но бюджет '
     'позволяет реализовать не все — отсюда задача отбора. Экономико-математическая '
     'модель портфеля (пункт 3.4 конспекта): переменная x = 1, если проект '
     'реализуется, и x = 0, если отклоняется:', first_indent=1.25, after=4)
npvs = [p['npv'] for p in D312['projects']]
inv = [p['invest'] for p in D312['projects']]
units = [p['units'] for p in D312['projects']]
obj = []
for j, v in enumerate(npvs):
    if j:
        obj.append(' + ')
    obj.append(ru(v))
    obj.append(s1('x', j + 1))
mpara(mm.op('max '), 'NPV = ', *obj, label='(7)')
cons = []
bud = []
for j, v in enumerate(inv):
    if j:
        bud.append(' + ')
    bud.append(ru(v))
    bud.append(s1('x', j + 1))
cons.append([mm.op('бюджет:    '), *bud, ' ≤ 100'])
dem = []
for j, v in enumerate(units):
    if not v:
        continue
    if dem:
        dem.append(' + ')
    dem.append(ru(v))
    dem.append(s1('x', j + 1))
cons.append([mm.op('спрос:     '), *dem, ' ≤ 550'])
cons.append([mm.op('исключение:    '), s1('x', 1), ' + ', s1('x', 2), ' ≤ 1'])
cons.append([mm.op('исключение:    '), s1('x', 3), ' + ', s1('x', 5), ' ≤ 1'])
cons.append([s1('x', 'i'), ' ∈ {0; 1},   i = 1, …, 6'])
mm.para(mm.cases(*cons), align='center', before=4, after=6)
para('В Excel модель собрана так же, как в задании 2.4: в столбце J — значения '
     'переменных x, в столбце G — левые части ограничений формулами СУММПРОИЗВ, '
     'справа — знаки и правые части (рисунок 3.11).', first_indent=1.25, after=4)
figure('inv_312_model.png',
       'Целевая функция и проверка ограничений модели формулами книги')
para('Шаг 2. Решение. Так как проектов шесть, полный перебор даёт 2⁶ = 64 набора '
     'переменных; допустимыми (по бюджету, спросу и взаимоисключению) оказались '
     f'{len(D312["all_feas"])} из них. Они упорядочены по значению целевой функции '
     'в таблице 4 и приведены на листе (рисунок 3.12).', first_indent=1.25, after=4)
rows12 = [[('{' + '; '.join(z['names']) + '}'), ru(z['npv']), ru(z['spend']),
           ru(z['units']), 'оптимальный' if k == 0 else 'допустим']
          for k, z in enumerate(D312['all_feas'])]
tbl(['Портфель', 'Σ NPV — целевая функция', 'Σ I·x ≤ 100', 'Выпуск ≤ 550', 'Статус'],
    rows12, [2.6, 4.4, 2.9, 2.8, 3.0],
    'Допустимые варианты портфеля, упорядоченные по Σ NPV', font_size=12)
figure('inv_312_portfeli.png', 'Перебор допустимых портфелей на листе 3.12')
figure('inv_312_portfolios.png', 'Допустимые варианты портфеля (убывание Σ NPV)')
best = D312['best']
para('Шаг 3. Анализ решения. Оптимальный набор — X* = (0; 1; 0; 1; 0; 0), то есть '
     f'проекты 2 и 4: Σ NPV = {ru(best["npv"])} млн руб. при затратах '
     f'{ru(best["spend"])} млн руб. и выпуске {ru(best["units"])} ед. '
     f'Остаток бюджета {ru(100 - best["spend"])} млн руб. не позволяет добавить ни '
     'одного из оставшихся проектов: наименьшие вложения среди них — 40 млн руб. у '
     'проекта 3, а проект 6 требует все 100 млн руб. Ограничение по спросу активно не '
     'полностью: 400 ед. из 550.', first_indent=1.25, after=4)
para('Для сравнения: ближайший конкурирующий набор {2; 3} даёт Σ NPV = '
     f'{ru(D312["all_feas"][1]["npv"])}, а самый «сильный» по отдельности проект 6 — '
     f'{ru(npvs[5])} млн руб. Он хорош сам по себе, но «съедает» весь бюджет, поэтому '
     'два проекта 2 и 4 в сумме выгоднее на '
     f'{ru(best["npv"] - D312["all_feas"][2]["npv"])} млн руб.',
     first_indent=1.25, after=6)
para('Ответ: оптимальный портфель — проекты 2 и 4; доходность портфеля '
     f'Σ NPV = {ru(best["npv"])} млн руб., использовано {ru(best["spend"])} из '
     '100 млн руб. бюджета, выпуск однородной продукции 400 ед. при спросе '
     'не более 550 ед.', first_indent=1.25, after=10)

# ============================== ПРИЛОЖЕНИЕ А ==================================
heading('ПРИЛОЖЕНИЕ А', page_break=True)
para('Листы Excel-книги «Эконометрика УСР 3 (Пась А.М.) — задания 3.4, 3.6, '
     '3.12.xlsx». В ячейках с результатами записаны формулы: множитель '
     'дисконтирования 1/(1+r)^t, суммы по годам, СУММПРОИЗВ для левой части '
     'ограничений модели, — поэтому при изменении исходных данных книга '
     'пересчитывается автоматически.', first_indent=1.25, after=8)
figure('inv_full_34.png', 'Лист 3.4 — проект и критический уровень инфляции')
figure('inv_full_36.png', 'Лист 3.6 — показатели проекта и сценарии инфляции')
figure('inv_full_312.png', 'Лист 3.12 — модель отбора портфеля и перебор наборов')

# ============================== ПРИЛОЖЕНИЕ Б ==================================
heading('ПРИЛОЖЕНИЕ Б', page_break=True)
para('Краткие методические сведения по теме 3 (конспект ЭиЭММиМ, с. 20–27).',
     first_indent=1.25, after=8)
para('1. Операции наращения и дисконтирования. Если капитал P положен под '
     'процентную ставку r, то через n периодов наращенная сумма (будущая '
     'стоимость) равна', first_indent=1.25, after=4)
mpara('S = P', sp('(1 + r)', 'n'), ',    ', 'α = ', sp('(1 + r)', 'n'),
      '  — коэффициент аккумуляции', label='(3.1)')
para('Обратная операция — нахождение текущей стоимости суммы S:', first_indent=1.25,
     after=4)
mpara('PV = ', F('S', sp('(1 + r)', 'n')), ',    δ = ', sp('(1 + r)', '−n'),
      '  — коэффициент дисконтирования', label='(3.2)')
para('2. Свободный денежный поток проекта: C = П − I в каждом году k (поступления '
     'минус выплаты). По нему строятся показатели эффективности (3.3)–(3.9), '
     'приведённые в начале РЕШЕНИЯ: проект эффективен при NPV > 0, что равносильно '
     'PI > 1 и IRR > r; при NPV = 0 подключают другие критерии, например срок '
     'окупаемости. Всегда DPP > PP: дисконтированный поток «опоздает» окупить '
     'инвестиции.', first_indent=1.25, after=6)
para('3. Влияние инфляции. Если P0 и P1 — цены потребительской корзины в начале и '
     'в конце периода, то уровень инфляции i = (P1 − P0)/P0, а реальная ставка '
     'связана с номинальной так:', first_indent=1.25, after=4)
mpara('r̂ = ', F('1 + r', '1 + i'), ' − 1,    ', mm.op('откуда'), '    r = r̂ + i + r̂·i',
      label='(3.9)')
para('Поэтому рост ожидаемой инфляции увеличивает знаменатели всех множителей '
     'дисконтирования и снижает NPV: отсюда и «критический уровень инфляции», '
     'который найден в заданиях 3.4 и 3.6.', first_indent=1.25, after=6)
para('4. Модель оптимального портфеля инвестиционных проектов при ограниченных '
     'финансовых ресурсах (пункт 3.4 конспекта):', first_indent=1.25, after=4)
mm.para(mm.cases(
    [mm.op('max '), mm.sum('i=1', 'n', [s1('NPV', 'i'), s1('x', 'i')])],
    [mm.sum('i=1', 'n', [s1('I', '0i'), s1('x', 'i')]), ' ≤ ', s1('I', 'o')],
    [s1('x', 'i'), ' ∈ {0; 1},   i = 1, …, n'],
), align='center', before=4, after=6)
para('В конкретных задачах добавляются ограничения на выпуск однородной продукции '
     '(спрос) и попарное взаимоисключение проектов — так получилось в задании 3.12. '
     'При малом числе проектов задача решается полным перебором наборов 0/1, что и '
     'реализовано на листе 3.12; для больших множеств используется целочисленное '
     'линейное программирование («Поиск решения» с бинарным ограничением на x).',
     first_indent=1.25, after=6)
para('5. Анализ чувствительности — исследование того, как изменится NPV (и вывод о '
     'проекте) при варьировании цены, объёма выпуска, издержек, ставки налога, '
     'нормы дисконтирования и уровня инфляции. В отчёте он выполнен по уровню '
     'инфляции: листы 3.4 и 3.6 содержат таблицы вариантов и графики NPV(i), а '
     'граница эффективности найдена точно из уравнения NPV(i) = 0.',
     first_indent=1.25, after=10)

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
