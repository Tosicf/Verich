# -*- coding: utf-8 -*-
"""Страница «ОГЛАВЛЕНИЕ» для отчётов УСР: настоящее поле Word + свой упаковщик страниц.

Как пользоваться
----------------
    from toc import Toc
    TOC = Toc(doc)
    ...
    def heading(text):            # в строитель отчёта
        p = para(text, align=CENTER, bold=True, ...)
        TOC.add(p, 1)            # помечает заголовок: outlineLvl + закладка
        return p
    ...
    TOC.finish()                 # вставляет ОГЛАВЛЕНИЕ и проставляет номера
    doc.save(OUT)

Что делает ``finish()``
-----------------------
1. Перед первым заголовком вставляет разрыв страницы, заголовок «ОГЛАВЛЕНИЕ» и
   строки «название … N», собранные в поле ``{ TOC \\o "1-2" \\h \\z \\u }``;
   каждый номер — поле ``{ PAGEREF _TocNNNN \\h }``, а сам заголовок получает
   закладку, на которую это поле ссылается (и гиперссылку для Ctrl-клика).
2. Прогоняет тело документа через ``paginate()`` — грубый измеритель высоты
   текстовых строк, таблиц, картинок и формул — и записывает посчитанные номера
   страниц в кэш полей.
3. Ставит ``w:dirty="true"`` и ``<w:updateFields w:val="true"/>``, поэтому Word
   при открытии сам пересчитает оглавление: цифры будут точными, даже если наша
   оценка ошибётся на страницу. В редакторах, которые поля не пересчитывают
   (Google Docs, просмотрщик), видны наши числа.
"""
import math

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
EMU_PER_PT = 12700.0
EMU_PER_CM = 360000.0
LINE_RULE = 1.15          # одиночный интервал Word = 1,15 кегля

# средняя ширина символа Times New Roman в кеглях
_EM = dict(space=0.250, narrow=0.250, mid=0.300, wide=0.833,
           low=0.500, up=0.665, digit=0.500)
_NARROW = ".,:;'’‘|!"
_MID = "()-/[]{}?\"“”«»≈≤≥·"
_WIDE = "%—–±×÷"


def _el(tag, attrs=None, text=None, space=False):
    e = OxmlElement(tag)
    for k, v in (attrs or {}).items():
        e.set(qn(k) if ':' in k and not k.startswith('xml') else '{http://www.w3.org/XML/1998/namespace}' + k
              if k.startswith('xml') else qn(k), str(v))
    if space:
        e.set(qn('xml:space'), 'preserve')
    if text is not None:
        e.text = text
    return e


def _run_el(size=14, bold=False, italic=False, font='Times New Roman'):
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rpr.append(_el('w:rFonts', {k: font for k in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia')}))
    if bold:
        rpr.append(_el('w:b'))
        rpr.append(_el('w:bCs'))
    if italic:
        rpr.append(_el('w:i'))
    rpr.append(_el('w:color', {'w:val': '000000'}))
    rpr.append(_el('w:sz', {'w:val': int(size * 2)}))
    rpr.append(_el('w:szCs', {'w:val': int(size * 2)}))
    r.append(rpr)
    return r


def _text_el(s, size=14, bold=False, italic=False):
    r = _run_el(size, bold, italic)
    r.append(_el('w:t', text=s, space=True))
    return r


# ============================================================ измерение текста ==
def text_pt(s, size):
    """Приблизительная ширина текста в пунктах (метрика Times New Roman)."""
    w = 0.0
    for ch in s:
        if ch == ' ':
            w += _EM['space']
        elif ch in _NARROW:
            w += _EM['narrow']
        elif ch in _MID:
            w += _EM['mid']
        elif ch in _WIDE:
            w += _EM['wide']
        elif ch.isdigit():
            w += _EM['digit']
        elif ('A' <= ch <= 'Z') or ('А' <= ch <= 'Я') or ch == 'Ё':
            w += _EM['up']
        else:
            w += _EM['low']
    return w * size


def p_text(p):
    """Видимый текст элемента (текст формул и полей не считаем буквами)."""
    out = []
    for t in p.iter(W + 't'):
        anc = t.getparent()
        if anc is not None and M in anc.tag:
            continue
        out.append(t.text or '')
    return ''.join(out)


def p_size(p, default=14.0):
    sz = p.find('.//' + W + 'sz')
    if sz is not None and sz.get(qn('w:val')):
        return int(sz.get(qn('w:val'))) / 2.0
    return default


def p_spacing(p):
    sp = p.find(W + 'pPr/' + W + 'spacing')
    before = after = 0.0
    line = 1.0
    if sp is not None:
        before = int(sp.get(qn('w:before')) or 0) / 20.0
        after = int(sp.get(qn('w:after')) or 0) / 20.0
        rule = sp.get(qn('w:lineRule'))
        ln = sp.get(qn('w:line'))
        if ln and rule in (None, 'auto'):
            line = int(ln) / 240.0
    return before, after, line


# ================================================================ упаковщик ====
def measure_par(p, usable_w):
    """Высота параграфа с учётом отступов, в пунктах."""
    size = p_size(p)
    before, after, line = p_spacing(p)
    h = before + after
    line_h = size * line * LINE_RULE
    if p.find('.//' + WP + 'extent') is not None:          # рисунок
        ext = p.find('.//' + WP + 'extent')
        return h + int(ext.get('cy')) / EMU_PER_PT + line_h
    if p.find('.//' + M + 'oMath') is not None:            # формула
        m = p.find('.//' + M + 'm')
        rows = len(m.findall(M + 'mr')) if m is not None else 0
        big = (p.find('.//' + M + 'f') is not None or
               p.find('.//' + M + 'nary') is not None)
        disp = p.find('.//' + M + 'oMathPara') is not None
        if rows:
            h += (rows * 1.18 + 0.45) * line_h
        elif disp and big:
            h += 1.85 * line_h
        elif disp:
            h += 1.4 * line_h
        else:
            h += 1.15 * line_h
        txt = p_text(p).strip()
        if len(txt) > 55:
            h += math.ceil(text_pt(txt, size) / usable_w) * line_h
        return h
    txt = p_text(p)
    ind = 0.0
    el = p.find(W + 'pPr/' + W + 'ind')
    if el is not None:
        for a in ('w:firstLine', 'w:left'):
            v = el.get(qn(a))
            if v:
                ind += abs(int(v)) / 20.0
    width = max(30.0, usable_w - ind)
    lines = max(1, math.ceil(text_pt(txt, size) / width))
    return h + lines * line_h


def measure_row(tr, grid, usable_w):
    cells = tr.findall(W + 'tc')
    rh = 0.0
    for i, tc in enumerate(cells):
        tcw = tc.find(W + 'tcPr/' + W + 'tcW')
        if tcw is not None and tcw.get(qn('w:w')) and tcw.get(qn('w:type')) == 'dxa':
            cw = int(tcw.get(qn('w:w'))) / 20.0
        elif i < len(grid):
            cw = grid[i]
        else:
            cw = usable_w / max(1, len(cells))
        cw = max(14.0, cw - 11.0)
        size = p_size(tc, default=12.0)
        txt = ' '.join(p_text(p) for p in tc.findall(W + 'p')).strip()
        lines = max(1, math.ceil(text_pt(txt, size) / cw)) if txt else 1
        rh = max(rh, lines * size * LINE_RULE + 4.5)
    return rh or (12.0 * LINE_RULE + 4.5)


def paginate(body, sec):
    """{id(элемента): номер страницы} для всех блоков тела документа."""
    usable_w = (sec.page_width - sec.left_margin - sec.right_margin) / EMU_PER_PT
    usable_h = (sec.page_height - sec.top_margin - sec.bottom_margin) / EMU_PER_PT
    pages = {}
    page, y = 1, 0.0

    def place(el, h):
        nonlocal page, y
        pages[id(el)] = page
        y += h

    def flush():
        nonlocal page, y
        if y > 1.0:
            page += 1
        y = 0.0

    grid = []
    for el in list(body):
        if el.tag == W + 'p':
            hard = any(b.get(qn('w:type')) == 'page' for b in el.findall('.//' + W + 'br'))
            pPr = el.find(W + 'pPr')
            if pPr is not None and pPr.find(W + 'pageBreakBefore') is not None:
                flush()
                place(el, 0.0)
                continue
            h = measure_par(el, usable_w)
            if y + h > usable_h + 0.5:
                flush()
            place(el, h)
            if hard:
                flush()
            elif pPr is not None and pPr.find(W + 'keepNext') is not None and \
                    usable_h - y < 2.2 * p_size(el) * LINE_RULE:
                flush()
        elif el.tag == W + 'tbl':
            grid = [int(g.get(qn('w:w'))) / 20.0 for g in el.findall('.//' + W + 'gridCol')]
            for tr in el.findall(W + 'tr'):
                rh = measure_row(tr, grid, usable_w)
                if y + rh > usable_h + 0.5:
                    flush()
                pages.setdefault(id(el), page)
                y += rh
        elif el.tag == W + 'sectPr':
            continue
    return pages, page


# =================================================================== класс Toc ==
class Toc:
    def __init__(self, doc, title='ОГЛАВЛЕНИЕ', size=14, levels=2):
        self.doc = doc
        self.title = title
        self.size = size
        self.levels = levels
        self.items = []           # dict(level, text, el, bookmark, num_run)
        self._bm = 0

    # -- регистрация заголовка -------------------------------------------------
    def add(self, p, level=1, text=None, bookmark=True):
        el = p._p if hasattr(p, '_p') else p
        self._mark(el, level)
        bm = None
        if bookmark:
            self._bm += 1
            bm = '_Toc%08d' % (abs(id(self)) % 10 ** 7 + self._bm)
            self._bookmark(el, self._bm, bm)
        self.items.append(dict(level=level, text=(text or p_text(el)).strip(),
                               el=el, bm=bm, num=None))
        return p

    def _mark(self, p_el, level):
        pPr = p_el.find(W + 'pPr')
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_el.insert(0, pPr)
        if pPr.find(W + 'outlineLvl') is None:
            pPr.append(_el('w:outlineLvl', {'w:val': level - 1}))

    def _bookmark(self, p_el, bid, name):
        start = _el('w:bookmarkStart', {'w:id': bid, 'w:name': name})
        end = _el('w:bookmarkEnd', {'w:id': bid})
        pPr = p_el.find(W + 'pPr')
        (pPr.addnext if pPr is not None else lambda e: p_el.insert(0, e))(start)
        p_el.append(end)

    # -- вставка страницы оглавления -------------------------------------------
    def finish(self, before=None, leading_break=False, trailing_break=True):
        """Вставить страницу «ОГЛАВЛЕНИЕ».

        before         — элемент, перед которым вставлять (по умолчанию — первый
                         заголовок верхнего уровня, т.е. сразу после титула);
        leading_break  — нужен ли разрыв страницы перед ОГЛАВЛЕНИЕМ (в наших
                         отчётах титул уже заканчивается разрывом);
        trailing_break — разрыв после ОГЛАВЛЕНИЯ, чтобы содержимое началось с
                         чистой страницы (как в образце: ОГЛАВЛЕНИЕ — стр. 2).
        """
        body = self.doc.element.body
        if before is None:
            before = next((it['el'] for it in self.items if it['level'] == 1), None)
        if before is None:
            return self
        anchor = before
        blocks = ([self._page_break()] if leading_break else []) + [self._title()]
        entries = []
        for it in self.items:
            p, num = self._entry(it)
            entries.append((p, num, it))
            blocks.append(p)
        if entries:
            self._wrap_field(entries[0][0], entries[-1][0])
        if trailing_break:
            blocks.append(self._page_break())
        for b in blocks:
            anchor.addprevious(b)
        pages, total = paginate(body, self.doc.sections[0])
        for p, num, it in entries:
            n = pages.get(id(it['el']), 1)
            if num is not None:
                num.text = str(n)
        self._update_fields()
        self.page_count = total
        return self

    def _page_break(self):
        p = OxmlElement('w:p')
        r = _run_el(self.size)
        r.append(_el('w:br', {'w:type': 'page'}))
        p.append(r)
        return p

    def _title(self):
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        pPr.append(_el('w:jc', {'w:val': 'center'}))
        pPr.append(_el('w:spacing', {'w:before': int(16 * 20), 'w:after': int(12 * 20),
                                     'w:line': 240, 'w:lineRule': 'auto'}))
        pPr.append(_el('w:keepNext'))
        p.append(pPr)
        p.append(_text_el(self.title.upper(), size=self.size, bold=True))
        return p

    def _entry(self, it):
        lvl, text, bm = it['level'], it['text'], it['bm']
        size = self.size - (0 if lvl == 1 else 0)
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        tabs = OxmlElement('w:tabs')
        tabs.append(_el('w:tab', {'w:val': 'right', 'w:leader': 'dot', 'w:pos': 9350}))
        pPr.append(tabs)
        pPr.append(_el('w:spacing', {'w:before': 0, 'w:after': int(3 * 20),
                                     'w:line': 240, 'w:lineRule': 'auto'}))
        if lvl > 1:
            pPr.append(_el('w:ind', {'w:left': int((0.75 + 0.75 * (lvl - 2)) * 567),
                                     'w:firstLine': 0}))
        p.append(pPr)
        inner = OxmlElement('w:hyperlink') if bm else p
        if bm:
            inner.set(qn('w:anchor'), bm)
            inner.set(qn('w:history'), '1')
        inner.append(_text_el(text, size=size, bold=(lvl == 1)))
        rt = _run_el(size)
        rt.append(OxmlElement('w:tab'))
        inner.append(rt)
        num = None
        if bm:
            r = _run_el(size)
            r.append(_el('w:fldChar', {'w:fldCharType': 'begin'}))
            inner.append(r)
            r = _run_el(size)
            r.append(_el('w:instrText', text=' PAGEREF %s \\h ' % bm, space=True))
            inner.append(r)
            r = _run_el(size)
            r.append(_el('w:fldChar', {'w:fldCharType': 'separate'}))
            inner.append(r)
            num_run = _text_el('1', size=size)
            num = num_run.find(W + 't')
            inner.append(num_run)
            r = _run_el(size)
            r.append(_el('w:fldChar', {'w:fldCharType': 'end'}))
            inner.append(r)
        if bm:
            p.append(inner)
        return p, num

    def _wrap_field(self, first_p, last_p):
        """Поле TOC охватывает все строки: begin/separate — в первой, end — в последней."""
        target = first_p.find(W + 'hyperlink')
        idx = list(first_p).index(target) if target is not None else len(list(first_p)) - 0
        runs = []
        r = _run_el(self.size)
        r.append(_el('w:fldChar', {'w:fldCharType': 'begin', 'w:dirty': 'true'}))
        runs.append(r)
        r = _run_el(self.size)
        r.append(_el('w:instrText', text=' TOC \\o "1-%d" \\h \\z \\u ' % self.levels,
                     space=True))
        runs.append(r)
        r = _run_el(self.size)
        r.append(_el('w:fldChar', {'w:fldCharType': 'separate'}))
        runs.append(r)
        for r in reversed(runs):
            first_p.insert(idx, r)
        r = _run_el(self.size)
        r.append(_el('w:fldChar', {'w:fldCharType': 'end'}))
        last_p.append(r)

    def _update_fields(self):
        st = self.doc.settings.element
        if st.find(W + 'updateFields') is None:
            st.insert(0, _el('w:updateFields', {'w:val': 'true'}))

    # -- справки ---------------------------------------------------------------
    def entries(self):
        return [(it['level'], it['text']) for it in self.items]
