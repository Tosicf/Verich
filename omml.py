# -*- coding: utf-8 -*-
"""Помощник красивых формул (OMML) для python-docx.

Позволяет писать в .docx настоящие формулы Word (не картинки и не текст):
дроби, индексы, степени, суммы с пределами, «система из n строк» с фигурной
скобкой, стрелки →, верхние/нижние индексы у скобок.

Использование:
    f = Math(doc)
    f.para(f.eq(f.frac(f.sub('NPV', 'i'), f.sup('f', '*'))), label='(3.4)',
           align='center')

Каждая формула — обычный абзац Word с <m:oMath>, дальше редактор freely правит.
"""
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

M = 'm:'
W = 'w:'

ALIGN = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'left': WD_ALIGN_PARAGRAPH.LEFT,
         'just': WD_ALIGN_PARAGRAPH.JUSTIFY, 'right': WD_ALIGN_PARAGRAPH.RIGHT}


def _el(tag, **attrs):
    el = OxmlElement(M + tag)
    for k, v in attrs.items():
        el.set(qn(M + k), v)
    return el


class Math:
    """Конструктор OMML-фрагментов, вставляемых в документ doc."""

    def __init__(self, doc, size=14, font='Cambria Math'):
        self.doc = doc
        self.size = size          # pt
        self.font = font

    # ---------------------------------------------------------------- runs ----
    def r(self, text, sty=None):
        """Математический запуск: <m:r> (<m:rPr>) <w:rPr> <m:t> — как пишет Word."""
        om = _el('r')
        if sty:
            mrpr = _el('rPr')
            st = _el('sty'); st.set(qn(M + 'val'), sty)
            mrpr.append(st)
            om.append(mrpr)
        rpr_w = OxmlElement(W + 'rPr')
        rfonts = OxmlElement(W + 'rFonts')
        for a in ('ascii', 'hAnsi', 'cs'):
            rfonts.set(qn(W + a), self.font)
        rpr_w.append(rfonts)
        sz = OxmlElement(W + 'sz'); sz.set(qn('w:val'), str(int(self.size * 2)))
        szcs = OxmlElement(W + 'szCs'); szcs.set(qn('w:val'), str(int(self.size * 2)))
        rpr_w.append(sz); rpr_w.append(szcs)
        om.append(rpr_w)
        t = OxmlElement(M + 't')
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        om.append(t)
        return om

    def t(self, text):
        """Обычный текст формулы (переменные — курсивом делает сам Word)."""
        return self.r(text)

    def op(self, text):
        """Прямой (не курсив) текст: max, min, σ, русские подписи."""
        return self.r(text, sty='p')

    # -------------------------------------------------------------- atoms ----
    def sub(self, base, sb):
        e = _el('sSub')
        g = _el('e'); self._extend(g, base); e.append(g)
        g = _el('sub'); self._extend(g, sb); e.append(g)
        return e

    def sup(self, base, sp):
        e = _el('sSup')
        g = _el('e'); self._extend(g, base); e.append(g)
        g = _el('sup'); self._extend(g, sp); e.append(g)
        return e

    def subsup(self, base, sb, sp):
        e = _el('sSubSup')
        g = _el('e'); self._extend(g, base); e.append(g)
        g = _el('sub'); self._extend(g, sb); e.append(g)
        g = _el('sup'); self._extend(g, sp); e.append(g)
        return e

    def frac(self, num, den):
        e = _el('f')
        g = _el('num'); self._extend(g, num); e.append(g)
        g = _el('den'); self._extend(g, den); e.append(g)
        return e

    def sqrt(self, e_):  # noqa: C901
        e = _el('rad')
        pr = _el('degHide'); pr.set(qn(M + 'val'), '1')
        rpr = _el('radPr'); rpr.append(pr)
        e.append(rpr)
        g = _el('deg'); e.append(g)
        g = _el('e'); self._extend(g, e_); e.append(g)
        return e

    def sum(self, lo, hi, body):
        """Σ с нижним и верхним пределами (пределы сверху/снизу, не справа)."""
        e = _el('nary')
        rpr = _el('naryPr')
        for tag, val in (('chr', '∑'), ('limLoc', 'undOvr'), ('subHide', '0'), ('supHide', '0')):
            c = _el(tag); c.set(qn(M + 'val'), val); rpr.append(c)
        e.append(rpr)
        g = _el('sub'); self._extend(g, lo); e.append(g)
        g = _el('sup'); self._extend(g, hi); e.append(g)
        g = _el('e'); self._extend(g, body); e.append(g)
        return e

    def brk(self, *items, beg='(', end=')'):
        """Скобка вокруг перечисленных элементов."""
        d = _el('d')
        pr = _el('dPr')
        b = _el('begChr'); b.set(qn(M + 'val'), beg); pr.append(b)
        e2 = _el('endChr'); e2.set(qn(M + 'val'), end); pr.append(e2)
        d.append(pr)
        e = _el('e')
        for it in items:
            self._extend(e, it)
        d.append(e)
        return d

    def cases(self, *rows):
        """Система (m:m с m:mr-строками) в большой левой фигурной скобке."""
        rows = [r if isinstance(r, (list, tuple)) else [r] for r in rows]
        m = _el('m')
        mp = _el('mPr')
        mc = _el('mcs')
        col = _el('mc')
        col.set(qn(M + 'algn'), 'left')
        mc.append(col)
        mp.append(mc)
        m.append(mp)
        for row in rows:
            mr = _el('mr')
            e = _el('e')
            self._extend(e, row)
            mr.append(e)
            m.append(mr)
        return self.brk(m, beg='{', end='')

    def eqarr(self, *rows):
        """Несколько строк, выровненных по центру (можно со списком элементов)."""
        rows = [r if isinstance(r, (list, tuple)) else [r] for r in rows]
        e = _el('eqArr')
        pr = _el('eqArrPr'); al = _el('baseJc'); al.set(qn(M + 'val'), 'center')
        pr.append(al); e.append(pr)
        for row in rows:
            g = _el('e')
            self._extend(g, row)
            e.append(g)
        return e

    def _wrap(self, item):
        """str → run, список/кортеж → последовательность, иначе OMML-элемент."""
        raise TypeError('use _extend for sequences')

    def _extend(self, parent, item):
        if isinstance(item, (int, float)):
            parent.append(self.r(str(item)))
        elif isinstance(item, str):
            parent.append(self.r(item))
        elif isinstance(item, (list, tuple)):
            for it in item:
                self._extend(parent, it)
        else:
            parent.append(item)
        return parent

    # ------------------------------------------------------------- output ----
    def omath(self, *items):
        om = OxmlElement(M + 'oMath')
        for it in items:
            self._extend(om, it)
        return om

    def math_para(self, items, jc='center'):
        """<m:oMathPara> — «блочная» формула Word: дроби в полный рост,
        пределы Σ сверху/снизу, выравнивание по центру или по левому краю."""
        omp = OxmlElement(M + 'oMathPara')
        pr = OxmlElement(M + 'oMathParaPr')
        j = OxmlElement(M + 'jc'); j.set(qn(M + 'val'), jc)
        pr.append(j)
        omp.append(pr)
        omp.append(self.omath(*items))
        return omp

    def para(self, *items, label=None, align='center', before=6, after=6, line=None,
             break_before=False):
        """Абзац с формулой; label — метка вида «(3.4)» справа."""
        p = self.doc.add_paragraph()
        p.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.CENTER)
        pf = p.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        if line:
            pf.line_spacing = line
        if break_before:
            p.add_run().add_break(WD_BREAK.PAGE)
        p._p.append(self.math_para(items, 'center' if align == 'center' else 'left'))
        if label:
            r = OxmlElement(W + 'r')
            rpr = OxmlElement(W + 'rPr')
            rf = OxmlElement(W + 'rFonts')
            for a in ('ascii', 'hAnsi', 'cs'):
                rf.set(qn(W + a), 'Times New Roman')
            rpr.append(rf)
            sz = OxmlElement(W + 'sz'); sz.set(qn(W + 'val'), str(int(self.size * 2)))
            rpr.append(sz)
            r.append(rpr)
            t = OxmlElement(W + 't'); t.text = '        ' + label
            t.set(qn('xml:space'), 'preserve')
            r.append(t)
            p._p.append(r)
        return p

    # ----------------------------------------------------- текстовый хелпер ---
    def text(self, s, size=None):
        return self.r(s)
