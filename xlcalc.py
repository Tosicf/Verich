# -*- coding: utf-8 -*-
"""Крошечный вычислитель формул листа openpyxl (для скриншотов «как в Excel»).

Поддерживает то, что встречается в наших книгах: ссылки на ячейки (в т.ч. с $),
диапазоны в SUM / СУММПРОИЗВ, арифметику + - * / ^, скобки.
"""
import re

from openpyxl.utils import column_index_from_string, range_boundaries

REF = re.compile(r'\$?([A-Za-z]{1,3})\$?(\d+)')


def _rc(ref):
    col, row = ref[0], int(ref[1])
    return int(row), column_index_from_string(col.upper())


class SheetCalc:
    def __init__(self, ws):
        self.ws = ws
        self.cache = {}

    def raw(self, r, c):
        return self.ws.cell(row=r, column=c).value

    def value(self, r, c):
        key = (r, c)
        if key in self.cache:
            return self.cache[key]
        self.cache[key] = 0
        v = self.raw(r, c)
        if isinstance(v, str) and v.startswith('='):
            v = self.eval(v[1:])
        self.cache[key] = v
        return v

    def eval(self, expr):
        expr = expr.replace('\n', ' ')
        # функции над диапазонами
        def sumprod(m):
            a, b = m.group(1), m.group(2)
            c1, r1, c2, r2 = range_boundaries(a.replace('$', ''))
            d1, s1, d2, s2 = range_boundaries(b.replace('$', ''))
            va = [self.value(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
            vb = [self.value(r, c) for r in range(s1, s2 + 1) for c in range(d1, d2 + 1)]
            return repr(sum((x or 0) * (y or 0) for x, y in zip(va, vb)))

        def summ(m):
            c1, r1, c2, r2 = range_boundaries(m.group(1).replace('$', ''))
            vals = [self.value(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
            return repr(sum(v or 0 for v in vals))

        expr = re.sub(r'(?:СУММПРОИЗВ|SUMPRODUCT)\(\s*([^,;]+)[,;]\s*([^)]+)\)', sumprod, expr)
        expr = re.sub(r'SUM\(\s*([^)]+)\)', summ, expr)
        expr = expr.replace(';', ',')
        # одиночные ссылки
        expr = REF.sub(lambda m: repr(self.value(*_rc(m.groups())) or 0), expr)
        expr = expr.replace('^', '**')
        try:
            return eval(expr, {'__builtins__': {}}, {})
        except Exception:
            return None

    def text(self, r, c, nd=4):
        v = self.value(r, c)
        if v is None:
            return ''
        if isinstance(v, str):
            return v.strip()
        if isinstance(v, float):
            if abs(v - round(v)) < 1e-9 and abs(v) < 1e12:
                return str(int(round(v)))
            s = f'{v:.{nd}f}'.rstrip('0').rstrip('.')
            return s.replace('.', ',')
        return str(v)
