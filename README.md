# Verich

Репозиторий учебно-самостоятельных работ (УСР) по дисциплине «Эконометрика и
исследование операций в экономике». Каждый отчёт собирается скриптами из
корня репозитория — правки вносятся в скрипты, а не в .docx.

## Готовые отчёты (`outputs/`)

| тема | отчёт | книга Excel |
|---|---|---|
| 1. Модели межотраслевого баланса (задача 1.6, вариант б) | `Эконометрика УСР 1 (Пась А.М.).docx` | — |
| 2. Модели и методы многоцелевой оптимизации (задание 2.4) | `Эконометрика УСР 2 (Пась А.М.) — задание 2.4.docx` | `…задание 2.4.xlsx` |
| 3. Модели анализа и оценки инвестиционных проектов (задания 3.4, 3.6, 3.12) | `Эконометрика УСР 3 (Пась А.М.) — задания 3.4, 3.6, 3.12.docx` | `…задания 3.4, 3.6, 3.12.xlsx` |

`Эконометрика УСР 5.docx/.xlsx` — пример из методички (чужая работа), не пересобирается.

## Сборка

Нужны `openpyxl`, `python-docx`, `matplotlib`, `numpy`, `scipy`:

```bash
python3 -m venv /tmp/venv && /tmp/venv/bin/pip install openpyxl python-docx matplotlib numpy scipy
```

Тема 1:

```bash
python make_mobr_screenshots.py    # кадры листов Excel-решения → report_assets/
python build_mobr_report.py        # → outputs/…УСР 1….docx
```

Тема 2 (`mmolo_models.py` — модели, `solve_mmolo.py` — ЛП-расчёт в `/tmp/mmolo_results.json`):

```bash
python solve_mmolo.py
python build_mmolo_excel.py        # → outputs/…УСР 2….xlsx  (листы с живыми формулами)
python make_mmolo_screenshots.py   # кадры для РЕШЕНИЯ
python make_mmolo_appendix.py      # полные снимки листов для ПРИЛОЖЕНИЕ А
python build_mmolo_report.py       # → outputs/…УСР 2….docx
```

Тема 3 (`solve_invest.py` — NPV/IRR/PI/PP/DPP/MIRR, критическая инфляция и
перебор 0/1 для портфеля, результат в `/tmp/invest_results.json`):

```bash
python solve_invest.py
python build_invest_excel.py
python make_invest_screenshots.py
python make_invest_charts.py       # графики NPV(i) и сравнения портфелей
python build_invest_report.py
```

Формулы в отчётах — настоящие объекты Word (OMML), поэтому пересборка .docx
обязательна после любой правки.

## Общие модули

- `omml.py` — конструктор формул Word: `frac`, `sub`, `sup`, `subsup`, `sqrt`,
  `sum` (Σ с пределами), `brk`, `cases` (система с фигурной скобкой), `eqarr`;
  вывод в блочном режиме (`m:oMathPara`), `para(…, label='(3.4)')` добавляет номер.
- `toc.py` — страница «ОГЛАВЛЕНИЕ»: `TOC.add(p, level)` помечает заголовок
  (`outlineLvl` + закладка), `TOC.finish()` вставляет оглавление полем Word
  `{ TOC \o "1-2" \h \z \u }` со ссылками `{ PAGEREF }`. Номера страниц в кэше
  считает собственный упаковщик страниц (`paginate`), а Word при открытии
  пересчитывает поля сам (`w:dirty` + `<w:updateFields/>`), поэтому цифры всегда
  актуальны. Чтобы поле обновилось вручную: Ctrl+A → F9 (в Word).
- `xlcalc.py` — мини-калькулятор формул листа openpyxl (SUM, SUMPRODUCT, `^`);
  нужен, чтобы на кадрах Excel были видны вычисленные значения.
- `build_excel.py`, `make_screenshots.py`, `build_report.py` — то же для темы 5
  (пример из методички).

## Оформление (единое для всех отчётов)

Times New Roman 14 pt, интервал 1,0, выравнивание по ширине; поля 3,0 / 1,5 /
2,0 / 2,0 см; структура ТИТУЛ → ОГЛАВЛЕНИЕ → ЗАДАЧА → РЕШЕНИЕ → ПРИЛОЖЕНИЕ А
(кадры листов Excel) → ПРИЛОЖЕНИЕ Б (методика); таблицы тонкие чёрно-белые,
рисунки в естественной плотности (300 dpi), нумерация страниц в колонтитуле,
титул без номера.
