# -*- coding: utf-8 -*-
"""Графики к отчёту УСР (тема 3): зависимость NPV от уровня инфляции и
сравнение вариантов портфеля. Данные — из /tmp/invest_results.json.
"""
import json
import os

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, 'report_assets')
D = json.load(open('/tmp/invest_results.json'))
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False


def _style(ax):
    ax.grid(True, ls=':', lw=0.7, color='#9a9a9a')
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    ax.tick_params(labelsize=9)


def curve(fname, curve, marks, title, xlabel, xmax=None, root=None, root_label=None,
           ymin=None, hlines=0.0):
    xs = [p[0] for p in curve]
    ys = [p[1] for p in curve]
    fig, ax = plt.subplots(figsize=(7.2, 3.5), dpi=200)
    ax.margins(y=0.12)
    ax.plot([100 * x for x in xs], ys, color='#1f5c99', lw=2.0, zorder=3)
    if hlines is not None:
        ax.axhline(hlines, color='#c0392b', lw=1.0, ls='--', zorder=2)
    for i, (x, y, lab) in enumerate(marks):
        ax.plot(100 * x, y, 'o', ms=5.5, color='#e67e22' if i == 0 else '#2e7d32', zorder=4)
        if i == 0:
            ax.annotate(lab, (100 * x, y), textcoords='offset points', xytext=(9, -15),
                        fontsize=8.5, color='#333333', ha='left')
        else:
            ax.annotate(lab, (100 * x, y), textcoords='offset points', xytext=(-9, -16),
                        fontsize=8.5, color='#333333', ha='right')
    if root is not None:
        ax.plot(100 * root, 0, marker='*', ms=13, color='#c0392b', zorder=5)
        ax.axvline(100 * root, color='#c0392b', lw=0.9, ls='-.', zorder=1)
        ax.annotate(root_label, (100 * root, 0), textcoords='offset points',
                    xytext=(-8, -30), fontsize=9, color='#c0392b', fontweight='bold',
                    ha='right')
    ax.set_title(title, fontsize=10.5)
    ax.set_xlabel(xlabel, fontsize=9.5)
    ax.set_ylabel('NPV, ден. ед.', fontsize=9.5)
    if xmax:
        ax.set_xlim(0, xmax)
    if ymin:
        ax.set_ylim(bottom=ymin)
    _style(ax)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, fname), dpi=200, facecolor='white')
    plt.close(fig)
    print('saved', fname)


def main():
    os.makedirs(OUT, exist_ok=True)
    d34, d36 = D['t34'], D['t36']
    marks34 = [(r['i'], r['npv'], 'i = %g %% → NPV = %.2f' % (100 * r['i'], r['npv']))
               for r in d34['rows'][:5]]
    curve('inv_34_curve.png', d34['curve'], marks34,
          'Зависимость NPV проекта от уровня инфляции (задание 3.4)',
          'уровень инфляции i, %', xmax=62, root=d34['i_max'],
          root_label='i = %.4f%% → NPV = 0' % (100 * d34['i_max']))
    marks36 = [(r['i'], r['npv'], 'i = %g %% → NPV = %.2f' % (100 * r['i'], r['npv']))
               for r in d36['rows'][:3]]
    curve('inv_36_curve.png', d36['curve'], marks36,
          'NPV проекта при различных уровнях инфляции (задание 3.6)',
          'уровень инфляции i, %', xmax=26, root=d36['i_crit'],
          root_label='i = %.3f%% → NPV = 0' % (100 * d36['i_crit']), ymin=-60)

    # ------ 3.12: NPV проектов и сравнение допустимых портфелей ------
    prj = D['t312']['projects']
    chosen = set(D['t312']['best']['names'])
    fig, ax = plt.subplots(figsize=(7.2, 3.0), dpi=200)
    xs = range(len(prj))
    vals = [p['npv'] for p in prj]
    cols = ['#1f5c99' if p['name'] in chosen else '#c9d7e6' for p in prj]
    bars = ax.bar(xs, vals, color=cols, edgecolor='#37506b', lw=0.6)
    for b, p in zip(bars, prj):
        ax.annotate('%.2f' % p['npv'], (b.get_x() + b.get_width() / 2, p['npv']),
                    xytext=(0, 3), textcoords='offset points', ha='center', fontsize=8.5)
    ax.axhline(0, color='#666666', lw=0.8)
    ax.set_xticks(list(xs))
    ax.set_xticklabels(['проект %s' % p['name'] for p in prj], fontsize=9)
    ax.set_ylabel('NPV, млн руб.', fontsize=9.5)
    ax.set_title('NPV проектов при ставке дисконтирования 15 % '
                 '(отобраны в портфель — тёмным)', fontsize=10.5)
    ax.set_ylim(0, max(vals) * 1.2)
    _style(ax)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, 'inv_312_npv.png'), dpi=200, facecolor='white')
    plt.close(fig)
    print('saved inv_312_npv.png')

    feas = d312 = D['t312']['all_feas'][:10]
    labels = ['{' + '; '.join(z['names']) + '}' for z in feas]
    vals = [z['npv'] for z in feas]
    fig, ax = plt.subplots(figsize=(7.2, 3.2), dpi=200)
    colors = ['#1f5c99'] + ['#9dc3e6'] * (len(vals) - 1)
    bars = ax.bar(range(len(vals)), vals, color=colors, edgecolor='#37506b', lw=0.6)
    for b, v in zip(bars, vals):
        ax.annotate('%.2f' % v, (b.get_x() + b.get_width() / 2, v), xytext=(0, 3),
                    textcoords='offset points', ha='center', fontsize=8)
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, fontsize=8, rotation=20, ha='right')
    ax.set_ylabel('Σ NPV, млн руб.', fontsize=9.5)
    ax.set_title('Допустимые варианты портфеля (убывание Σ NPV), оптимальный — выделен',
                 fontsize=10.5)
    ax.set_ylim(0, max(vals) * 1.18)
    _style(ax)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, 'inv_312_portfolios.png'), dpi=200, facecolor='white')
    plt.close(fig)
    print('saved inv_312_portfolios.png')


if __name__ == '__main__':
    main()
