#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Юнит-прогон gap-aware выбора пика профиля (без БД и xlsx).

Запуск: python3 backend/analyzers/test_profile_gaps.py
Проверяет: исключение граничных отсчётов у «дыр» при выборе пика; регресс без
дыр; дыра в начале/в конце; две смежные дыры (общий отсчёт исключается раз);
сплошные дыры → пика нет; построение часовой сетки.
"""
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from profile_analyzer import _peak_excluding_gaps, _hour_grid  # noqa: E402

base = datetime(2026, 7, 1, 0, 0)


def H(n):
    return base + timedelta(hours=n)


def grid(nmax):
    return [H(i) for i in range(nmax + 1)]


def check(name, cond):
    print(("OK   " if cond else "FAIL ") + name)
    assert cond, name


def main():
    # 1) Спайки до/после дыры исключаются, пик — не граничный.
    hourly = {H(0): 0.5, H(1): 9.9, H(6): 8.8, H(7): 2.0}   # дыра h2..h5
    pdt, praw, gaps, exc = _peak_excluding_gaps(hourly, grid(7))
    check("example: peak=h7 (2.0), не спайки 9.9/8.8", pdt == H(7) and praw == 2.0)
    check("example: gaps=1, excluded=2", gaps == 1 and exc == 2)

    # 2) Без дыр — пик как раньше (регресс).
    hourly = {H(0): 1, H(1): 2, H(2): 9, H(3): 3, H(4): 1}
    pdt, praw, gaps, exc = _peak_excluding_gaps(hourly, grid(4))
    check("no-gap: peak=h2 (9), gaps=0", pdt == H(2) and praw == 9 and gaps == 0 and exc == 0)

    # 3) Дыра в начале — исключается только «после».
    hourly = {H(2): 9.0, H(3): 2, H(4): 1}
    pdt, praw, gaps, exc = _peak_excluding_gaps(hourly, grid(4))
    check("gap-start: h2 исключён, peak=h3 (2)", pdt == H(3) and gaps == 1 and exc == 1)

    # 4) Дыра в конце — исключается только «до».
    hourly = {H(0): 1, H(1): 2, H(2): 9}
    pdt, praw, gaps, exc = _peak_excluding_gaps(hourly, grid(4))
    check("gap-end: h2 исключён, peak=h1 (2)", pdt == H(1) and gaps == 1 and exc == 1)

    # 5) Две смежные дыры, общий отсчёт h3 исключается один раз.
    hourly = {H(0): 1, H(1): 5, H(3): 9, H(5): 4, H(6): 2}   # отсутствуют h2, h4
    pdt, praw, gaps, exc = _peak_excluding_gaps(hourly, grid(6))
    check("two-gaps: gaps=2, excluded=3, пик не h3", gaps == 2 and exc == 3 and pdt != H(3))

    # 6) Сплошные дыры — пика нет.
    pdt, praw, gaps, exc = _peak_excluding_gaps({}, grid(4))
    check("all-gap: peak None", pdt is None)

    # 7) Часовая сетка из меток (в т.ч. получасовых).
    g = _hour_grid([H(0), H(0) + timedelta(minutes=30), H(3)])
    check("hour_grid: h0..h3", g == grid(3))

    print("\nВСЕ ТЕСТЫ ПРОЙДЕНЫ")


if __name__ == '__main__':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    main()
