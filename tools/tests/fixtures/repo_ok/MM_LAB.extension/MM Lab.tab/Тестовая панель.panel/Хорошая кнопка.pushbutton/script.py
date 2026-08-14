#! python3
# -*- coding: utf-8 -*-
"""Хорошая кнопка

Эталонная тестовая кнопка для чекера конвенции MM LAB.
Ничего не делает с моделью Revit — тело-заглушка без Revit-вызовов.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Хорошая\nкнопка"
__author__ = "GENPRO LAB"

import os
import sys


def main():
    """Заглушка: реальная кнопка работала бы с Revit API через lib."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return 0 if script_dir else 1


if __name__ == "__main__":
    sys.exit(main())
