#! python3
# -*- coding: utf-8 -*-
"""Название кнопки  # TODO: замени

Что делает кнопка, кратко.  # TODO: опиши

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Новая\nкнопка"  # TODO: замени (\n переносит подпись кнопки на новую строку)
__author__ = "GENPRO LAB"  # TODO: замени при необходимости

# Канонический lib-бутстрап (D-15) — единственная допустимая форма.
# Путь рассчитан на рабочее положение кнопки ВНУТРИ
# "MM Lab.tab/<Панель>.panel/": пока папка лежит в templates/, lib не
# находится — это норма, шаблон запускается только после копирования.
import os
import sys

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

import revit_compat


COMMAND_NAME = u"Новая кнопка"  # TODO: замени на имя своей кнопки


def main():
    """Точка входа кнопки: чтение модели + транзакционный каркас."""
    # Гейт версии Revit (D-03): на неподдерживаемой версии показывает
    # диалог со списком поддерживаемых версий и завершает скрипт.
    revit_compat.require_supported_version(COMMAND_NAME)

    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(COMMAND_NAME, u"Открой проект Revit и повтори команду.")
        return
    doc = uidoc.Document

    # Пример чтения модели: считаем стены проекта.
    # TODO: замени на сбор нужных тебе элементов
    walls = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    count = revit_compat.iter_count(walls)

    # Любое изменение модели — только внутри транзакции:
    # Commit в try, RollBack + raise в except.
    transaction = Transaction(doc, COMMAND_NAME)
    transaction.Start()
    try:
        # TODO: логика изменения модели (пример читает и НИЧЕГО не меняет)
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    # TODO: замени итоговый отчёт под свою кнопку
    TaskDialog.Show(COMMAND_NAME, u"Стен в проекте: {0}".format(count))


try:
    main()
except SystemExit:
    pass  # require_supported_version уже показал свой диалог
except Exception as ex:
    TaskDialog.Show(COMMAND_NAME, u"Ошибка:\n{0}".format(ex))
