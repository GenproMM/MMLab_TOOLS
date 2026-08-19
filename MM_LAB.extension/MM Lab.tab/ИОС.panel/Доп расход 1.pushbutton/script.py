#! python3
# -*- coding: utf-8 -*-
"""Доп. расход = 1

Устанавливает значение 1 в параметр «Доп. расход» у воздуховодов, фитингов,
арматуры, воздухораспределителей и гибких воздуховодов.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Доп.\nрасход 1"
__author__ = "GENPRO LAB"

# Канонический lib-бутстрап (D-15).
import os
import sys

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM_LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

import revit_compat
from ios_common_helpers import set_additional_flow_value


COMMAND_NAME = u"Доп. расход = 1"
TARGET_VALUE = 1.0


def main(doc):
    """Точка входа: устанавливает параметр «Доп. расход» в 1 у элементов воздуховодов."""
    revit_compat.require_supported_version(COMMAND_NAME)

    transaction = Transaction(doc, COMMAND_NAME)
    transaction.Start()
    try:
        updated_count = set_additional_flow_value(doc, TARGET_VALUE)
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    TaskDialog.Show(COMMAND_NAME, u"Изменено: {0} элементов".format(updated_count))


def _entry():
    """Готовит doc/uidoc и вызывает main (doc/uidoc — параметрами, правило 18)."""
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(COMMAND_NAME, u"Открой проект Revit и повтори команду.")
        return
    main(uidoc.Document)


try:
    _entry()
except SystemExit:
    pass  # require_supported_version уже показал свой диалог
except Exception as ex:
    TaskDialog.Show(COMMAND_NAME, u"Ошибка:\n{0}".format(ex))
