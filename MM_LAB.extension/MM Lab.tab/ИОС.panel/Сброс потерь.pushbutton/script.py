#! python3
# -*- coding: utf-8 -*-
"""Сброс потерь

Сбрасывает параметр «Метод определения потерь» на значение «Не определено»
у фитингов и арматуры воздуховодов.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = u"Сброс\nпотерь"
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

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

import revit_compat
from ios_common_helpers import collect_elements
from ios_common_helpers import ensure_loss_method_undefined
from ios_common_helpers import get_duct_loss_method_not_defined_server_id


COMMAND_NAME = u"Сброс потерь"


def main(doc):
    """Точка входа: сбрасывает метод определения потерь у фитингов и арматуры воздуховодов."""
    revit_compat.require_supported_version(COMMAND_NAME)

    not_defined_server_id = get_duct_loss_method_not_defined_server_id()
    if not not_defined_server_id:
        TaskDialog.Show(
            COMMAND_NAME,
            u"Не удалось определить GUID сервера для значения «Не определено».\n"
            u"Проверь список серверов метода потерь в текущем Revit.",
        )
        return

    targets = collect_elements(
        doc,
        BuiltInCategory.OST_DuctFitting,
        BuiltInCategory.OST_DuctAccessory,
    )

    updated_count = 0
    already_count = 0
    skipped_count = 0

    transaction = Transaction(doc, COMMAND_NAME)
    transaction.Start()
    try:
        for element in targets:
            family_instance = element if hasattr(element, "MEPModel") else None
            if family_instance is None:
                skipped_count += 1
                continue

            outcome = ensure_loss_method_undefined(family_instance, not_defined_server_id)
            if outcome == "updated":
                updated_count += 1
            elif outcome == "already":
                already_count += 1
            else:
                skipped_count += 1

        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    TaskDialog.Show(
        COMMAND_NAME,
        u"Переведено в «Не определено»: {0} элементов\n"
        u"Уже были в «Не определено»: {1}\n"
        u"Пропущено: {2}".format(updated_count, already_count, skipped_count),
    )


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
