#! python3
# -*- coding: utf-8 -*-

import os
import sys

__title__ = u"Сброс потерь"
__author__ = "MM Lab"
__doc__ = u"Сбрасывает параметр «Метод определения потерь» на значение «Не определено»."


SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXTENSION_ROOT, "lib")
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

from ios_common_helpers import collect_elements
from ios_common_helpers import ensure_loss_method_undefined
from ios_common_helpers import get_document
from ios_common_helpers import get_duct_loss_method_not_defined_server_id
from ios_common_helpers import show_error


COMMAND_NAME = u"Сброс потерь"


try:
    doc = get_document(COMMAND_NAME)
    if doc:
        updated_count = 0
        already_count = 0
        skipped_count = 0

        not_defined_server_id = get_duct_loss_method_not_defined_server_id()
        if not not_defined_server_id:
            TaskDialog.Show(
                COMMAND_NAME,
                u"Не удалось определить GUID сервера для значения «Не определено».\n"
                u"Проверь список серверов метода потерь в текущем Revit.",
            )
        else:
            targets = collect_elements(
                doc,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
            )

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
except Exception as ex:
    show_error(COMMAND_NAME, ex)
