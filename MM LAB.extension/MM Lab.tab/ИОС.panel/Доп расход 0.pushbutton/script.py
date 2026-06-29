#! python3
# -*- coding: utf-8 -*-

import os
import sys

__title__ = u"Доп. расход = 0"
__author__ = "MM Lab"
__doc__ = u"Устанавливает значение 0 в параметр «Доп. расход»."

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXTENSION_ROOT, "lib")
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

from ios_common_helpers import get_document
from ios_common_helpers import set_additional_flow_value
from ios_common_helpers import show_error


COMMAND_NAME = u"Доп. расход = 0"
TARGET_VALUE = 0.0


try:
    doc = get_document(COMMAND_NAME)
    if doc:
        transaction = Transaction(doc, COMMAND_NAME)
        transaction.Start()
        try:
            updated_count = set_additional_flow_value(doc, TARGET_VALUE)
            transaction.Commit()
        except Exception:
            transaction.RollBack()
            raise

        TaskDialog.Show(COMMAND_NAME, u"Изменено: {0} элементов".format(updated_count))
except Exception as ex:
    show_error(COMMAND_NAME, ex)
