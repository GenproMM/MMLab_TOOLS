#! python3
# -*- coding: utf-8 -*-
"""Конфузор/Диффузор

Анализирует переходы (transition) воздуховодной сети: сравнивает площади
входного и выходного коннекторов и записывает результат в параметр
«Конфузор» — 1, если переход сужается (конфузор), 0, если расширяется
(диффузор).

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Конфузор\nДиффузор"
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
from Autodesk.Revit.DB import FlowDirectionType
from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

import revit_compat
from ios_common_helpers import collect_elements
from ios_common_helpers import get_connector_area
from ios_common_helpers import get_hvac_connectors
from ios_common_helpers import get_parameter_by_names
from ios_common_helpers import is_writable
from ios_common_helpers import nearly_equal
from ios_common_helpers import normalize_text
from ios_common_helpers import to_text


COMMAND_NAME = u"Конфузор/Диффузор"
AREA_TOLERANCE = 1e-6


def main(doc):
    """Точка входа: классифицирует переходы воздуховодов как конфузор/диффузор."""
    revit_compat.require_supported_version(COMMAND_NAME)

    confuser_count = 0
    diffuser_count = 0
    skipped_count = 0

    fittings = collect_elements(doc, BuiltInCategory.OST_DuctFitting)

    transaction = Transaction(doc, COMMAND_NAME)
    transaction.Start()
    try:
        for element in fittings:
            family_instance = element if hasattr(element, "MEPModel") else None
            if family_instance is None or family_instance.MEPModel is None:
                skipped_count += 1
                continue

            try:
                part_type = family_instance.MEPModel.PartType
            except Exception:
                skipped_count += 1
                continue

            if normalize_text(to_text(part_type)) != "transition":
                skipped_count += 1
                continue

            hvac_connectors = get_hvac_connectors(family_instance)
            if len(hvac_connectors) != 2:
                skipped_count += 1
                continue

            in_connector = None
            out_connector = None
            for connector in hvac_connectors:
                if connector.Direction == FlowDirectionType.In:
                    in_connector = connector
                elif connector.Direction == FlowDirectionType.Out:
                    out_connector = connector

            if in_connector is None or out_connector is None:
                skipped_count += 1
                continue

            in_area = get_connector_area(in_connector)
            out_area = get_connector_area(out_connector)
            if in_area <= 0.0 or out_area <= 0.0 or nearly_equal(in_area, out_area, AREA_TOLERANCE):
                skipped_count += 1
                continue

            confuser_parameter = get_parameter_by_names(family_instance, u"Конфузор", u"Confuser")
            if not is_writable(confuser_parameter) or confuser_parameter.StorageType != StorageType.Integer:
                skipped_count += 1
                continue

            is_confuser = in_area > out_area
            new_value = 1 if is_confuser else 0
            if confuser_parameter.AsInteger() != new_value:
                confuser_parameter.Set(new_value)

            if is_confuser:
                confuser_count += 1
            else:
                diffuser_count += 1

        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    total = confuser_count + diffuser_count
    TaskDialog.Show(
        COMMAND_NAME,
        u"Классифицировано переходов: {0}\nКонфузоров: {1}\nДиффузоров: {2}\nПропущено: {3}".format(
            total,
            confuser_count,
            diffuser_count,
            skipped_count,
        ),
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
