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

    fittings = collect_elements(doc, BuiltInCategory.OST_DuctFitting)

    no_mepmodel_count = 0
    part_type_error_count = 0
    part_type_counts = {}
    probe_lines = []
    connector_probe = []

    for element in fittings:
        family_instance = element if hasattr(element, "MEPModel") else None
        if family_instance is None or family_instance.MEPModel is None:
            no_mepmodel_count += 1
            continue

        try:
            part_type = family_instance.MEPModel.PartType
        except Exception as ex:
            part_type_error_count += 1
            if not probe_lines:
                probe_lines.append(u"PartType бросает: {0}".format(ex))
            continue

        # Проба аксессоров на первом успешно прочитанном элементе.
        if not probe_lines:
            probe_lines.append(u"type = {0}".format(type(part_type).__name__))
            try:
                probe_lines.append(u"str() = [{0}]".format(str(part_type)))
            except Exception as ex:
                probe_lines.append(u"str() ОШИБКА: {0}".format(ex))
            try:
                probe_lines.append(u"ToString() = [{0}]".format(part_type.ToString()))
            except Exception as ex:
                probe_lines.append(u"ToString() ОШИБКА: {0}".format(ex))
            try:
                probe_lines.append(u"to_text() = [{0}]".format(to_text(part_type)))
            except Exception as ex:
                probe_lines.append(u"to_text() ОШИБКА: {0}".format(ex))
            try:
                probe_lines.append(u"normalize = [{0}]".format(normalize_text(to_text(part_type))))
            except Exception as ex:
                probe_lines.append(u"normalize ОШИБКА: {0}".format(ex))

        try:
            key = part_type.ToString()
        except Exception:
            key = to_text(part_type)
        part_type_counts[key] = part_type_counts.get(key, 0) + 1

        # Для первого элемента, похожего на переход, снимаем картину коннекторов.
        if not connector_probe and "transition" in normalize_text(key):
            raw = []
            manager = family_instance.MEPModel.ConnectorManager
            for connector in manager.Connectors:
                raw.append(u"  domain={0} dir={1}".format(connector.Domain, connector.Direction))
            connector_probe.append(u"Сырых коннекторов: {0}".format(len(raw)))
            connector_probe.extend(raw)
            connector_probe.append(
                u"get_hvac_connectors вернул: {0}".format(len(get_hvac_connectors(family_instance)))
            )

    lines = [u"ДИАГНОСТИКА (модель не изменялась)", u""]
    lines.append(u"Всего фитингов: {0}".format(len(fittings)))
    lines.append(u"Нет MEPModel: {0}".format(no_mepmodel_count))
    lines.append(u"Ошибка PartType: {0}".format(part_type_error_count))
    lines.append(u"")
    lines.append(u"Аксессоры на 1-м элементе:")
    lines.extend(probe_lines)
    lines.append(u"")
    lines.append(u"Различные PartType (топ-12):")
    ranked = sorted(part_type_counts.items(), key=lambda kv: kv[1], reverse=True)
    for key, count in ranked[:12]:
        lines.append(u"  [{0}] = {1}".format(key, count))

    if connector_probe:
        lines.append(u"")
        lines.append(u"Коннекторы 1-го перехода:")
        lines.extend(connector_probe)

    TaskDialog.Show(COMMAND_NAME, u"\n".join(lines))


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
