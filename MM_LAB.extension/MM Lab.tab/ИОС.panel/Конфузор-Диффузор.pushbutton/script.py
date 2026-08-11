#! python3
# -*- coding: utf-8 -*-

import os
import sys

__title__ = u"Конфузор/Диффузор"
__author__ = "MM Lab"
__doc__ = u"Определяет, является ли переход конфузором или диффузором."

SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXTENSION_ROOT, "lib")
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import FlowDirectionType
from Autodesk.Revit.DB import StorageType
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

from ios_common_helpers import collect_elements
from ios_common_helpers import get_connector_area
from ios_common_helpers import get_hvac_connectors
from ios_common_helpers import get_parameter_by_names
from ios_common_helpers import get_document
from ios_common_helpers import is_writable
from ios_common_helpers import nearly_equal
from ios_common_helpers import normalize_text
from ios_common_helpers import show_error
from ios_common_helpers import to_text


COMMAND_NAME = u"Конфузор/Диффузор"
AREA_TOLERANCE = 1e-6


try:
    doc = get_document(COMMAND_NAME)
    if doc:
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
except Exception as ex:
    show_error(COMMAND_NAME, ex)
