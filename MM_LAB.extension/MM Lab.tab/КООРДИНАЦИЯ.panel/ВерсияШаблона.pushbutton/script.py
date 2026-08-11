# -*- coding: utf-8 -*-
"""Формирование версии шаблона из даты изменения листов.

Берет максимальную дату из системного параметра "Дата изменения"
(в блоке "Даты утверждения/изменения листов") на листах проекта,
преобразует ее в формат vYYMMDD и записывает в общий параметр
GP_01_ВерсияШаблона.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Версия\nШаблона"
__author__ = "MM LAB"
__doc__ = "Пишет GP_01_ВерсияШаблона по последней дате изменения листов"

# === IMPORTS ===
import sys
import os
import re
import clr
from datetime import datetime
from System import Guid

# === VENDOR PATH ===
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_ROOT = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "..", "..", ".."))
VENDOR_DIR = os.path.join(EXTENSION_ROOT, "lib")
if os.path.isdir(VENDOR_DIR) and VENDOR_DIR not in sys.path:
    sys.path.insert(0, VENDOR_DIR)

# === REVIT API ===
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    StorageType,
    Transaction,
)
from Autodesk.Revit.UI import TaskDialog

# === VERSION COMPAT ===
try:
    from Autodesk.Revit.DB import UnitTypeId
    _UNIT_MM = UnitTypeId.Millimeters
except ImportError:
    from Autodesk.Revit.DB import DisplayUnitType
    _UNIT_MM = DisplayUnitType.DUT_MILLIMETERS

# === PYREVIT ===
from pyrevit import script

# === CONFIGURATION ===
TARGET_PARAMETER_NAME = u"GP_01_ВерсияШаблона"
TARGET_PARAMETER_GUID = Guid("6e4772dc-dffd-487b-9e95-76c56f4a66aa")
DATE_PARAM_FALLBACK_NAMES = (u"Дата изменения", u"Revision Date")
NO_TARGET_PARAM_MESSAGE = u"Добавьте параметр GP_01_ВерсияШаблона в проект"


# === HELPERS ===
def to_text(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except NameError:
        return str(value)
    except:
        return str(value)


def get_string_from_param(param):
    if param is None:
        return u""

    try:
        value_str = param.AsValueString()
        if value_str:
            return to_text(value_str).strip()
    except:
        pass

    try:
        if param.StorageType == StorageType.String:
            return to_text(param.AsString()).strip()
    except:
        pass

    return u""


def get_sheet_revision_date_text(sheet):
    # Системное имя параметра "Дата изменения" для листа.
    param = None
    try:
        param = sheet.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION_DATE)
    except:
        param = None

    value = get_string_from_param(param)
    if value:
        return value

    for param_name in DATE_PARAM_FALLBACK_NAMES:
        try:
            fallback_param = sheet.LookupParameter(param_name)
            value = get_string_from_param(fallback_param)
            if value:
                return value
        except:
            pass

    return u""


def parse_revision_date(date_text):
    text = to_text(date_text).strip()
    if not text:
        return None

    match = re.search(r"(\d{1,2})[./\\-](\d{1,2})[./\\-](\d{2,4})", text)
    if not match:
        return None

    day = int(match.group(1))
    month = int(match.group(2))
    year = int(match.group(3))
    if year < 100:
        year += 2000

    try:
        return datetime(year, month, day)
    except:
        return None


def get_latest_project_revision_date(doc):
    sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType()
    latest_dt = None

    for sheet in sheets:
        date_text = get_sheet_revision_date_text(sheet)
        if not date_text:
            continue

        parsed_dt = parse_revision_date(date_text)
        if parsed_dt is None:
            continue

        if latest_dt is None or parsed_dt > latest_dt:
            latest_dt = parsed_dt

    return latest_dt


def format_template_version(date_dt):
    return u"v{0:02d}{1:02d}{2:02d}".format(date_dt.year % 100, date_dt.month, date_dt.day)


def collect_target_parameters(doc):
    targets = []

    def is_target_parameter(parameter):
        if parameter is None:
            return False

        try:
            if parameter.Definition is None:
                return False
            if to_text(parameter.Definition.Name) != TARGET_PARAMETER_NAME:
                return False
        except:
            return False

        try:
            return parameter.GUID == TARGET_PARAMETER_GUID
        except:
            # Если API не возвращает GUID, считаем совпадение по имени допустимым.
            return True

    try:
        proj_info = doc.ProjectInformation
        if proj_info is not None:
            p = proj_info.LookupParameter(TARGET_PARAMETER_NAME)
            if is_target_parameter(p):
                targets.append((proj_info, p))
    except:
        pass

    # Резервный сценарий: параметр может быть привязан к листам.
    if not targets:
        sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType()
        for sheet in sheets:
            try:
                p = sheet.LookupParameter(TARGET_PARAMETER_NAME)
                if is_target_parameter(p):
                    targets.append((sheet, p))
            except:
                pass

    return targets


def set_template_version(doc, version_value):
    target_params = collect_target_parameters(doc)
    if not target_params:
        TaskDialog.Show(__title__, NO_TARGET_PARAM_MESSAGE)
        return False

    updated = 0
    transaction = Transaction(doc, u"Запись GP_01_ВерсияШаблона")

    try:
        transaction.Start()
        for _, parameter in target_params:
            if parameter is None or parameter.IsReadOnly:
                continue
            if parameter.Set(version_value):
                updated += 1
        transaction.Commit()
    except:
        try:
            transaction.RollBack()
        except:
            pass
        raise

    if updated == 0:
        TaskDialog.Show(__title__, u"Параметр найден, но недоступен для записи.")
        return False

    return True


# === MAIN ===
def main():
    output = script.get_output()

    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(__title__, u"Откройте проект Revit и повторите команду.")
        return

    doc = uidoc.Document
    latest_date = get_latest_project_revision_date(doc)
    if latest_date is None:
        TaskDialog.Show(__title__, u"Не найдены корректные даты изменения на листах проекта.")
        return

    new_version = format_template_version(latest_date)
    success = set_template_version(doc, new_version)
    if not success:
        return

    output.print_md(u"**Обновление завершено**")
    output.print_md(u"Последняя дата изменения: **{0}**".format(latest_date.strftime("%d.%m.%y")))
    output.print_md(u"Новое значение `{0}`: **{1}**".format(TARGET_PARAMETER_NAME, new_version))
    TaskDialog.Show(__title__, u"Параметр GP_01_ВерсияШаблона обновлен: {0}".format(new_version))


if __name__ == "__main__":
    main()
