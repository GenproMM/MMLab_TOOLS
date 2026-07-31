# -*- coding: utf-8 -*-
"""
Пример: папочный скрипт с кастомным окном настроек.
Скрипт валидирует параметр Mark у элементов выбранной категории.
Пользователь выбирает категорию и вводит шаблон марки через кастомное окно.
"""

SCRIPT_ID = "vor_c1d4a6b8"
SCRIPT_NAME = u"\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u043c\u0430\u0440\u043a\u0438 \u043f\u043e \u0448\u0430\u0431\u043b\u043e\u043d\u0443"
SCRIPT_DESCRIPTION = u"\u041f\u0440\u043e\u0432\u0435\u0440\u044f\u0435\u0442 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 Mark \u0443 \u044d\u043b\u0435\u043c\u0435\u043d\u0442\u043e\u0432 \u0432\u044b\u0431\u0440\u0430\u043d\u043d\u043e\u0439 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 \u043d\u0430 \u0441\u043e\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0438\u0435 \u0448\u0430\u0431\u043b\u043e\u043d\u0443. \u0418\u0441\u043f\u043e\u043b\u044c\u0437\u0443\u0435\u0442 \u0441\u043e\u0431\u0441\u0442\u0432\u0435\u043d\u043d\u043e\u0435 \u043e\u043a\u043d\u043e \u043d\u0430\u0441\u0442\u0440\u043e\u0435\u043a."
HAS_SETTINGS = True

import os
import sys

from pyrevit import revit, DB
from core.validation_engine import ValidationResult


def show_settings(doc, current_settings):
    """Показать кастомное окно настроек."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.join(script_dir, "lib")
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)

    from settings_window import CheckSettingsWindow
    dlg = CheckSettingsWindow(doc, current_settings)
    return dlg.show_dialog()


def run(doc, section, project, settings=None):
    if not settings:
        settings = {}

    category_name = settings.get("category", "")
    pattern = settings.get("pattern", "")

    if not category_name or not pattern:
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message=u"\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u0442\u0435 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044e \u0438 \u0448\u0430\u0431\u043b\u043e\u043d \u043c\u0430\u0440\u043a\u0438 \u0432 \u043d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0430\u0445 \u0441\u043a\u0440\u0438\u043f\u0442\u0430.",
            skip_summary=True
        )

    # Маппинг категорий
    cat_map = {
        u"\u0421\u0442\u0435\u043d\u044b": DB.BuiltInCategory.OST_Walls,
        u"\u0414\u0432\u0435\u0440\u0438": DB.BuiltInCategory.OST_Doors,
        u"\u041e\u043a\u043d\u0430": DB.BuiltInCategory.OST_Windows,
        u"\u041f\u0435\u0440\u0435\u043a\u0440\u044b\u0442\u0438\u044f": DB.BuiltInCategory.OST_Floors,
        u"\u041a\u043e\u043b\u043e\u043d\u043d\u044b": DB.BuiltInCategory.OST_Columns,
    }

    bic = cat_map.get(category_name)
    if not bic:
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message=u"\u041d\u0435\u0438\u0437\u0432\u0435\u0441\u0442\u043d\u0430\u044f \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f: {}".format(category_name),
            skip_summary=True
        )

    try:
        elements = (DB.FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .ToElements())

        if not elements:
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=True,
                message=u"\u042d\u043b\u0435\u043c\u0435\u043d\u0442\u044b \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 '{}' \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b.".format(category_name),
                skip_summary=True
            )

        problems = []
        for elem in elements:
            mark_param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_MARK)
            if mark_param:
                value = mark_param.AsString()
                if not value or not value.strip():
                    problems.append(elem)
                elif pattern and not value.startswith(pattern):
                    problems.append(elem)

        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=len(problems) == 0,
            message=u"\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f: {}. \u041f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e: {}. \u041f\u0440\u043e\u0431\u043b\u0435\u043c: {}. \u0428\u0430\u0431\u043b\u043e\u043d: '{}'".format(
                category_name, len(elements), len(problems), pattern),
            elements=[p.Id for p in problems],
            skip_summary=True
        )

    except Exception as e:
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message=u"\u041e\u0448\u0438\u0431\u043a\u0430: {}".format(str(e)),
            skip_summary=True
        )
