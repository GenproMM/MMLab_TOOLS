# -*- coding: utf-8 -*-
"""
Пример: простая проверка — найти стены без значения Mark.
Читает модель, не модифицирует. Не требует настроек.
"""

SCRIPT_ID = "vor_e4a12f07"
SCRIPT_NAME = "\u0421\u0442\u0435\u043D\u044B \u0431\u0435\u0437 \u043C\u0430\u0440\u043A\u0438"
SCRIPT_DESCRIPTION = "\u041F\u0440\u043E\u0432\u0435\u0440\u044F\u0435\u0442 \u0432\u0441\u0435 \u0441\u0442\u0435\u043D\u044B \u0432 \u043C\u043E\u0434\u0435\u043B\u0438 \u0438 \u043D\u0430\u0445\u043E\u0434\u0438\u0442 \u0442\u0435, \u0443 \u043A\u043E\u0442\u043E\u0440\u044B\u0445 \u043D\u0435 \u0437\u0430\u043F\u043E\u043B\u043D\u0435\u043D \u043F\u0430\u0440\u0430\u043C\u0435\u0442\u0440 'Mark' (\u041C\u0430\u0440\u043A\u0430). \u0412\u043E\u0437\u0432\u0440\u0430\u0449\u0430\u0435\u0442 \u0441\u043F\u0438\u0441\u043E\u043A \u043F\u0440\u043E\u0431\u043B\u0435\u043C\u043D\u044B\u0445 \u0441\u0442\u0435\u043D."

from pyrevit import revit, DB
from core.validation_engine import ValidationResult


def run(doc, section, project, settings=None):
    """Найти стены с пустым параметром Mark."""
    try:
        walls = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_Walls)
                 .WhereElementIsNotElementType()
                 .ToElements())

        if not walls:
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=True,
                message="\u0421\u0442\u0435\u043D\u044B \u0432 \u043C\u043E\u0434\u0435\u043B\u0438 \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u044B.",
                skip_summary=True
            )

        problems = []
        for wall in walls:
            mark_param = wall.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_MARK)
            if mark_param:
                value = mark_param.AsString()
                if not value or not value.strip():
                    problems.append(wall)

        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=len(problems) == 0,
            message="\u0412\u0441\u0435\u0433\u043E \u0441\u0442\u0435\u043D: {}. \u0411\u0435\u0437 \u043C\u0430\u0440\u043A\u0438: {}".format(len(walls), len(problems)),
            elements=[w.Id for w in problems],
            skip_summary=True
        )

    except Exception as e:
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message="\u041E\u0448\u0438\u0431\u043A\u0430: {}".format(str(e)),
            skip_summary=True
        )
