# -*- coding: utf-8 -*-
"""
Пример: скрипт модификации — заполнить пустые Mark у стен.
Демонстрирует паттерн с транзакцией для изменения модели.
"""

SCRIPT_ID = "vor_5f7a2e93"
SCRIPT_NAME = u"\u0410\u0432\u0442\u043e\u043c\u0430\u0440\u043a\u0438\u0440\u043e\u0432\u043a\u0430 \u0441\u0442\u0435\u043d"
SCRIPT_DESCRIPTION = u"\u041d\u0430\u0445\u043e\u0434\u0438\u0442 \u0441\u0442\u0435\u043d\u044b \u0441 \u043f\u0443\u0441\u0442\u044b\u043c \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u043c Mark \u0438 \u0437\u0430\u043f\u043e\u043b\u043d\u044f\u0435\u0442 \u0435\u0433\u043e \u043f\u043e \u0448\u0430\u0431\u043b\u043e\u043d\u0443 'W001', 'W002' \u0438 \u0442.\u0434. \u041c\u043e\u0434\u0438\u0444\u0438\u0446\u0438\u0440\u0443\u0435\u0442 \u043c\u043e\u0434\u0435\u043b\u044c."

from pyrevit import revit, DB
from core.validation_engine import ValidationResult


def run(doc, section, project, settings=None):
    try:
        walls = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_Walls)
                 .WhereElementIsNotElementType()
                 .ToElements())

        if not walls:
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=True,
                message=u"\u0421\u0442\u0435\u043d\u044b \u0432 \u043c\u043e\u0434\u0435\u043b\u0438 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b.",
                skip_summary=True
            )

        # Находим стены без марки и определяем максимальный существующий номер
        empty_walls = []
        existing_nums = []
        for wall in walls:
            mark_param = wall.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_MARK)
            if mark_param:
                val = mark_param.AsString()
                if not val or not val.strip():
                    empty_walls.append(wall)
                else:
                    # Пытаемся извлечь число из существующей марки
                    try:
                        num = int("".join(c for c in val if c.isdigit()))
                        existing_nums.append(num)
                    except ValueError:
                        pass

        if not empty_walls:
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=True,
                message=u"\u0412\u0441\u0435 \u0441\u0442\u0435\u043d\u044b \u0438\u043c\u0435\u044e\u0442 \u043c\u0430\u0440\u043a\u0438. \u041d\u0438\u0447\u0435\u0433\u043e \u043d\u0435 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u043e.",
                skip_summary=True
            )

        # Начинаем нумерацию с максимального существующего + 1
        next_num = max(existing_nums) + 1 if existing_nums else 1

        modified = 0
        errors = []
        t = DB.Transaction(doc, u"\u0410\u0432\u0442\u043e\u043c\u0430\u0440\u043a\u0438\u0440\u043e\u0432\u043a\u0430 \u0441\u0442\u0435\u043d")
        t.Start()

        try:
            for wall in empty_walls:
                mark_param = wall.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_MARK)
                if mark_param:
                    mark_param.Set("W{:03d}".format(next_num))
                    next_num += 1
                    modified += 1

            t.Commit()

            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=True,
                message=u"\u041f\u0440\u0438\u0441\u0432\u043e\u0435\u043d\u043e \u043c\u0430\u0440\u043e\u043a: {}".format(modified),
                skip_summary=True
            )

        except Exception as e:
            t.RollBack()
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=False,
                message=u"\u041e\u0448\u0438\u0431\u043a\u0430 \u043f\u0440\u0438 \u0437\u0430\u043f\u0438\u0441\u0438: {}".format(str(e)),
                skip_summary=True
            )

    except Exception as e:
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message=u"\u041e\u0448\u0438\u0431\u043a\u0430: {}".format(str(e)),
            skip_summary=True
        )
