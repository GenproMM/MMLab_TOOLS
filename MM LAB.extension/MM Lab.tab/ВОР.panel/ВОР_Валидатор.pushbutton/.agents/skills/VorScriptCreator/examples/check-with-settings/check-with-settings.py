# -*- coding: utf-8 -*-
"""
Пример: скрипт с настройками — проверить параметры на выбранных листах.
Использует SETTINGS_SCHEMA (generic-окно) с разными типами настроек.
"""

SCRIPT_ID = "vor_b83c9d51"
SCRIPT_NAME = u"\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u043b\u0438\u0441\u0442\u043e\u0432"
SCRIPT_DESCRIPTION = u"\u041f\u0440\u043e\u0432\u0435\u0440\u044f\u0435\u0442 \u043d\u0430\u043b\u0438\u0447\u0438\u0435 \u0437\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u044b\u0445 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u0432 'Designed By' \u0438 'Checked By' \u043d\u0430 \u0432\u044b\u0431\u0440\u0430\u043d\u043d\u044b\u0445 \u043b\u0438\u0441\u0442\u0430\u0445."
HAS_SETTINGS = True
SETTINGS_SCHEMA = [
    {"key": "selected_sheets", "type": "sheet_list", "label": u"\u041b\u0438\u0441\u0442\u044b \u0434\u043b\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438",
     "sortable": True, "hide_unselected": True},
    {"key": "check_designed", "type": "checkbox", "label": u"\u041f\u0440\u043e\u0432\u0435\u0440\u044f\u0442\u044c Designed By",
     "checkbox_label": u"\u041f\u0440\u043e\u0432\u0435\u0440\u044f\u0442\u044c Designed By"},
    {"key": "check_checked", "type": "checkbox", "label": u"\u041f\u0440\u043e\u0432\u0435\u0440\u044f\u0442\u044c Checked By",
     "checkbox_label": u"\u041f\u0440\u043e\u0432\u0435\u0440\u044f\u0442\u044c Checked By"},
    {"key": "mode", "type": "select", "label": u"\u0420\u0435\u0436\u0438\u043c \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438",
     "options": [u"\u0421\u0442\u0440\u043e\u0433\u0438\u0439 (\u043e\u0431\u0430 \u043f\u043e\u043b\u044f \u043e\u0431\u044f\u0437\u0430\u0442\u0435\u043b\u044c\u043d\u044b)", u"\u041c\u044f\u0433\u043a\u0438\u0439 (\u0445\u043e\u0442\u044f \u0431\u044b \u043e\u0434\u043d\u043e)"]},
]

from pyrevit import revit, DB
from core.validation_engine import ValidationResult


def run(doc, section, project, settings=None):
    if not settings:
        settings = {}

    selected = settings.get("selected_sheets", [])

    if not selected:
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message=u"\u041d\u0435 \u0432\u044b\u0431\u0440\u0430\u043d\u043e \u043d\u0438 \u043e\u0434\u043d\u043e\u0433\u043e \u043b\u0438\u0441\u0442\u0430. \u041e\u0442\u043a\u0440\u043e\u0439\u0442\u0435 \u043d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438 \u0441\u043a\u0440\u0438\u043f\u0442\u0430.",
            skip_summary=True
        )

    check_designed = settings.get("check_designed", True)
    check_checked = settings.get("check_checked", True)
    mode = settings.get("mode", u"\u0421\u0442\u0440\u043e\u0433\u0438\u0439 (\u043e\u0431\u0430 \u043f\u043e\u043b\u044f \u043e\u0431\u044f\u0437\u0430\u0442\u0435\u043b\u044c\u043d\u044b)")

    all_sheets = (DB.FilteredElementCollector(doc)
                  .OfClass(DB.ViewSheet)
                  .WhereElementIsNotElementType()
                  .ToElements())
    sheet_map = {s.SheetNumber: s for s in all_sheets}

    problems = []
    checked = 0
    for num in selected:
        if num not in sheet_map:
            continue
        sheet = sheet_map[num]
        checked += 1

        has_designed = True
        has_checked = True

        if check_designed:
            designed = sheet.get_Parameter(DB.BuiltInParameter.SHEET_DESIGNED_BY)
            if designed and (not designed.AsString() or not designed.AsString().strip()):
                has_designed = False

        if check_checked:
            checked_by = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CHECKED_BY)
            if checked_by and (not checked_by.AsString() or not checked_by.AsString().strip()):
                has_checked = False

        if mode == u"\u0421\u0442\u0440\u043e\u0433\u0438\u0439 (\u043e\u0431\u0430 \u043f\u043e\u043b\u044f \u043e\u0431\u044f\u0437\u0430\u0442\u0435\u043b\u044c\u043d\u044b)":
            if not has_designed or not has_checked:
                problems.append(sheet)
        else:
            if not has_designed and not has_checked:
                problems.append(sheet)

    return ValidationResult(
        check_name=SCRIPT_NAME,
        passed=len(problems) == 0,
        message=u"\u041f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e \u043b\u0438\u0441\u0442\u043e\u0432: {}. \u041f\u0440\u043e\u0431\u043b\u0435\u043c\u043d\u044b\u0445: {}".format(checked, len(problems)),
        elements=[s.Id for s in problems],
        skip_summary=True
    )
