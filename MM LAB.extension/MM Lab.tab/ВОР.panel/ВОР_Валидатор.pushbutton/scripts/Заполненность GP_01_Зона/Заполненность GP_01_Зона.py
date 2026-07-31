# -*- coding: utf-8 -*-
"""Проверяет заполненность параметра GP_01_Зона у элементов."""

SCRIPT_ID = "vor_a3f2c891"
SCRIPT_NAME = u"\u0417\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u044c GP_01_\u0417\u043e\u043d\u0430"
SCRIPT_DESCRIPTION = u"\u041f\u0440\u043e\u0432\u0435\u0440\u044f\u0435\u0442 \u0437\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u044c \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u0430 GP_01_\u0417\u043e\u043d\u0430 \u0443 \u044d\u043b\u0435\u043c\u0435\u043d\u0442\u043e\u0432 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0439, \u043a \u043a\u043e\u0442\u043e\u0440\u044b\u043c \u043e\u043d \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d."
HAS_SETTINGS = True

import os
import sys

from pyrevit import revit, DB
from core.validation_engine import ValidationResult

PARAM_NAME = u"GP_01_\u0417\u043e\u043d\u0430"

# Кэш данных последнего прогона — используется show_results для повторного
# открытия окна результатов без перевыполнения проверки.
_last_results_data = None


def show_settings(doc, current_settings):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.join(script_dir, "lib")
    old_path = list(sys.path)
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
    if "settings_window" in sys.modules:
        del sys.modules["settings_window"]
    try:
        from settings_window import ZoneSettingsWindow
    finally:
        sys.path[:] = old_path
    dlg = ZoneSettingsWindow(doc, current_settings)
    return dlg.show_dialog()


# ================================================================
# Helpers
# ================================================================

def _detect_categories_with_bic(doc):
    """Найти категории с параметром PARAM_NAME. Возвращает [(cat_name, BuiltInCategory)]."""
    result = []
    try:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            definition = iterator.Key
            if definition.Name == PARAM_NAME:
                binding = iterator.Current
                cat_set = binding.Categories
                cat_iter = cat_set.GetEnumerator()
                while cat_iter.MoveNext():
                    cat = cat_iter.Current
                    result.append((cat.Name, cat.BuiltInCategory))
                break
    except Exception:
        pass
    result.sort(key=lambda x: x[0])
    return result


def _get_param_value(element, param_name):
    """Получить строковое значение параметра по имени."""
    for p in element.Parameters:
        if p.Definition.Name == param_name:
            val = p.AsString()
            if val is None:
                val = p.AsValueString()
            return val
    return None


def _is_empty(value):
    """Проверить что значение параметра пустое."""
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def _get_elem_info(doc, elem):
    """Получить {"id": ElementId, "family": str, "type": str} элемента."""
    type_name = ""
    family_name = ""

    try:
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            type_name = elem_type.Name or ""
    except Exception:
        pass

    try:
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            fam = elem_type.Family
            if fam:
                family_name = fam.Name or ""
    except Exception:
        pass

    if not family_name:
        try:
            p = elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
            if p:
                family_name = p.AsValueString() or ""
        except Exception:
            pass

    if not family_name:
        try:
            for p in elem.Parameters:
                if p.StorageType == DB.StorageType.ElementId:
                    eid = p.AsElementId()
                    if eid and eid.IntegerValue > 0:
                        ref = doc.GetElement(eid)
                        if ref and isinstance(ref, DB.Family):
                            family_name = ref.Name or ""
                            break
        except Exception:
            pass

    if not type_name:
        try:
            type_name = elem.Name or ""
        except Exception:
            pass

    if not family_name and elem.Category:
        try:
            family_name = elem.Category.Name or ""
        except Exception:
            pass

    param_val = _get_param_value(elem, PARAM_NAME) or ""
    return {"id": elem.Id, "family": family_name, "type": type_name, "param_value": param_val}


def _show_results_window(doc, results_data):
    """Показать плавающее окно результатов (non-modal)."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.join(script_dir, "lib")
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)

    import __main__
    old_win = getattr(__main__, '_gp01zona_results_window', None)
    if old_win:
        try:
            old_win.Close()
        except Exception:
            pass

    from results_window import ResultsWindow
    win = ResultsWindow(doc, results_data)
    __main__._gp01zona_results_window = win
    win.Show()


# ================================================================
# Main validation
# ================================================================

def run(doc, section, project, settings=None):
    global _last_results_data
    try:
        if not settings:
            settings = {}
        categories = settings.get("categories", {})

        detected = _detect_categories_with_bic(doc)
        if not detected:
            import System
            System.Windows.MessageBox.Show(
                u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 GP_01_\u0417\u043e\u043d\u0430 \u043d\u0435 \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d \u043d\u0438 \u043a \u043e\u0434\u043d\u043e\u0439 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 \u0432 \u043f\u0440\u043e\u0435\u043a\u0442\u0435.",
                SCRIPT_NAME,
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning
            )
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=False,
                message=u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 GP_01_\u0417\u043e\u043d\u0430 \u043d\u0435 \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d \u043d\u0438 \u043a \u043e\u0434\u043d\u043e\u0439 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 \u0432 \u043f\u0440\u043e\u0435\u043a\u0442\u0435.",
                skip_summary=True,
            )

        total_checked = 0
        total_problems = 0
        all_problem_ids = []
        results_data = []

        for cat_name, bic in detected:
            if not categories.get(cat_name, True):
                continue

            try:
                elements = (DB.FilteredElementCollector(doc)
                            .OfCategory(bic)
                            .WhereElementIsNotElementType()
                            .ToElements())
            except Exception:
                elements = []

            problems_in_cat = 0
            checked_in_cat = 0
            problem_elements = []

            for elem in elements:
                checked_in_cat += 1
                value = _get_param_value(elem, PARAM_NAME)
                if _is_empty(value):
                    problems_in_cat += 1
                    all_problem_ids.append(elem.Id)
                    problem_elements.append(_get_elem_info(doc, elem))

            total_checked += checked_in_cat
            total_problems += problems_in_cat

            results_data.append({
                "category": cat_name,
                "checked": checked_in_cat,
                "problems": problems_in_cat,
                "elements": problem_elements,
            })

        # Кэшируем данные прогона для show_results (окно открывается
        # по кнопке «Открыть результат» в окне прогона, а не здесь).
        _last_results_data = results_data

        messages = []
        for rd in results_data:
            if rd["checked"] > 0:
                messages.append(
                    u"{}: \u043f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e {}, "
                    u"\u043f\u0440\u043e\u0431\u043b\u0435\u043c {}".format(
                        rd["category"], rd["checked"], rd["problems"]
                    )
                )

        if not messages:
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=True,
                message=u"\u041d\u0435\u0442 \u044d\u043b\u0435\u043c\u0435\u043d\u0442\u043e\u0432 \u0434\u043b\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438.",
                skip_summary=True,
            )

        result_msg = u"\n".join(messages)
        result_msg += (
            u"\n\u0418\u0442\u043e\u0433\u043e: \u043f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e {}, "
            u"\u043f\u0440\u043e\u0431\u043b\u0435\u043c {}".format(
                total_checked, total_problems
            )
        )

        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=(total_problems == 0),
            message=result_msg,
            elements=all_problem_ids,
            skip_summary=True,
        )

    except Exception as e:
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message=u"\u041e\u0448\u0438\u0431\u043a\u0430: {}".format(str(e)),
            skip_summary=True,
        )


def show_results(doc, section, project, settings=None):
    """Открыть окно результатов последнего прогона.

    Вызывается окном прогона (run_window) по кнопке «Открыть результат».
    Данные берутся из кэша _last_results_data (заполняется в run()).
    """
    global _last_results_data
    if not _last_results_data:
        return
    # Есть ли вообще проблемы для показа?
    has_problems = any(
        rd.get("problems", 0) > 0 for rd in _last_results_data
    )
    if not has_problems:
        import System
        System.Windows.MessageBox.Show(
            u"\u041f\u0440\u043e\u0431\u043b\u0435\u043c \u043d\u0435\u0442.",
            SCRIPT_NAME,
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Information
        )
        return
    _show_results_window(doc, _last_results_data)
