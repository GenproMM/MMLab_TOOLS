# -*- coding: utf-8 -*-
"""Проверяет заполненность параметров 01_GP_КодЕдиницы, 01_GP_КодВидаРаботы_Тип
у семейств и 01_GP_КодЕдиницы_Мат у материалов."""

SCRIPT_ID = "vor_c5e8d4a2"
SCRIPT_NAME = u"\u0417\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u044c GP_01_\u041a\u043e\u0434\u044b"
SCRIPT_DESCRIPTION = u"\u041f\u0440\u043e\u0432\u0435\u0440\u044f\u0435\u0442 \u0437\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u044c \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u0432 01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b, 01_GP_\u041a\u043e\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043e\u0442\u044b_\u0422\u0438\u043f \u0443 \u0441\u0435\u043c\u0435\u0439\u0441\u0442\u0432 \u0438 01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b_\u041c\u0430\u0442 \u0443 \u043c\u0430\u0442\u0435\u0440\u0438\u0430\u043b\u043e\u0432."
HAS_SETTINGS = True

import os
import sys

from pyrevit import revit, DB
from core.validation_engine import ValidationResult


PARAM_FAMILY_CODE_UNIT = u"01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b"
PARAM_FAMILY_WORK_TYPE = u"01_GP_\u041a\u043e\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043e\u0442\u044b_\u0422\u0438\u043f"
PARAM_MATERIAL_CODE = u"01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b_\u041c\u0430\u0442"

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
        from settings_window import KodySettingsWindow
    finally:
        sys.path[:] = old_path
    dlg = KodySettingsWindow(doc, current_settings)
    return dlg.show_dialog()


# ================================================================
# Helpers
# ================================================================

def _detect_categories_for_params(doc):
    """Найти категории для трёх параметров. Возвращает:
    {"family": [(cat_name, BuiltInCategory)],
     "material": [(cat_name, BuiltInCategory)]}
    """
    result = {"family": [], "material": []}
    family_bics = set()
    material_bics = set()

    target_params = {
        PARAM_FAMILY_CODE_UNIT: "family",
        PARAM_FAMILY_WORK_TYPE: "family",
        PARAM_MATERIAL_CODE: "material",
    }

    try:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            definition = iterator.Key
            group = target_params.get(definition.Name)
            if group is None:
                continue
            binding = iterator.Current
            cat_set = binding.Categories
            cat_iter = cat_set.GetEnumerator()
            while cat_iter.MoveNext():
                cat = cat_iter.Current
                bic = cat.BuiltInCategory
                if group == "family" and bic not in family_bics:
                    family_bics.add(bic)
                    result["family"].append((cat.Name, bic))
                elif group == "material" and bic not in material_bics:
                    material_bics.add(bic)
                    result["material"].append((cat.Name, bic))
    except Exception:
        pass

    result["family"].sort(key=lambda x: x[0])
    result["material"].sort(key=lambda x: x[0])
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


def _get_elem_info(doc, elem, group):
    """Получить информацию об элементе.
    group: "family" или "material"
    """
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

    kod_ed = _get_param_value(elem, PARAM_FAMILY_CODE_UNIT) or ""
    kod_vr = _get_param_value(elem, PARAM_FAMILY_WORK_TYPE) or ""
    kod_mat = _get_param_value(elem, PARAM_MATERIAL_CODE) or ""

    if group == "family":
        ed_empty = _is_empty(kod_ed)
        vr_empty = _is_empty(kod_vr)
        if ed_empty and vr_empty:
            marker = "both"
        elif ed_empty:
            marker = "kod_ed"
        else:
            marker = "kod_vr"
    else:
        marker = "mat"

    return {
        "id": elem.Id,
        "unique_id": elem.UniqueId,
        "family": family_name,
        "type": type_name,
        "kod_ed": kod_ed,
        "kod_vr": kod_vr,
        "kod_mat": kod_mat,
        "empty_marker": marker,
    }


def _unique_id_to_element_id(doc, unique_id):
    """Конвертировать UniqueId в ElementId. None если элемент удалён."""
    try:
        elem = doc.GetElement(unique_id)
        if elem:
            return elem.Id
    except Exception:
        pass
    return None


def _load_excluded_ids(doc, settings):
    """Загрузить списки исключений из настроек.
    Возвращает {"family": (set(ElementId), list(str)),
                 "material": (set(ElementId), list(str))}
    """
    result = {}
    for key in ("excluded_family", "excluded_material"):
        group = key.replace("excluded_", "")
        uids = (settings or {}).get(key, [])
        valid_uids = []
        elem_ids = set()
        for uid in uids:
            eid = _unique_id_to_element_id(doc, uid)
            if eid is not None:
                elem_ids.add(eid)
                valid_uids.append(uid)
        result[group] = (elem_ids, valid_uids)
    return result


def _show_results_window(doc, results_data):
    """Показать плавающее окно результатов (non-modal)."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.join(script_dir, "lib")
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)

    import __main__
    attr_name = "_gp01kody_results_window"
    old_win = getattr(__main__, attr_name, None)
    if old_win:
        try:
            old_win.Close()
        except Exception:
            pass

    from results_window import KodyResultsWindow
    win = KodyResultsWindow(doc, results_data)
    setattr(__main__, attr_name, win)
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

        detected = _detect_categories_for_params(doc)
        if not detected["family"] and not detected["material"]:
            import System
            System.Windows.MessageBox.Show(
                u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b 01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b / 01_GP_\u041a\u043e\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043e\u0442\u044b_\u0422\u0438\u043f / 01_GP_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b_\u041c\u0430\u0442 \u043d\u0435 \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d\u044b \u043d\u0438 \u043a \u043e\u0434\u043d\u043e\u0439 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438.",
                SCRIPT_NAME,
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning
            )
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=False,
                message=u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b \u043d\u0435 \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d\u044b \u043d\u0438 \u043a \u043e\u0434\u043d\u043e\u0439 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438.",
                skip_summary=True,
            )

        excluded = _load_excluded_ids(doc, settings)
        excluded_family_ids, excluded_family_uids = excluded["family"]
        excluded_material_ids, excluded_material_uids = excluded["material"]

        family_problems = []
        family_exceptions = []
        material_problems = []
        material_exceptions = []

        total_family_checked = 0
        total_material_checked = 0

        # Семейства
        for cat_name, bic in detected["family"]:
            if not categories.get(cat_name, True):
                continue
            try:
                elements = (DB.FilteredElementCollector(doc)
                            .OfCategory(bic)
                            .WhereElementIsNotElementType()
                            .ToElements())
            except Exception:
                elements = []

            for elem in elements:
                total_family_checked += 1
                kod_ed = _get_param_value(elem, PARAM_FAMILY_CODE_UNIT)
                kod_vr = _get_param_value(elem, PARAM_FAMILY_WORK_TYPE)
                if not _is_empty(kod_ed) and not _is_empty(kod_vr):
                    continue

                info = _get_elem_info(doc, elem, "family")
                if elem.Id in excluded_family_ids:
                    family_exceptions.append(info)
                else:
                    family_problems.append(info)

        # Материалы
        for cat_name, bic in detected["material"]:
            if not categories.get(cat_name, True):
                continue
            try:
                elements = (DB.FilteredElementCollector(doc)
                            .OfCategory(bic)
                            .WhereElementIsNotElementType()
                            .ToElements())
            except Exception:
                elements = []

            for elem in elements:
                total_material_checked += 1
                kod_mat = _get_param_value(elem, PARAM_MATERIAL_CODE)
                if not _is_empty(kod_mat):
                    continue

                info = _get_elem_info(doc, elem, "material")
                if elem.Id in excluded_material_ids:
                    material_exceptions.append(info)
                else:
                    material_problems.append(info)

        results_data = {
            "family": {
                "problems": family_problems,
                "exceptions": family_exceptions,
            },
            "material": {
                "problems": material_problems,
                "exceptions": material_exceptions,
            },
            "excluded_uids": {
                "family": excluded_family_uids,
                "material": excluded_material_uids,
            },
            "script_name": SCRIPT_NAME,
            "section": section,
            "project": project,
        }

        has_any = (family_problems or family_exceptions
                   or material_problems or material_exceptions)
        # Кэшируем данные прогона для show_results (окно открывается
        # по кнопке «Открыть результат» в окне прогона, а не здесь).
        _last_results_data = results_data

        # Сводка
        parts = []
        if total_family_checked > 0:
            parts.append(
                u"\u0421\u0435\u043c\u0435\u0439\u0441\u0442\u0432\u0430: \u043f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e {}, "
                u"\u043f\u0440\u043e\u0431\u043b\u0435\u043c {}, \u0438\u0441\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0439 {}".format(
                    total_family_checked,
                    len(family_problems),
                    len(family_exceptions),
                )
            )
        if total_material_checked > 0:
            parts.append(
                u"\u041c\u0430\u0442\u0435\u0440\u0438\u0430\u043b\u044b: \u043f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e {}, "
                u"\u043f\u0440\u043e\u0431\u043b\u0435\u043c {}, \u0438\u0441\u043a\u043b\u044e\u0447\u0435\u043d\u0438\u0439 {}".format(
                    total_material_checked,
                    len(material_problems),
                    len(material_exceptions),
                )
            )

        if not parts:
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=True,
                message=u"\u041d\u0435\u0442 \u044d\u043b\u0435\u043c\u0435\u043d\u0442\u043e\u0432 \u0434\u043b\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438.",
                skip_summary=True,
            )

        result_msg = u"\n".join(parts)
        total_problems = len(family_problems) + len(material_problems)

        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=(total_problems == 0),
            message=result_msg,
            elements=[e["id"] for e in family_problems]
                     + [e["id"] for e in material_problems],
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
    rd = _last_results_data
    fam = rd.get("family", {}) if isinstance(rd, dict) else {}
    mat = rd.get("material", {}) if isinstance(rd, dict) else {}
    has_any = (fam.get("problems") or fam.get("exceptions")
               or mat.get("problems") or mat.get("exceptions"))
    if not has_any:
        import System
        System.Windows.MessageBox.Show(
            u"\u041f\u0440\u043e\u0431\u043b\u0435\u043c \u043d\u0435\u0442.",
            SCRIPT_NAME,
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Information
        )
        return
    _show_results_window(doc, rd)
