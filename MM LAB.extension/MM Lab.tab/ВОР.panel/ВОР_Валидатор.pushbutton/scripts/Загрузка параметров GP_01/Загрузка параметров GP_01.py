# -*- coding: utf-8 -*-
# Загрузка общих параметров GP_01 из файла ГП_ФОП2025.txt
# и подключение их к нужным категориям элементов и материалов.

SCRIPT_ID = "vor_8c4f2a1e"
# "Загрузка параметров GP_01"
SCRIPT_NAME = u"\u0417\u0430\u0433\u0440\u0443\u0437\u043a\u0430 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u0432 GP_01"
# "Загружает общие параметры GP_01 из ФОП и подключает к категориям."
SCRIPT_DESCRIPTION = u"\u0417\u0430\u0433\u0440\u0443\u0436\u0430\u0435\u0442 \u043e\u0431\u0449\u0438\u0435 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b GP_01 \u0438\u0437 \u0424\u041e\u041f \u0438 \u043f\u043e\u0434\u043a\u043b\u044e\u0447\u0430\u0435\u0442 \u043a \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f\u043c."

import os
from pyrevit import revit, DB
from core.validation_engine import ValidationResult

# \\Srv-dfs\BIM\01_Ресурсы плагинов\2_Shared_parameters\ГП_ФОП2025.txt
SHARED_PARAMS_PATH = u"\\\\Srv-dfs\\BIM\\01_\u0420\u0435\u0441\u0443\u0440\u0441\u044b \u043f\u043b\u0430\u0433\u0438\u043d\u043e\u0432\\2_Shared_parameters\\\u0413\u041f_\u0424\u041e\u041f2025.txt"

# Основные категории для параметров ВОР
MAIN_CATEGORIES = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_StairsRailing,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_GenericModel,
    DB.BuiltInCategory.OST_Cornices,
    DB.BuiltInCategory.OST_SpecialityEquipment,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_MechanicalEquipment,
    DB.BuiltInCategory.OST_Stairs,
    DB.BuiltInCategory.OST_StairsRuns,
    DB.BuiltInCategory.OST_StairsLandings,
    DB.BuiltInCategory.OST_Ramps,
]

# (имя_параметра, экземплярный, [категории])
# GP_01_КодВидаРаботы_Экз, GP_01_КодВидаРаботы_Тип,
# GP_01_КодЕдиницы_Мат, GP_01_КодЕдиницы, GP_01_ВидРаботы
PARAMETERS = [
    (u"GP_01_\u041a\u043e\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043e\u0442\u044b_\u042d\u043a\u0437", True, MAIN_CATEGORIES),
    (u"GP_01_\u041a\u043e\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043e\u0442\u044b_\u0422\u0438\u043f", False, MAIN_CATEGORIES),
    (u"GP_01_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b_\u041c\u0430\u0442", True, [DB.BuiltInCategory.OST_Materials]),
    (u"GP_01_\u041a\u043e\u0434\u0415\u0434\u0438\u043d\u0438\u0446\u044b", False, MAIN_CATEGORIES),
    (u"GP_01_\u0412\u0438\u0434\u0420\u0430\u0431\u043e\u0442\u044b", True, MAIN_CATEGORIES),
]


def _find_definition(def_file, param_name):
    """Find ExternalDefinition by name in shared parameters file."""
    group_iter = def_file.Groups.GetEnumerator()
    while group_iter.MoveNext():
        group = group_iter.Current
        defn_iter = group.Definitions.GetEnumerator()
        while defn_iter.MoveNext():
            defn = defn_iter.Current
            if defn.Name == param_name:
                return defn
    return None


def _get_binding_info(binding_map, param_name):
    """Get existing binding info. Returns (categories_set, is_instance) or (None, None)."""
    iterator = binding_map.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        if iterator.Key.Name == param_name:
            binding = iterator.Current
            cats = set()
            cat_enum = binding.Categories.GetEnumerator()
            while cat_enum.MoveNext():
                cats.add(cat_enum.Current.BuiltInCategory)
            is_inst = binding.GetType().Name == "InstanceBinding"
            return cats, is_inst
    return None, None


def run(doc, section, project, settings=None):
    """Load and bind GP_01 shared parameters."""
    try:
        app = doc.Application
        errors = []
        warnings = []
        bound_count = 0
        already_ok = 0

        # "Файл общих параметров не найден"
        if not os.path.exists(SHARED_PARAMS_PATH):
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=False,
                message=u"\u0424\u0430\u0439\u043b \u043e\u0431\u0449\u0438\u0445 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u0432 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d: " + SHARED_PARAMS_PATH,
                elements=[],
                skip_summary=True
            )

        # Ensure shared parameters file path is set
        path_changed = False
        current_path = app.SharedParametersFilename or u""
        if current_path.lower() != SHARED_PARAMS_PATH.lower():
            app.SharedParametersFilename = SHARED_PARAMS_PATH
            path_changed = True

        # Open shared parameters file
        def_file = app.OpenSharedParameterFile()
        if def_file is None:
            # "Не удалось открыть файл общих параметров"
            return ValidationResult(
                check_name=SCRIPT_NAME,
                passed=False,
                message=u"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043e\u0442\u043a\u0440\u044b\u0442\u044c \u0444\u0430\u0439\u043b \u043e\u0431\u0449\u0438\u0445 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u0432",
                elements=[],
                skip_summary=True
            )

        binding_map = doc.ParameterBindings

        for param_name, is_instance, required_cats in PARAMETERS:
            # Find definition in shared parameters file
            definition = _find_definition(def_file, param_name)
            if definition is None:
                # "Параметр '{}' не найден в файле общих параметров"
                errors.append(
                    u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 '{}' \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d \u0432 \u0444\u0430\u0439\u043b\u0435 \u043e\u0431\u0449\u0438\u0445 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u0432".format(param_name)
                )
                continue

            # Check existing binding
            bound_cats, bound_is_inst = _get_binding_info(binding_map, param_name)

            if bound_cats is not None:
                required_set = set(required_cats)
                if required_set.issubset(bound_cats) and bound_is_inst == is_instance:
                    already_ok += 1
                    continue
                else:
                    # "Параметр '{}' привязан с другими категориями"
                    warnings.append(
                        u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 '{}' \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d \u0441 \u0434\u0440\u0443\u0433\u0438\u043c\u0438 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f\u043c\u0438".format(param_name)
                    )
                    continue

            # Create category set and binding
            cat_set = app.Create.NewCategorySet()
            for bic in required_cats:
                cat = doc.Settings.Categories.get_Item(bic)
                cat_set.Insert(cat)

            if is_instance:
                binding = app.Create.NewInstanceBinding(cat_set)
            else:
                binding = app.Create.NewTypeBinding(cat_set)

            # Insert binding in transaction
            t = DB.Transaction(doc, "Bind shared parameter")
            t.Start()
            try:
                success = binding_map.Insert(definition, binding)
                if success:
                    t.Commit()
                    bound_count += 1
                else:
                    t.RollBack()
                    # "Не удалось привязать параметр '{}'"
                    errors.append(
                        u"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u0442\u044c \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 '{}'".format(param_name)
                    )
            except Exception:
                t.RollBack()
                errors.append(
                    u"\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u0442\u044c \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 '{}'".format(param_name)
                )

        # Build result message
        parts = []
        if path_changed:
            # "Путь к файлу общих параметров обновлён"
            parts.append(u"\u041f\u0443\u0442\u044c \u043a \u0444\u0430\u0439\u043b\u0443 \u043e\u0431\u0449\u0438\u0445 \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u0432 \u043e\u0431\u043d\u043e\u0432\u043b\u0451\u043d")
        if bound_count > 0:
            # "Привязано: {}"
            parts.append(u"\u041f\u0440\u0438\u0432\u044f\u0437\u0430\u043d\u043e: {}".format(bound_count))
        if already_ok > 0:
            # "Уже привязаны: {}"
            parts.append(u"\u0423\u0436\u0435 \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d\u044b: {}".format(already_ok))
        for w in warnings:
            parts.append(u"! " + w)
        for e in errors:
            parts.append(u"X " + e)

        passed = len(errors) == 0
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=passed,
            message=u"\n".join(parts) if parts else u"OK",
            elements=[],
            skip_summary=True
        )

    except Exception as e:
        # "Ошибка: {}"
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message=u"\u041e\u0448\u0438\u0431\u043a\u0430: {}".format(str(e)),
            elements=[],
            skip_summary=True
        )
