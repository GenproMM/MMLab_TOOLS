---
quick_id: 260724-win
status: complete
---

# Quick Task 260724-win: приёмка кнопки IFC_Окна

## Task
Принять и адаптировать сторонний скрипт классификации окон по МССК под
конвенцию MM LAB (процедура `/mm-adopt-script`).

- Источник: `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Окна.pushbutton/script.py`
  (лежал как `pending_adoption` в `tools/convention_baseline.json`).
- Панель: **АРХИТЕКТУРА** (категория `OST_Windows`, IFC-классификация окон).
- Полный аналог уже принятой кнопки `IFC_Двери` (те же общие параметры
  `GP_01_КодКлассифМССК` / `GP_01_ИмяКлассифМССК`, тот же WinForms-паттерн).

## Adaptation summary
- **MM001/MM002**: добавлены `#! python3` и `# -*- coding: utf-8 -*-`.
- **MM003**: файл записан UTF-8 без BOM.
- **MM004**: русский docstring со строками «Совместимость:» и «Зависимости:».
- **MM009**: `from Autodesk.Revit.DB import *` → явные импорты имён.
- **MM012**: `pyrevit.forms` (SelectFromList/alert) → WinForms `CheckedListBox`
  (`select_families`) и `TaskDialog` (`revit_ui_helpers.alert` / `confirm`).
- **MM010**: `LookupParameter("GP_01_…")` → `revit_compat.get_shared_parameter(GUID)`;
  `LookupParameter("Модель")` → `revit_compat.get_parameter(…, ALL_MODEL_MODEL, "Модель")`.
- **MM011**: все голые `except:` → `except Exception`.
- **MM014**: добавлен канонический lib-бутстрап `_SCRIPT_DIR`/`_EXTENSION_DIR`/`_LIB_DIR`.
- **MM005/MM006**: созданы `bundle.yaml` (title/tooltip ru+en) и `README.md`.
- **MM007**: кнопка зарегистрирована в `АРХИТЕКТУРА.panel/bundle.yaml` (после `IFC_Двери`).
- **D-03**: `main()` начинается с `require_supported_version(COMMAND_NAME)`.
- **Правило 10**: транзакция — `Start()` перед `try`, `Commit()` в `try`,
  `RollBack()`+`raise` в `except`.
- **Правило 18**: `doc` берётся в `_entry()` и передаётся в `main(doc)` аргументом.
- **2024-совместимость**: `Id.IntegerValue` → `revit_compat.element_id_value(Id)`.
- Иконка `icons8-окно-96.png` → `icon.png`.
- Функционально-эквивалентное упрощение: избыточный перебор
  `GetParameters("Модель")` свёрнут в `_model_has_balcony()` (экземпляр + тип);
  коды классификации сохранены дословно (`ЭЛ 30 18 09` / `ЭЛ 30 18 40`).

## Verify
- `py -3 tools/check_convention.py "…/IFC_Окна.pushbutton" --strict` → exit 0 (0/0).
- `py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json`
  → exit 0 (17 проверено, 0 ошибок, 0 предупреждений).
- Запись из `pending_adoption` удалена.
- Runtime-истина (UAT) — прогон кнопки в живом Revit после pyRevit Reload.
