---
quick_id: 260831-flr
status: complete
---

# Quick Task 260831-flr: приёмка кнопки IFC_Перекрытия

## Task
Принять и адаптировать сторонний скрипт классификации перекрытий по МССК под
конвенцию MM LAB (процедура `/mm-adopt-script`).

- Источник: `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Перекрытия.pushbutton/script.py`
  (лежал в рабочем дереве untracked, записи в `convention_baseline.json` не было).
- Панель: **АРХИТЕКТУРА** (`FloorType` / `OST_Floors`, IFC-классификация перекрытий) —
  подтверждено пользователем.
- Аналог уже принятых кнопок `IFC_Двери` и `IFC_Окна` (те же общие параметры
  `GP_01_КодКлассифМССК` / `GP_01_ИмяКлассифМССК`, тот же WinForms-паттерн).

## Adaptation summary
- **MM001/MM002**: добавлены `#! python3` и `# -*- coding: utf-8 -*-`.
- **MM003**: файл записан UTF-8 без BOM.
- **MM004**: русский docstring со строками «Совместимость:» и «Зависимости:»;
  опечатка `__autor__` исправлена на `__author__ = "GENPRO LAB"`.
- **MM008/MM012**: `from pyrevit import revit, DB, forms` → явные импорты
  `Autodesk.Revit.DB/UI` и `System.Windows.Forms`; `forms.SelectFromList` →
  WinForms `CheckedListBox` (`select_floor_types`), `forms.alert` →
  `TaskDialog` (`revit_ui_helpers.alert` / `confirm`).
- **MM010**: `LookupParameter("GP_11_Группирование")` убран (см. изменение
  поведения); чтение имени типоразмера — `revit_compat.get_parameter(…,
  SYMBOL_NAME_PARAM)` вместо прямого `get_Parameter(bip)` (D-04, pythonnet 3.x).
- **MM005/MM006**: созданы `bundle.yaml` (title/tooltip ru+en) и `README.md`.
- **MM007**: кнопка зарегистрирована в `АРХИТЕКТУРА.panel/bundle.yaml`
  (после `IFC_Окна`).
- **MM014**: добавлен канонический lib-бутстрап `_SCRIPT_DIR`/`_EXTENSION_DIR`/`_LIB_DIR`.
- **D-03**: код верхнего уровня свёрнут в `main(doc)`, которая начинается с
  `revit_compat.require_supported_version(COMMAND_NAME)`.
- **Правило 10**: транзакция — `Start()` перед `try`, `Commit()` в `try`,
  `RollBack()`+`raise` в `except` (в исходнике `except` глотал ошибку и
  показывал `forms.alert`).
- **Правило 11/18**: верхний уровень обёрнут в `try/except Exception` с
  `TaskDialog`; `doc` берётся в `_entry()` и передаётся в `main(doc)` аргументом.
- **2024-совместимость**: сравнение `Id` / `GetTypeId()` через
  `revit_compat.element_id_value` вместо прямого сравнения ElementId.
- Иконка исходника оставлена как `icon.png`.

## Изменение поведения (решение пользователя)
Исходник писал код `ЭЛ 30 10 40` в «Комментарии» (`ALL_MODEL_INSTANCE_COMMENTS`)
и «Перекрытие» в `GP_11_Группирование`, при этом `__doc__` обещал GP_01.
По решению пользователя приведено к обещанному и к соседним кнопкам:
`ЭЛ 30 10 40` → `GP_01_КодКлассифМССК`, «Перекрытие» → `GP_01_ИмяКлассифМССК`
(GUID те же, что в `IFC_Двери`/`IFC_Окна`). «Комментарии» и `GP_11` больше
не трогаются; docstring, tooltip и README приведены к фактическому поведению.

## Verify
- `py -3 tools/check_convention.py "…/IFC_Перекрытия.pushbutton" --strict`
  → exit 0 (0 ошибок, 0 предупреждений).
- `py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json`
  → exit 1, но **ни одного нарушения по IFC_Перекрытия**: остаток —
  предсуществующий долг `IFC_Стены.pushbutton` (untracked, ещё не принята)
  и `СНиП` / `СНиП_ФОП25` (нет записей в baseline). Приёмкой не затронут.
- Runtime-истина (UAT) — прогон кнопки в живом Revit после pyRevit Reload.
