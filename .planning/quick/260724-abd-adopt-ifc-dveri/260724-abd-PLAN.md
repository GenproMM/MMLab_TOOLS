---
quick_id: 260724-abd
status: complete
---

# Quick Task 260724-abd: приёмка кнопки IFC_Двери

## Task
Принять сторонний скрипт `IFC_Двери.pushbutton` (панель АРХИТЕКТУРА) в конвенцию
MM LAB через `/mm-adopt-script`. Исходник классифицирует двери по МССК и заполняет
общие параметры `GP_01_КодКлассифМССК` / `GP_01_ИмяКлассифМССК` по назначению
помещений с обеих сторон двери (FromRoom/ToRoom, `GP_23_Назначение`), с отдельным
распознаванием ворот и люков.

## Adaptation summary
- MM001/MM002: шапка `#! python3` + `# -*- coding: utf-8 -*-`.
- MM004: docstring с «Совместимость: Revit 2020 / 2022 / 2024» и «Зависимости: нет».
- MM009: убран `from Autodesk.Revit.DB import *` — явные импорты.
- MM012: `pyrevit.forms` убран — `forms.SelectFromList` → WinForms `CheckedListBox`
  (`select_families`), `forms.alert` → `revit_ui_helpers.alert` / `TaskDialog` (`confirm`).
- MM011: все голые `except:` → `except Exception:`.
- MM010: `LookupParameter("GP_01_КодКлассифМССК"/"GP_01_ИмяКлассифМССК"/"GP_23_Назначение")`
  → `revit_compat.get_shared_parameter(el, GUID)`. GUID трёх параметров взяты из ФОП
  `ГП_ФОП2025.txt` (предоставлен пользователем):
  - `GP_01_КодКлассифМССК` → `4df18cfa-3e5b-4e84-aef6-5ac3385a7d4f`
  - `GP_01_ИмяКлассифМССК` → `91f5b762-e361-462c-8611-3d952be1777b`
  - `GP_23_Назначение` → `0b3dbc34-30a7-4278-b1c5-8ba8819f9db4`
  `LookupParameter("Модель")` → `revit_compat.get_parameter(el, BuiltInParameter.ALL_MODEL_MODEL, "Модель")`.
- MM014: канонический lib-бутстрап (3 уровня `..`) вместо кастомного.
- MM005/MM006: добавлены `bundle.yaml` и `README.md`.
- Правило 9/10/18 AGENTS.md: `revit_compat.require_supported_version` в начале `main()`;
  транзакция `Start()` до `try`, `RollBack()+raise` в `except`; `doc` передаётся
  параметром вместо модульного глобала.
- Иконка `icons8-дверь-100.png` → `icon.png`.

## Verify
```bash
py -3 tools/check_convention.py "MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Двери.pushbutton" --strict
py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json
```
Оба — exit 0, 0 ошибок, 0 предупреждений.
