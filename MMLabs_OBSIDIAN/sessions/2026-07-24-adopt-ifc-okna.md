---
tags: [session, convention, adoption]
date: 2026-07-24
---

# 2026-07-24 — приёмка кнопки IFC_Окна под конвенцию MM LAB

## Что сделано
Сторонний скрипт классификации окон по МССК принят и адаптирован под конвенцию MM LAB.

## Ключевые файлы
- `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Окна.pushbutton/script.py` — адаптированный скрипт
- `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Окна.pushbutton/bundle.yaml` — подпись и подсказка
- `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Окна.pushbutton/README.md` — документация
- `.planning/quick/260724-win-adopt-ifc-okna/` — quick task (PLAN + SUMMARY)
- Коммит: `c85e0f8` (приёмка: кнопка IFC_Окна адаптирована под конвенцию MM LAB)

## Фактическая работа
1. **Исходный отчёт чекера** (шаг 2 `/mm-adopt-script`): 16 нарушений (3 ошибки, 13 предупреждений) — MM001/002/004/009/010/011/012, множественные `except:` и `LookupParameter("...")`.

2. **Адаптация (шаг 5)**: замена на конвенцию согласно AGENTS.md.
   - Шапка: `#! python3` и `# -*- coding: utf-8 -*-` на первые две строки.
   - Docstring: русский с «Совместимость: 2020/2022/2024» и «Зависимости: нет».
   - Импорты: `from Autodesk.Revit.DB import *` → явные имена (MM009).
   - Диалоги: `pyrevit.forms.SelectFromList` / `forms.alert` → WinForms `CheckedListBox` + `TaskDialog` (как в соседней `IFC_Двери`), см. [[WinForms-паттерн для множественного выбора]].
   - Параметры: все `LookupParameter("строка")` → `revit_compat.get_shared_parameter(GUID)` (GP_01_*) или `revit_compat.get_parameter(..., BuiltInParameter.ALL_MODEL_MODEL, "Модель")` (MM010).
   - Обработка ошибок: `except:` → `except Exception` (7 мест, MM011).
   - lib-бутстрап: добавлен канонический блок `_SCRIPT_DIR`/`_EXTENSION_DIR`/`_LIB_DIR` (MM014, D-15).
   - Версия Revit: `main()` начинается с `require_supported_version(COMMAND_NAME)` (D-03).
   - Транзакция: `Start()` вне `try`, `Commit()` в `try`, `RollBack()+raise` в `except` (правило 10).
   - doc/uidoc: в `_entry()` получаются из `__revit__`, передаются в `main(doc)` аргументом (правило 18).
   - ElementId: `Id.IntegerValue` → `revit_compat.element_id_value(Id)` (2024-совместимость).
   - Иконка: `icons8-окно-96.png` → `icon.png`.

3. **РЕВЬЮ-ГЕЙТ (шаг 7)**: показана сводка правок с MM-кодами, полный diff новых файлов; пользователь одобрил.

4. **Регистрация (шаг 8)**: `IFC_Окна` добавлена в `layout:` файла `АРХИТЕКТУРА.panel/bundle.yaml` (после `IFC_Двери`); запись удалена из `pending_adoption` в `tools/convention_baseline.json`.

5. **Гейт-проверка**: `py -3 tools/check_convention.py "…/IFC_Окна.pushbutton" --strict` → exit 0 (0/0); общий прогон `--all --baseline` → exit 0 (17 проверено).

6. **Quick task**: `.planning/quick/260724-win-adopt-ifc-okna/` (260724-win-PLAN.md + 260724-win-SUMMARY.md).

7. **Коммит (шаг 10)**: `c85e0f8`, пофайловый стейджинг (D-17); push не выполнялся (D-18 — ждёт подтверждения).

## Параллельная сессия
Во время работы конкурентный процесс закоммитил соседнюю кнопку `IFC_Двери` (`a92c955`) — тесная взаимозависимость: та же панель, те же общие параметры GP_01, тот же WinForms-паттерн. STATE.md при этом захватил мою строку `260724-win` и уже заказана она в HEAD. Энтанглмент разрешился чистым образом: IFC_Двери и IFC_Окна — отдельные коммиты, оба зарегистрированы в layout'е.

## Ключевые решения
- **Зеркало паттерна IFC_Двери**: функционально идентичная кнопка дала готовый шаблон адаптации (WinForms + revit_compat + `revit_ui_helpers.alert`).
- **Упрощение детекции балконного блока**: исходная логика дублировала проверку `LookupParameter("Модель")` через `GetParameters("Модель")` — свёрнуто в helper `_model_has_balcony()` (экземпляр + тип через `revit_compat.get_parameter`).
- **Строгая коммит-политика**: три общих файла (`bundle.yaml`, `baseline.json`, `STATE.md`) содержали работу двух сессий — выбрана чистая позиция «Только IFC_Окна» (IFC_Двери уже закоммичена отдельно).

## Результат
Кнопка принята, адаптирована, зарегистрирована, чистится по коду, залогирована в quick task. Runtime-истина (UAT) отложена: требует pyRevit Reload и прогона в живом Revit (2020/2022/2024).

## Ссылки
- [[Конвенция MM LAB скриптов]] (AGENTS.md)
- [[Процедура приёмки mm-adopt-script]]
- [[WinForms-паттерн для множественного выбора]] (IFC_Двери)
- [[Общие параметры GP_01_КодКлассифМССК и GP_01_ИмяКлассифМССК]]
