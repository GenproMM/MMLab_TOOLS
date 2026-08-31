---
tags: [session, adoption, convention]
date: 2026-08-31
---

# 2026-08-31 - Приёмка кнопки IFC_Перекрытия под конвенцию MM LAB

## Что сделано

Завершена приёмка сторонней кнопки классификации перекрытий по МССК
в процедуре `/mm-adopt-script`:

1. **Статическая проверка:** утилита `tools/check_convention.py` выявила
   нарушения MM001, MM002, MM004, MM005, MM006, MM007, MM008, MM010, MM012.

2. **Решения пользователя (шаг 3–4):**
   - Панель: **АРХИТЕКТУРА** (категория OST_Floors, IFC, рядом с IFC_Двери/Окна).
   - Параметры: переписать на `GP_01_*` вместо «Комментарии»/`GP_11_Группирование`.
   - Состав: только `GP_01_КодКлассифМССК` + `GP_01_ИмяКлассифМССК`, 
     `GP_11` и «Комментарии» не трогаются.

3. **Адаптация (шаг 5):**
   - Шапка: `#! python3` + `# -*- coding: utf-8 -*-` + docstring с
     «Совместимость/Зависимости», исправлена опечатка `__autor__` → `__author__`.
   - Импорты: `pyrevit.forms` (SelectFromList/alert) заменены явными импортами
     `Autodesk.Revit.DB/UI` + `System.Windows.Forms` + WinForms-диалоги
     `CheckedListBox` + `TaskDialog` (паттерн IFC_Двери).
   - Параметры: `LookupParameter("строка")` убран, имя типоразмера читается
     через `revit_compat.get_parameter(…, SYMBOL_NAME_PARAM)`, общие параметры
     пишутся через `revit_compat.get_shared_parameter(…, GUID)`.
   - Каркас: код свёрнут в `main(doc)` с `require_supported_version` в начале,
     `_entry()` готовит `doc` и передаёт параметром (правило 18).
   - Транзакция: `Start()` перед `try`, `RollBack()`+`raise` вместо глотания
     ошибок в `except` (правило 10).
   - ElementId: сравнения через `revit_compat.element_id_value` (Revit 2024+).
   - Созданы `bundle.yaml` (title ru/en, tooltip ru/en), `README.md` по шаблону.

4. **Гейт-проверка (шаг 6):** `py -3 tools/check_convention.py … --strict`
   → **exit 0** (0 ошибок, 0 предупреждений).

5. **РЕВЬЮ-ГЕЙТ (шаг 7):** пользователь одобрил правки («да»).

6. **Регистрация (шаг 8):**
   - Кнопка добавлена в `АРХИТЕКТУРА.panel/bundle.yaml` в layout
     (после `IFC_Окна`).
   - Общий прогон `--all --baseline` — exit 1, но ни одной строки по
     `IFC_Перекрытия` (долг — `IFC_Стены` и два СНиП, не в scope этой сессии).

7. **Quick task (шаг 9):** заведены артефакты `260831-flr-PLAN.md` и
   `260831-flr-SUMMARY.md` в `.planning/quick/260831-flr-adopt-ifc-perekrytiya/`,
   строка добавлена в `STATE.md`.

8. **Коммит (шаг 10):** пофайловый стейджинг 8 файлов (button, panel layout,
   quick task, STATE). Хеш `a01ace7`.

## Ключевые файлы

| Файл | Статус |
|------|--------|
| `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Перекрытия.pushbutton/script.py` | ✅ созданы |
| `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Перекрытия.pushbutton/bundle.yaml` | ✅ созданы |
| `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Перекрытия.pushbutton/README.md` | ✅ созданы |
| `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Перекрытия.pushbutton/icon.png` | ✅ скопирована |
| `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/bundle.yaml` | ✅ обновлена |
| `.planning/quick/260831-flr-adopt-ifc-perekrytiya/*` | ✅ созданы |
| `.planning/STATE.md` | ✅ обновлена |

## Решения и обоснование

1. **Переписал на GP_01:** исходник обещал `GP_01_*`, но писал куда-то ещё.
   По уточнению оказалось, что пользователь хотел фактически адаптировать к
   соседним кнопкам (IFC_Двери/Окна). Решено: код и имя МССК в `GP_01_*`,
   старые целевые параметры исключены.

2. **WinForms вместо pyrevit.forms:** MM012 — форм не работают под CPython
   (pyRevit 5+). Использовал паттерн из IFC_Двери — `CheckedListBox` +
   `TaskDialog` вместо `SelectFromList` (ворнинг снят).

3. **revit_compat.get_parameter для SYMBOL_NAME_PARAM:** прямой вызов
   `get_Parameter(BuiltInParameter)` падает на pythonnet 3.x с TypeError
   (enum воспринимается как int). Каскадный фолбэк через `revit_compat`
   (шаг 3: BuiltInParameter → имя → LookupParameter) гарантирует работу
   на всех версиях (D-04).

## Контакты с пользователем

- Выбор панели и состава — интерактивные вопросы, согласованы ответы.
- Ревью-гейт — показано полное содержимое + сводка правок; одобрено.
- Push — отложен; коммит локальный.

## Связи к документации

- [[Конвенция MM LAB скриптов]] — AGENTS.md, правила MM001–MM014 + D-01/D-03/D-04.
- [[Процедура приёмки mm-adopt-script]] — 11 шагов, гейты D-08/D-10/D-17/D-18.
- [[WinForms-паттерн для множественного выбора]] — CheckedListBox + Panel + Button.
- [[Общие параметры GP_01_КодКлассифМССК и GP_01_ИмяКлассифМССК]] — GUID сохранены из IFC_Двери.
- [[revit_compat как абстракция мультиверсионности]] — весь каскад `get_parameter`.

## Доп. замечания

- Иконка исходника оставлена (без переименования в icon.png).
- Имя в layout добавлено после `IFC_Окна` (консистентно с блоком IFC-кнопок).
- Runtime UAT отложена (требует pyRevit Reload и живого Revit 2020/2022/2024).
