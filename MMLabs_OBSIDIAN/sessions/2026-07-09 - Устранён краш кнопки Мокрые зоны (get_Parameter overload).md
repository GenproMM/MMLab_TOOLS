---
tags: [sessions, log, debugging, architecture, revit-api]
date: 2026-07-09
---

# Устранён краш кнопки «Мокрые зоны» на резолвинге перегрузки get_Parameter

## Информация о сессии
- Модель: claude-opus-4-8
- Дата: 2026-07-09
- Изменено файлов: 5

## Изменённые файлы
- `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Мокрые зоны.pushbutton/script.py`
- `.planning/debug/wet-zones-get-parameter.md`
- `MMLabs_OBSIDIAN/knowledge/debugging/Замена get_Parameter на LookupParameter небезопасна для локализованного Revit.md`
- `MMLabs_OBSIDIAN/00-home/index.md`
- `MMLabs_OBSIDIAN/sessions/2026-07-09 - Устранён краш кнопки Мокрые зоны (get_Parameter overload).md`

## Симптом
- Кнопка «Мокрые зоны» (панель АРХИТЕКТУРА) падала на одном ПК с `TypeError: No method matches given arguments for get_Parameter: (<class 'int'>)`, строка 103 в `get_rooms_in_phase`. На остальных машинах работала.

## Что сделано
- Причина (уже была диагностирована в `.planning/debug/wet-zones-get-parameter.md`): под `#! python3` (pythonnet) на Revit 2024+ enum `BuiltInParameter` воспринимается как обычный Python `int`, и резолвер не может выбрать перегрузку `get_Parameter` (конкурируют `Guid / BuiltInParameter / Definition / string`).
- Добавлен helper `_get_param(element, bip)` с явным выбором перегрузки через `element.get_Parameter.__overloads__[BuiltInParameter](bip)` и fallback на прямой вызов — на уже рабочих средах (Revit 2020/2022, IronPython, старый pythonnet) поведение не меняется.
- Через helper проведены **все 6** вызовов `get_Parameter(BuiltInParameter.*)` в файле (`ROOM_PHASE`, `ROOM_NAME` ×3, `ROOM_NUMBER` ×2), а не только строка 103 — иначе краш повторился бы на следующем вызове в `is_wet_room`.
- Debug-заметка переведена в статус `resolved` с описанием применённого фикса.

## Ключевое решение
- Первоначальная рекомендация debug-заметки (`LookupParameter("Phase")`) **отклонена**: `LookupParameter` ищет по локализованному имени параметра, у русскоязычной команды Revit русский («Стадия»/«Имя»/«Номер»), поэтому английские строки вернули бы `None` → молчаливый пропуск всех помещений (тихая регрессия хуже краша).
- Это уточняет прежнюю заметку [[Несуществующие имена BuiltInParameter ломают кнопки ИОС при вызове]], где `LookupParameter(string)` предлагался как универсальная замена.

## Контекст
- Живая проверка в Revit на проблемном ПК не выполнялась (headless-среда разработки) — оставлен follow-up на подтверждение пользователем.
- `python -m py_compile` — OK.

## Следующие шаги
- [[Замена get_Parameter на LookupParameter небезопасна для локализованного Revit]]
- [[Мультиверсионная совместимость Revit API через адаптер версии]]
- [[Отсутствие тестов повышает риск регрессий при правках]]
