---
tags: [knowledge, debugging, revit-api, architecture, i18n]
date: 2026-07-09
---

# Замена get_Parameter на LookupParameter небезопасна для локализованного Revit

Проблема:
- `get_Parameter(BuiltInParameter.X)` под `#! python3` (pythonnet) на Revit 2024+ падает с `No method matches given arguments for get_Parameter: (<class 'int'>)` — enum приходит как Python `int`, и перегрузки `(Guid / BuiltInParameter / Definition / string)` неоднозначны.
- Наивная замена на `LookupParameter("Phase")` / `("Name")` / `("Number")` **не является безопасной универсальной заменой**: `LookupParameter` резолвит параметр по локализованному имени (`Definition.Name`). На русском Revit это «Стадия» / «Имя» / «Номер», поэтому английские строки вернут `None` → молчаливый пропуск данных вместо краша (тихая регрессия хуже явной ошибки).

Базовое решение:
- Для BuiltInParameter'ов уходить не в строковый `LookupParameter`, а в **явный выбор перегрузки**: `element.get_Parameter.__overloads__[BuiltInParameter](bip)` с fallback на прямой вызов. Снимает неоднозначность и включает приведение аргумента к enum, оставаясь языконезависимым.
- `LookupParameter(string)` применять только для параметров, чьё имя стабильно и не локализуется (кастомные общие параметры, напр. `GP_23_*`, `Высота этажа`).
- Где есть чистое свойство API — предпочитать его (`Room.Number`, `Element.Name`) вместо чтения BuiltInParameter.

Где встречалось:
- Кнопка «Мокрые зоны»: 6 вызовов `get_Parameter(BuiltInParameter.*)` (ROOM_PHASE/ROOM_NAME/ROOM_NUMBER) переведены на helper `_get_param`.
- «Экспорт ПСО» (тоже `#! python3`) успешно использует `room.get_Parameter(DB.BuiltInParameter.ROOM_NAME)` — паттерн рабочий, проблема лишь в резолвинге перегрузки на конкретной среде.

Связанные утверждения:
- [[Несуществующие имена BuiltInParameter ломают кнопки ИОС при вызове]]
- [[Мультиверсионная совместимость Revit API через адаптер версии]]
- [[Продукт ориентирован на BIM-команды в русскоязычном контуре]]
- [[Тихие except блоки скрывают критические ошибки в ИОС-скриптах]]
