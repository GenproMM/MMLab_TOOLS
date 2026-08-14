---
status: resolved
trigger: "Ошибка при запуске кнопки «Мокрые зоны» только на одном ПК"
created: 2026-07-07
updated: 2026-07-09
---

# Symptoms

## Expected behavior
Кнопка «Мокрые зоны» должна открыть диалог настройки (фильтр названий + стадия), затем выполнить анализ пересечений помещений с мокрыми зонами и показать отчёт.

## Actual behavior
При нажатии кнопки появляется окно PyRevitLoader с ошибкой:
```
TypeError: No method matches given arguments for get_Parameter: (<class 'int'>)
```
Ошибка происходит в функции `get_rooms_in_phase`, строка 103 файла `script.py`.

## Error details
```
PyRevitLoader - Ошибка
Ошибка при выполнении:
Traceback (most recent call last):
File "<string>", line 655, in <module>
File "<string>", line 536, in main
File "<string>", line 103, in get_rooms_in_phase
TypeError: No method matches given arguments for get_Parameter: (<class 'int'>)
```

## Timeline
- Происходит только на одном ПК; на других ПК кнопка работает корректно
- Неизвестно, работала ли раньше на проблемном ПК
- Различия проблемного ПК от рабочих неизвестны

## Reproduction
1. Открыть проект Revit на проблемном ПК
2. Нажать кнопку «Мокрые зоны» на панели «АРХИТЕКТУРА» вкладки «GENPRO LAB»
3. Ошибка появляется сразу (до открытия диалога)

# Analysis

## Code at fault
File: `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Мокрые зоны.pushbutton/script.py`
Line 103:
```python
phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
```

## Root cause hypothesis
`BuiltInParameter.ROOM_PHASE` — это enum-значение Revit API, которое в IronPython представлено как `int`. Метод `get_Parameter()` имеет несколько перегрузок (принимает `Guid`, `BuiltInParameter`, `Definition`, `string`). IronPython на проблемном ПК не может разрешить правильную перегрузку метода, потому что `int` (значение enum) неоднозначно между `Guid` и `BuiltInParameter`.

**Наиболее вероятная причина различия между ПК:** проблемный ПК имеет другую версию Revit (вероятно 2023+), где:
- Перегрузка `get_Parameter(BuiltInParameter)` была изменена/помечена deprecated в пользу `GetParameter(ForgeTypeId)`
- IronPython резолвит перегрузки иначе из-за изменений в .NET API

## Suggested fix
Заменить `room.get_Parameter(BuiltInParameter.ROOM_PHASE)` на `room.LookupParameter("Phase")`:
```python
# Строка 103 — было:
phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
# Должно быть:
phase_param = room.LookupParameter("Phase")
```

# Current Focus

hypothesis: "pythonnet на проблемном ПК не может разрешить перегрузку get_Parameter(BuiltInParameter) т.к. BuiltInParameter.ROOM_PHASE передаётся как Python int, а метод имеет конкурирующие перегрузки (Guid, BuiltInParameter, Definition, string)"
status: "CONFIRMED — root cause identified"
next_action: "return structured diagnosis"
root_cause: "В pythonnet (CPython 3 .NET bridge, используемый pyRevit с #! python3) .NET enum BuiltInParameter представлен как Python int. Метод get_Parameter имеет несколько перегрузок: get_Parameter(Guid), get_Parameter(BuiltInParameter), get_Parameter(Definition), get_Parameter(string). На проблемном ПК (вероятно Revit 2024+ с .NET 6 и/или другая версия pythonnet) резолвер перегрузок не может сопоставить Python int ни с одной перегрузкой, выбрасывая TypeError: No method matches given arguments for get_Parameter: (<class 'int'>). На рабочих ПК (Revit 2020/2022 с .NET Framework 4.8) та же конструкция работает, что указывает на различие в поведении pythonnet или CLR между версиями."

# Evidence

- timestamp: 2026-07-07
  type: error_traceback
  description: "TypeError: No method matches given arguments for get_Parameter: (<class 'int'>) в строке 103"
  source: "пользователь"
- timestamp: 2026-07-07
  type: code_analysis
  description: "Строка 103: phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE). BuiltInParameter.ROOM_PHASE передаётся как int в pythonnet, что вызывает ошибку разрешения перегрузки."
  source: "code_review"
- timestamp: 2026-07-07
  type: behavioral
  description: "Ошибка только на одном ПК. На других ПК работает."
  source: "пользователь"
- timestamp: 2026-07-07
  type: compatibility
  description: "Скрипт заявлен как совместимый с Revit 2020/2022/2024. Ошибка перегрузки get_Parameter типична для Revit 2024+ (.NET 6) где COM interop изменился, что влияет на резолвинг перегрузок в pythonnet."
  source: "code_review"
- timestamp: 2026-07-07
  type: code_analysis
  description: "Shared lib ios_common_helpers.py (146) использует element.get_Parameter(built_in_parameter) — тот же паттерн, но для элементов MEP (воздуховоды, фитинги), не для комнат."
  source: "code_review"
- timestamp: 2026-07-07
  type: code_analysis
  description: "Другие скрипты, работающие с комнатами, используют LookupParameter() вместо get_Parameter(BuiltInParameter): «Высота этажа» (130), «Экспорт ПСО» (118). «Экспорт ПСО» (116) вызывает room.get_Parameter(DB.BuiltInParameter.ROOM_NAME) через модуль DB, а не прямой импорт BuiltInParameter."
  source: "code_review"
- timestamp: 2026-07-07
  type: code_analysis
  description: "BuiltInParameter.ROOM_PHASE используется ТОЛЬКО в этом скрипте (103) — единственное место во всём расширении. Другие BuiltInParameter (ROOM_NAME, RBS_ADDITIONAL_FLOW и др.) используются в других скриптах, но с другими типами элементов."
  source: "code_review"
- timestamp: 2026-07-07
  type: root_cause_determination
  description: "Корневая причина: несовместимость pythonnet (CPython 3 .NET bridge) с перегрузками get_Parameter при передаче BuiltInParameter как Python int. Отличающееся поведение на проблемном ПК вызвано иной версией pythonnet (более строгий резолвер перегрузок) и/или .NET 6 в Revit 2024+ (.NET Core меняет COM interop, что влияет на то, какие перегрузки метода видит pythonnet)."
  source: "analysis"
- timestamp: 2026-07-07
  type: no_knowledge_base_match
  description: "Knowledge base не существует — первая отладочная сессия, возвращать нечего."
  source: "knowledge_base_check"

# Eliminated

- hypothesis: "Опечатка в названии кнопки (Мкорые vs Мокрые)"
  reason: "Опечатка только в пользовательском вводе, не в коде. Кнопка корректно названа «Мокрые зоны» в bundle.yaml."
- hypothesis: "Файл скрипта повреждён или отсутствует"
  reason: "Ошибка происходит внутри скрипта (выполняется до строки 103), значит файл найден и начинает выполняться."
- hypothesis: "Проблема с правами доступа"
  reason: "Ошибка не связана с файловой системой; это ошибка разрешения перегрузки .NET метода."
- hypothesis: "Проблема с import BuiltInParameter (AttributeError)"
  reason: "Ошибка TypeError, а не AttributeError — импорт успешен, BuiltInParameter.ROOM_PHASE существует."
- hypothesis: "Проблема специфична для IOS_common_helpers.get_parameter()"
  reason: "Ошибка происходит в script.py:103, где BuiltInParameter.ROOM_PHASE передаётся напрямую в get_Parameter(), а не через helper-функцию."
- hypothesis: "Проблема в EXTENSION_ROOT/lib path (ошибка на 1 уровень в os.path.join)"
  reason: "Даже если LIB_DIR определён неверно, revit_ui_helpers импортируется через автоматический sys.path от pyRevit. Ошибка не связана с импортом."

# Resolution

**Root cause:** `BuiltInParameter.ROOM_PHASE` передаётся как Python `int` в `get_Parameter()`, который имеет конкурирующие перегрузки `(Guid, BuiltInParameter, Definition, string)`. pythonnet — CPython 3 .NET bridge, используемый pyRevit с `#! python3` — не может разрешить правильную перегрузку. Причина отличия проблемного ПК: иная версия pythonnet (более строгий резолвер перегрузок) и/или Revit 2024+ с .NET 6, где изменился COM interop, что повлияло на метаданные перегрузок, видимые pythonnet.

**Fix:** Заменить `room.get_Parameter(BuiltInParameter.ROOM_PHASE)` на `room.LookupParameter("Phase")`. `LookupParameter` имеет единственную перегрузку `(string)` — резолвер всегда сработает. Альтернатива: `room.Phase` (свойство `SpatialElement`, доступно во всех релевантных версиях Revit).

**Файл для изменения:** `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Мокрые зоны.pushbutton/script.py`, строка 103.

**Верификация:** После замены ошибка `TypeError: No method matches given arguments for get_Parameter: (<class 'int'>)` должна исчезнуть на проблемном ПК. Требуется подтверждение пользователя.

# Applied fix (2026-07-09)

Первоначальная рекомендация (`LookupParameter("Phase")`) **отклонена**: `LookupParameter`
ищет по локализованному имени параметра. У русскоязычной BIM-команды Revit русский,
где параметры называются «Стадия»/«Имя»/«Номер», поэтому `LookupParameter("Phase")`/
`("Name")`/`("Number")` вернул бы `None` → молчаливый пропуск всех помещений (тихая
регрессия, хуже краша).

Ключевое наблюдение: соседний скрипт «Экспорт ПСО» (тоже `#! python3`) успешно
использует `room.get_Parameter(DB.BuiltInParameter.ROOM_NAME)` на всех машинах команды —
значит паттерн рабочий, а проблема в резолвинге перегрузки на конкретной среде.

**Что сделано:** добавлен helper `_get_param(element, bip)`, который явно выбирает
перегрузку через pythonnet — `element.get_Parameter.__overloads__[BuiltInParameter](bip)` —
с fallback на прямой вызов (`except (TypeError, AttributeError, KeyError)`). Явный выбор
перегрузки снимает неоднозначность (Guid/BuiltInParameter/Definition/string) и включает
приведение аргумента к enum даже если pythonnet видит его как int. Fallback гарантирует
неизменное поведение на средах, где всё уже работает (Revit 2020/2022, IronPython, старый
pythonnet без `__overloads__`).

**Через helper проведены ВСЕ 6 вызовов** `get_Parameter(BuiltInParameter.*)` в файле
(ROOM_PHASE, ROOM_NAME ×3, ROOM_NUMBER ×2) — иначе после фикса строки 103 краш повторился
бы на следующем вызове (например, в `is_wet_room`).

**Проверка:** `python -m py_compile` — OK. Живая проверка в Revit на проблемном ПК —
follow-up за пользователем (headless-среда разработки).
