---
tags: [sessions, log, debugging, architecture, revit-api, pythonnet]
date: 2026-07-20
---

# Исправлен краш скрипта «Мокрые зоны» из-за маршалирования IList в pythonnet

## Информация о сессии
- Модель: claude-opus-4-8
- Дата: 2026-07-20
- Изменено файлов: 1

## Изменённые файлы
- `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Мокрые зоны.pushbutton/script.py`

## Симптом
Скрипт «Мокрые зоны» падал только на части пользователей с ошибкой:
```
AttributeError: 'list' object has no attribute 'Count'
File "<string>", line 674, in <module>
File "<string>", line 598, in main
File "<string>", line 142, in get_room_boundary_loop
```
На других пользователях той же модели Revit (R2024) работал корректно.

## Что сделано
Найдена и исправлена коренная причина.

### Диагностика
- Функция `get_room_boundary_loop` вызывает `room.GetBoundarySegments(opts)` (Revit API, строка 138)
- Возвращаемый тип — `.NET IList<IList<BoundarySegment>>`
- На строках 142 и 147 код вызывает `.Count` (свойство .NET): `boundary.Count == 0` и `segments.Count == 0`
- Проблема в маршалировании pythonnet: разные версии pythonnet (2.5.x vs 3.x) маршалят общие `.NET IList<T>` по-разному
  - pythonnet 2.5.x (старый pyRevit): возвращает `.NET wrapper` с `.Count` ✓
  - pythonnet 3.x (новый pyRevit, Revit 2024+): возвращает нативный Python `list` без `.Count` ✗
- Пользователи с новым pyRevit/pythonnet 3.x на R2024 получали `AttributeError: 'list' object has no attribute 'Count'`
- Пользователи со старым pyRevit/pythonnet 2.x или IronPython не жаловались

### Решение
Заменить `.Count` на встроенную функцию Python `len()`:
- `len()` работает **идентично** на:
  - Python `list` (всегда есть `__len__`)
  - .NET `IList<T>` под pythonnet 2.x (wrapper реализует `__len__`)
  - .NET `IList<T>` под pythonnet 3.x (маршалится как Python `list`)
  - IronPython объектах (реализует `__len__`)

Исправлены строки 142 и 147 функции `get_room_boundary_loop`:
```python
# было:
if boundary is None or boundary.Count == 0:
if segments is None or segments.Count == 0:

# стало:
if boundary is None or len(boundary) == 0:
if segments is None or len(segments) == 0:
```

Оставлены без изменений строки 428, 501, 502 (`.Count` на WinForms коллекциях `dgv.Rows`, `cmb_phase.Items`) — те всегда геннуинные .NET объекты, маршалирование их не меняется между версиями.

## Ключевое решение
- `len()` — кроссплатформенное решение, работающее на всех версиях pythonnet, IronPython и будущих версиях pyRevit
- Это не замена на какой-то string-based fallback (как `LookupParameter`) — просто синтаксический сдвиг, безопасный для всех типов возврата Revit API

## Контекст
- Пользователь сообщил: «Проблема возникает на R2024» — это сужение указывает на версию-зависимый баг pythonnet маршалирования, что и подтвердилось
- Живая проверка в Revit на проблемных ПК не выполнялась (headless-среда разработки) — требуется подтверждение пользователем

## Следующие шаги
- [[Мультиверсионная совместимость Revit API через адаптер версии]] — документировать парад версий pythonnet и их маршалирование
- [[Отсутствие тестов повышает риск регрессий при правках]] — нужны автотесты для Revit API обвязки
