@AGENTS.md

<!-- Ниже — только специфичное для Claude Code; стандарт MM LAB — в AGENTS.md. -->

Команды /mm-* доступны как slash-команды из `.claude/commands/` (канонические
процедуры — `agents/commands/`, каталог — AGENTS.md §Команды MM LAB).

## GSD Release Map (Регламент 8.2.1 / 11.3)
Карта релизов синхронизируется с артефактами GSD скриптом `RELEASE_MAP/gsd_release_sync.py`.
Источник правды — CSV-экспорт листа «Скрипты_Карта релизов»
(`RELEASE_MAP/Сводный Реестр Плагинов - Скрипты_Карта релизов.csv`).
### Команда «Синхронизируй gsd» = `/mm-releasemap-download`
Когда пользователь говорит "Синхронизируй gsd" или "скачай карту релизов", выполни
процедуру `agents/commands/mm-releasemap-download.md` (кратко — `check` + `sync-docs`
+ разбор пропусков):
1. `python3 RELEASE_MAP/gsd_release_sync.py check` — валидация CSV «Карта релизов».
2. `python3 RELEASE_MAP/gsd_release_sync.py sync-docs` — генерация `.planning/release-map.json`, `ROADMAP.md`, `REQUIREMENTS.md`, `PROJECT.md` (с сохранением прогресса).
Эквивалент: `python3 RELEASE_MAP/gsd_release_sync.py sync`. После выполнения покажи, что изменилось.
### Правила синхронизации
Задание НЕ синхронизируется, если статус «Без статуса»/пустой либо пустое
«Название плагина»; `MVP = TRUE` → приоритет `MVP`, иначе `Обычный`. Пропуски
перечисляются в отчёте и в секции «Не синхронизировано» в `ROADMAP.md`.
Генерируемый текст `.planning/*.md` живёт между маркерами `RELEASE-MAP:BEGIN/END`;
всё вне маркеров (например, фазы GSD `## Phase N`) регенерацию переживает.
### Статусы заданий
Жизненный цикл: Не начато → В работе → Готово → Релиз. Менять статус:
`python3 RELEASE_MAP/gsd_release_sync.py status <ID> "<статус>"`.
CSV вручную не редактировать: состав заданий правится в Google-таблице.
