@AGENTS.md

<!-- Ниже — только специфичное для Claude Code; стандарт MM LAB — в AGENTS.md. -->

Команды /mm-* доступны как slash-команды из `.claude/commands/` (канонические
процедуры — `agents/commands/`, каталог — AGENTS.md §Команды MM LAB).

## GSD Release Map (Регламент 8.2.1 / 11.3)
Карта релизов синхронизируется с артефактами GSD скриптом `RELEASE_MAP/gsd_release_sync.py`.
### Команда «Синхронизируй gsd»
Когда пользователь говорит "Синхронизируй gsd", выполни:
1. `python3 RELEASE_MAP/gsd_release_sync.py check` — валидация CSV «Карта релизов».
2. `python3 RELEASE_MAP/gsd_release_sync.py sync-docs` — генерация `.planning/release-map.json`, `ROADMAP.md`, `REQUIREMENTS.md`, `PROJECT.md` (с сохранением прогресса).
Эквивалент: `python3 RELEASE_MAP/gsd_release_sync.py sync`. После выполнения покажи, что изменилось.
### Статусы заданий
Жизненный цикл: Не начато → В работе → Готово → Релиз. Менять статус:
`python3 RELEASE_MAP/gsd_release_sync.py status <ID> "<статус>"`.
