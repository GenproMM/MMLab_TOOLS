## graphify

This project has a knowledge graph at graphify-out/ with god nodes, community structure, and cross-file relationships.

Rules:
- For codebase questions, first run `graphify query "<question>"` when graphify-out/graph.json exists. Use `graphify path "<A>" "<B>"` for relationships and `graphify explain "<concept>"` for focused concepts. These return a scoped subgraph, usually much smaller than GRAPH_REPORT.md or raw grep output.
- If graphify-out/wiki/index.md exists, use it for broad navigation instead of raw source browsing.
- Read graphify-out/GRAPH_REPORT.md only for broad architecture review or when query/path/explain do not surface enough context.
- After modifying code, run `graphify update .` to keep the graph current (AST-only, no API cost).

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

## Obsidian Knowledge Vault
Хранилище знаний: ./MMLabs_OBSIDIAN
### При старте сессии
Прочитай 00-home/index.md и текущие приоритеты.md.
Если задача касается модуля — прочитай заметку из knowledge/.
### При завершении (пользователь: "сохрани сессию")
1. Создай заметку в sessions/ с датой
2. Обнови текущие приоритеты.md
3. Если решение — создай в knowledge/decisions/
4. Если баг — создай в knowledge/debugging/
5. Обнови index.md если новые заметки
6. Автоматически сделай git commit всех изменений сессии с сообщением в формате: "session: YYYY-MM-DD summary"