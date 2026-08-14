# 2026-08-14: Команда /mm-releasemap-download и переименование MM LAB.extension

**Сессия:** Claude Haiku 4.5  
**Результат:** 4 новых команды, парсер реальной Карты релизов, 369 правок пути расширения, тесты зелёные.

## Основная работа

### 1. Команда `/mm-releasemap-download` (новая)
- **Канонический файл:** `agents/commands/mm-releasemap-download.md` (8.2.1 Регламента)
- **Адаптеры:** `.claude/`, `.gemini/`, `.kilo/commands/` (4 файла всего)
- **Регистрация:** AGENTS.md (таблица команд), CLAUDE.md (фраза-триггер)
- **Интеграция:** естественный язык («Синхронизируй gsd» / «скачай карту релизов») → `/mm-releasemap-download`

### 2. Адаптация парсера `RELEASE_MAP/gsd_release_sync.py`
**Проблема:** скрипт читал эталонный CSV, реальный лист другой структуры  
**Решение:** переписан парсер под лист «Скрипты_Карта релизов» (Google Sheets):
- Наследование версии релиза от строк-заголовков (forward-fill)
- Колонки: `Название плагина`, `MVP` (TRUE/FALSE → приоритет), `Вес`, `Группа задач`, `Автор`, `Комментарий`
- Правила отсева: статус «Без статуса» и пустое «Название плагина» → не синхронизируются
- Результат: 8 заданий в 3 релизах (v250407, v250516, v251205), 4 пропущено с указанием причины

**Генерация в маркерах:** `.planning/*.md` — генерируемый текст между `RELEASE-MAP:BEGIN/END`, ручной текст вне маркеров переживает регенерацию (GSD-фазы в ROADMAP.md не затираются)

**Тесты:** новый файл `tools/tests/test_release_map_sync.py` — 18 тестов (правила отсева, наследование, приоритет, ошибки, сохранение ручного текста), все зелёные

### 3. Переименование MM LAB.extension → MM_LAB.extension
**Причина:** [коммит 8de634e](https://github.com/) переименовал каталог в git, пути в инструментах остались старыми  
**Размер:** 369 вхождений в 62 файлах (исключены только путь-токен, бренд «MM LAB» в заголовках не трогал)

**Затронутые зоны:**
- `tools/check_convention.py` — `EXTENSION_DIR_NAME`, 6 упоминаний в докстрингах
- `tools/convention_baseline.json` — 15 ключей секции `units`
- `tools/tests/test_check_convention.py` — 10 путей фикстур
- `tools/tests/fixtures/repo_ok` и `repo_bad` — `git mv` каталогов
- `tools/tests/test_revit_compat_contract.py` — пути в `COMPAT_PATH`
- Весь код: AGENTS.md (архитектура, правила 4/6/17), команды, скрипты, шаблон, Регламент
- Артефакты: `.planning/codebase/` (документация архитектуры) и `debug/`

**Результат гейта:**
- До правок: `Проверено: 0` (каталог не найден)
- После: **`Проверено: 19, ошибок: 10, предупреждений: 8`** (находит расширение)
- Baseline корректно гасит 15 легаси-кнопок

**Побочный вывод:** две кнопки `СНиП` и `СНиП_ФОП25` (коммит ccde781) не прошли приёмку — нет `#! python3`, `coding`, `bundle.yaml`, `README.md`, есть wildcard-импорты и `pyrevit.forms`. Решение за пользователем: `/mm-adopt-script` или `pending_adoption`.

## Тесты (74 всего)
✅ `tools/tests/test_mm_commands_catalog.py` — 7 тестов (обновлён слаг `mm-releasemap-download`)  
✅ `tools/tests/test_revit_compat_contract.py` — 6 тестов (пути заживают)  
✅ `tools/tests/test_check_convention.py` — 43 теста (фикстуры работают)  
✅ `tools/tests/test_release_map_sync.py` — 18 тестов (новый)  
✅ Остальное — общие тесты

## Ключевые решения

1. **Маркеры в `.planning/*.md`** — регенерируемый текст между маркерами, ручной снаружи. Первая миграция → бэкап в `.planning/_legacy/` (добавлен в `.gitignore`)
2. **Правила синхронизации CSV** — явное решение пользователя (2026-08-14): «Без статуса» не синхронизируется, `MVP=TRUE` → MVP, пустое имя плагина не синхронизируется
3. **Разделение файлов сессии** — спор про `.planning/` и `MMLabs_OBSIDIAN/` решён выбором пользователя (всё в один коммит)

## Следующие шаги

- Две непринятые кнопки `СНиП*` → `/mm-adopt-script` или запись в `pending_adoption` (статус гейта будет честным: 10 ошибок — это не потеря, это рабочий статус)
- Возможное расширение: флаг `--all` в `gsd_release_sync.py` для связки выполненных задач с сессиями (требует реверс-инженирии Obsidian-путей)

## Файлы сессии
- agents/commands/mm-releasemap-download.md (новая)
- .claude/commands/mm-releasemap-download.md (новая)
- .gemini/commands/mm-releasemap-download.toml (новая)
- .kilo/commands/mm-releasemap-download.md (новая)
- RELEASE_MAP/gsd_release_sync.py (переписан)
- RELEASE_MAP/README.md (переписан)
- tools/tests/test_release_map_sync.py (новая)
- tools/ (пути: check_convention.py, baseline, fixtures, test_revit_compat_contract.py)
- 62 файла с заменой MM LAB.extension
- .planning/, MMLabs_OBSIDIAN/ (побочно обновлены)
