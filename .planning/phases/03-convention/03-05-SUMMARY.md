---
phase: 03-convention
plan: 05
subsystem: convention
tags: [agents-md, claude-md, gemini-cli, kilocode, pyrevit, convention, docs]

# Dependency graph
requires:
  - phase: 03-convention (планы 03-01..03-04)
    provides: чекер MM000–MM014 (tools/check_convention.py) + baseline, revit_compat.py, шаблон templates/НоваяКнопка.pushbutton
provides:
  - AGENTS.md — канонический стандарт скриптов MM LAB (русский, 311 строк, 10 разделов)
  - CLAUDE.md как тонкий указатель @AGENTS.md + сохранённый дословно GSD Release Map
  - GEMINI.md (@AGENTS.md, memory import Gemini CLI)
  - .kilocode/rules/00-mmlab.md — указатель для старых версий Kilo Code
affects: [03-06 (mm-команды ссылаются на AGENTS.md §Команды MM LAB), 03-07, все будущие кнопки]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Один канонический AGENTS.md + тонкие per-agent указатели без дублей (D-05, D-07)"
    - "Мост Claude Code / Gemini CLI — строка @AGENTS.md первой непустой строкой контекст-файла"
    - "graphify-правило и Obsidian-поток живут в AGENTS.md; поток «сохрани сессию» свёрнут в /mm-save-session (D-21)"

key-files:
  created:
    - AGENTS.md
    - GEMINI.md
    - .kilocode/rules/00-mmlab.md
  modified:
    - CLAUDE.md

key-decisions:
  - "Канонический бутстрап-блок в AGENTS.md вынесен из отступа нумерованного списка на верхний уровень — текст блока дословно совпадает с шаблоном (копипаста без отступов)"
  - "Запрет легаси-бутстрапа сформулирован без литерального легаси-имени переменной — acceptance criterion требует отсутствия этого имени в файле; триггеры MM014 описаны словами"
  - "Белый список импортов в AGENTS.md включает Microsoft — сверен с HOST_IMPORT_ROOTS чекера (план перечислял сокращённый список)"
  - "Сообщение коммита Task 2 взято дословно из acceptance criteria плана"

patterns-established:
  - "Per-agent файлы содержат только указатель + агент-специфику: CLAUDE.md = @AGENTS.md + GSD Release Map + ссылка на .claude/commands/"
  - "Таблица MM-кодов в стандарте сверяется со словарём RULES чекера (severity до буквы)"

requirements-completed: [CONV-STD]

# Coverage metadata — per-deliverable verification
coverage:
  - id: D1
    description: "AGENTS.md — канонический стандарт: все обязательные разделы, таблица MM000–MM014 с severity из RULES, API revit_compat, канонический бутстрап дословно, шаблон сессионного коммита нового формата, 7 mm-слагов, без userEmail/currentDate"
    requirement: CONV-STD
    verification:
      - kind: other
        ref: "py -3 marker-check (14 обязательных маркеров из <verify> плана) — exit 0"
        status: pass
      - kind: other
        ref: "acceptance-гейт: 311 строк ≥ 180; severity MM000–MM014 == RULES; бутстрап-блок verbatim; EXTENSION_ROOT/userEmail/currentDate отсутствуют; «## Сессия» + «- Агент/модель:»; 7 слагов"
        status: pass
    human_judgment: false
  - id: D2
    description: "Тонкие указатели: CLAUDE.md (строка 1 @AGENTS.md, GSD-блок дословно, без graphify/Obsidian), GEMINI.md (@AGENTS.md), .kilocode/rules/00-mmlab.md; .claude/CLAUDE.md не изменён"
    requirement: CONV-STD
    verification:
      - kind: other
        ref: "py -3 verify Task 2 (первые строки, gsd_release_sync.py, отсутствие graphify query/MMLabs_OBSIDIAN) — exit 0"
        status: pass
      - kind: other
        ref: "difflib-сверка GSD-блока с git show HEAD:CLAUDE.md — VERBATIM (10 строк); git status .claude/CLAUDE.md — пуст"
        status: pass
    human_judgment: false

# Metrics
duration: 9min
completed: 2026-07-24
status: complete
---

# Phase 03 Plan 05: AGENTS.md + тонкие указатели агентов Summary

**AGENTS.md (311 строк, RU) — единственный полный текст конвенции MM LAB, сверенный с чекером и revit_compat; CLAUDE.md/GEMINI.md/.kilocode сведены к тонким указателям @AGENTS.md**

## Performance

- **Duration:** 9 min
- **Started:** 2026-07-24T11:58:50Z
- **Completed:** 2026-07-24T12:08:20Z
- **Tasks:** 2
- **Files modified:** 4

## Accomplishments

- AGENTS.md: 10 обязательных разделов — архитектура репозитория (first-party vs vendored lib), 18 нумерованных правил кода с MM-кодами, структура кнопки и layout-правила, таблица ломающих изменений Revit 2020/2022/2024 + 13 функций revit_compat, команды чекера с таблицей MM000–MM014 и семантикой baseline, каталог 7 mm-команд, Git-регламент с шаблоном сессионного коммита (новый формат «## Сессия»), правило graphify и Obsidian-хранилище (D-21)
- CLAUDE.md переписан: первая строка `@AGENTS.md`, GSD Release Map сохранён дословно (сверено difflib), graphify/Obsidian удалены (перенесены в AGENTS.md), добавлена ссылка на /mm-* адаптеры в .claude/commands/
- GEMINI.md и .kilocode/rules/00-mmlab.md созданы: три агента получают один стандарт (Claude — @import, Gemini — memory import, Kilo — нативно + файл правил для старых версий)
- Все файлы UTF-8 без BOM; symlink не использованы (D-05, Windows)

## Task Commits

Each task was committed atomically:

1. **Task 1: Написать AGENTS.md — канонический стандарт** - `f2558eb` (docs)
2. **Task 2: Тонкие указатели — CLAUDE.md, GEMINI.md, .kilocode/rules/00-mmlab.md** - `2ab2a33` (docs)

## Files Created/Modified

- `AGENTS.md` - канонический стандарт скриптов MM LAB (единственный источник правды)
- `CLAUDE.md` - тонкий указатель @AGENTS.md + Claude-специфика (GSD Release Map, /mm-* из .claude/commands/)
- `GEMINI.md` - тонкий указатель @AGENTS.md для Gemini CLI
- `.kilocode/rules/00-mmlab.md` - указатель для старых версий Kilo Code

## Decisions Made

- Канонический бутстрап-блок MM014 размещён в AGENTS.md без отступа списка: первоначальная версия внутри пункта 17 получала 4-пробельный отступ и переставала быть дословной копией шаблона; блок вынесен на верхний уровень (нумерация 18 задана явно)
- Запрет легаси-бутстрапа описан без литерального легаси-имени переменной (acceptance criterion требует его отсутствия в файле): триггеры MM014 переданы формулировкой «легаси-имя переменной „корня расширения“, подъём на 4 уровня .., иной sys.path-вызов»
- Белый список импортов дополнен `Microsoft` — по HOST_IMPORT_ROOTS реализации чекера (таблица стандарта обязана совпадать с реализацией до буквы)
- Сообщение коммита Task 2 — дословно из acceptance criteria плана: «docs(03-05): AGENTS.md + тонкие указатели агентов»

## Deviations from Plan

None - plan executed exactly as written.

## Threat Model Dispositions

- T-03-13 (Tampering, Git-регламент): mitigated — AGENTS.md §Git-регламент явно запрещает `git add .`/`-A`, требует пофайловый стейджинг и push только после подтверждения (D-17, D-18)
- T-03-14 (Repudiation, потеря правил): mitigated — GSD-блок сверён difflib с git-версией (VERBATIM), перенос graphify/Obsidian в AGENTS.md подтверждён гейтом до удаления из CLAUDE.md
- T-03-15 (Info disclosure): mitigated — подстроки userEmail/currentDate в AGENTS.md отсутствуют (проверено гейтом)

## Issues Encountered

None

## Notes

- `requirements mark-complete CONV-STD` пропущен: REQUIREMENTS.md генерируется из «Карты релизов» и содержит только задачи плагинов (T-xxx); условная метка CONV-STD в нём не заведена (зафиксировано в 03-RESEARCH §Phase Requirements)
- Форвард-ссылки AGENTS.md на `agents/commands/mm-*.md` — намеренные: каталог команд создаёт план 03-06 (допущено планом)

## Next Phase Readiness

- Стандарт готов; план 03-06 создаёт agents/commands/ + адаптеры .claude/.gemini/.kilo, на которые AGENTS.md уже ссылается фиксированными слагами
- Претензий/блокеров нет

---
*Phase: 03-convention*
*Completed: 2026-07-24*

## Self-Check: PASSED

- Файлы: AGENTS.md, GEMINI.md, .kilocode/rules/00-mmlab.md, CLAUDE.md, 03-05-SUMMARY.md — FOUND
- Коммиты: f2558eb, 2ab2a33 — FOUND
- Automated verify обеих задач — exit 0
