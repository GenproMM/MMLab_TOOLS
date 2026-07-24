---
phase: 03-convention
plan: 06
subsystem: agent-commands
tags: [pyrevit, mm-commands, procedures, git, obsidian, convention, adoption]

# Dependency graph
requires:
  - phase: 03-convention (03-03)
    provides: "CLI-контракт чекера check_convention.py (--json/--strict/--baseline/--all/--root, exit 0/1/2)"
  - phase: 03-convention (03-04)
    provides: "templates/НоваяКнопка.pushbutton (канонический бутстрап) + tools/convention_baseline.json"
  - phase: 03-convention (03-05)
    provides: "AGENTS.md — стандарт, §Команды MM LAB (каталог 7 команд), §Git-регламент (шаблон коммита)"
provides:
  - "agents/commands/ — 7 канонических процедур mm-команд (D-19, D-20), единственный полный текст каждой команды"
  - "mm-adopt-script: полный поток приёмки с гейтами D-08 (ревью до регистрации), D-10 (панель с авто-подсказкой), D-09 (quick task оба пути)"
  - "mm-save-session: шаблон коммита дословно из AGENTS.md, пофайловый стейджинг (D-17), push с подтверждением (D-18)"
  - "mm-update-repo: только fetch + pull --ff-only при чистом дереве; деструктивные команды запрещены текстом"
  - "mm-doctor: read-only диагностика в 6 групп проверок"
affects: [03-07 (адаптеры .claude/.gemini/.kilo ссылаются на эти файлы), mm-команды, приёмка скриптов]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Канонический формат процедуры: # /mm-<слаг> + ## Аргументы / ## Процедура / ## Гейты и запреты / ## Финал"
    - "Анти-дрейф: полный текст команды только в agents/commands/; адаптеры агентов — тонкие указатели (план 03-07)"

key-files:
  created:
    - agents/commands/mm-adopt-script.md
    - agents/commands/mm-check.md
    - agents/commands/mm-new-button.md
    - agents/commands/mm-new-compat.md
    - agents/commands/mm-save-session.md
    - agents/commands/mm-update-repo.md
    - agents/commands/mm-doctor.md
  modified: []

key-decisions:
  - "Все 7 процедур несут полный набор секций (включая «Гейты и запреты») — единый interfaces-контракт, даже где acceptance требовал только 3 секции"
  - "Эвристика панели (ИОС/АРХИТЕКТУРА/КООРДИНАЦИЯ) описана в mm-adopt-script и продублирована ссылкой в mm-new-button — D-10 действует одинаково в обеих командах"
  - "mm-new-compat предписывает чистить __pycache__ после py_compile (перенос known pitfall из verify-цепочек фазы)"
  - "CONV-* — условные метки фазы (в REQUIREMENTS.md, сгенерированном из Карты релизов, их нет) — шаг requirements mark-complete пропущен без правки REQUIREMENTS.md"

patterns-established:
  - "Формат mm-процедуры: нумерованные шаги с точными копируемыми командами (py -3 …, git -c core.quotepath=false …)"
  - "Claude-специфичный путь (/gsd-quick) всегда сопровождается ручной альтернативой для Gemini/Kilo"

requirements-completed: [CONV-ADAPT, CONV-REG, CONV-GSD, CONV-CHECK]

# Coverage metadata (#1602)
coverage:
  - id: D1
    description: "mm-adopt-script.md — полный поток приёмки: чекер --json → панель (D-10) → адаптация → --strict → РЕВЬЮ-ГЕЙТ (D-08) → регистрация → quick task (D-09) → пофайловый коммит (D-17)"
    requirement: "CONV-ADAPT"
    verification:
      - kind: other
        ref: "py -3 -c \"…req=['--strict','--json','РЕВЬЮ-ГЕЙТ',…]\" (Task 1 verify: контент-гейт 10 строк)"
        status: pass
    human_judgment: true
    rationale: "Соблюдение гейтов D-08/D-10 (diff показан, без «да» нет регистрации) проверяется только живым прогоном — Manual UAT фазы на IFC_Двери через адаптер плана 03-07"
  - id: D2
    description: "mm-check.md — прогон чекера (путь/--all+--baseline/--strict), трактовка exit 0/1/2, пересказ по MM-кодам, read-only"
    requirement: "CONV-CHECK"
    verification:
      - kind: other
        ref: "py -3 -c Task 2 verify (check_convention.py/--strict/--baseline присутствуют)"
        status: pass
    human_judgment: false
  - id: D3
    description: "mm-new-button.md — скаффолд из templates/НоваяКнопка.pushbutton, запрет кнопки внутри templates/, регистрация в layout, --strict до зелёного"
    verification:
      - kind: other
        ref: "py -3 -c Task 2 verify (templates//НоваяКнопка.pushbutton/layout/Reload/--strict присутствуют)"
        status: pass
    human_judgment: false
  - id: D4
    description: "mm-new-compat.md — SUPPORTED_VERSIONS + ревизия версионных веток + три точки синхронизации (revit_compat.py, AGENTS.md, контрактный тест)"
    verification:
      - kind: other
        ref: "py -3 -c Task 2 verify (SUPPORTED_VERSIONS/test_revit_compat/py_compile/AGENTS.md присутствуют)"
        status: pass
    human_judgment: false
  - id: D5
    description: "mm-save-session.md — файлы только текущей сессии, Obsidian-поток, шаблон коммита дословно из AGENTS.md §Git-регламент, push только после «да»"
    requirement: "CONV-GSD"
    verification:
      - kind: other
        ref: "py -3 verbatim-сверка блока шаблона коммита с AGENTS.md (template verbatim match: True) + Task 3 verify"
        status: pass
    human_judgment: true
    rationale: "Manual UAT фазы: живой прогон — коммит по шаблону создаётся, push без подтверждения не выполняется"
  - id: D6
    description: "mm-update-repo.md — только fetch + pull --ff-only при чистом дереве; reset --hard/clean/checkout --force/rebase запрещены текстом"
    verification:
      - kind: other
        ref: "py -3 -c Task 3 verify (--ff-only/status --porcelain/reset --hard присутствуют)"
        status: pass
    human_judgment: false
  - id: D7
    description: "mm-doctor.md — read-only диагностика: Revit vs SUPPORTED_VERSIONS, vendored lib, чекер+тесты, обязательные файлы, git, итоговая таблица"
    verification:
      - kind: other
        ref: "py -3 -c Task 3 verify (SUPPORTED_VERSIONS/openpyxl/--baseline/AGENTS.md присутствуют) + проверка 6 групп"
        status: pass
    human_judgment: false

# Metrics
duration: 9min
completed: 2026-07-24
status: complete
---

# Phase 03 Plan 06: Канонические процедуры mm-команд Summary

**7 канонических процедур mm-команд в agents/commands/ — единственный полный текст каждой команды с гейтами D-08/D-10/D-17/D-18, зафиксированными текстом; закрывает вторую половину цели фазы (команда приёмки CONV-ADAPT/REG/GSD)**

## Performance

- **Duration:** 9 min
- **Started:** 2026-07-24T12:12:57Z
- **Completed:** 2026-07-24T12:22:00Z
- **Tasks:** 3
- **Files modified:** 7 (все созданы)

## Accomplishments

- `mm-adopt-script.md` (168 строк) — полный поток приёмки стороннего скрипта: статический чекер `--json` → панель с авто-подсказкой по эвристике ИОС/АРХИТЕКТУРА/КООРДИНАЦИЯ (D-10) → адаптация по MM-кодам → `--strict` до exit 0 → блокирующий РЕВЬЮ-ГЕЙТ до регистрации (D-08) → bundle.yaml + снятие записи baseline → quick task обоими путями (`/gsd-quick` и ручные артефакты `.planning/quick/` для Gemini/Kilo, D-09) → пофайловый коммит (D-17)
- `mm-check.md`, `mm-new-button.md`, `mm-new-compat.md` — жизненный цикл: проверка (выбор режима чекера, трактовка exit-кодов), скаффолд из шаблона с запретом кнопки внутри `templates/` и обязательной регистрацией в layout, расширение матрицы Revit с тремя синхронными точками обновления
- `mm-save-session.md` (107 строк), `mm-update-repo.md`, `mm-doctor.md` — Git-регламент командой, а не памятью агента: шаблон сессионного коммита дословно из AGENTS.md (verbatim-сверка пройдена), push строго после «да» (D-18); обновление только fetch + ff-only при чистом дереве; read-only диагностика в 6 групп
- Слаги всех 7 файлов совпадают с каталогом AGENTS.md §Команды MM LAB — форвард-ссылки стандарта закрыты

## Task Commits

Each task was committed atomically:

1. **Task 1: Процедура /mm-adopt-script** - `182a065` (docs)
2. **Task 2: Процедуры /mm-check, /mm-new-button, /mm-new-compat** - `5f047b8` (docs)
3. **Task 3: Процедуры /mm-save-session, /mm-update-repo, /mm-doctor** - `e16b36b` (docs)

## Files Created/Modified

- `agents/commands/mm-adopt-script.md` - приёмка/адаптация стороннего скрипта (гейты D-08/D-09/D-10/D-11)
- `agents/commands/mm-check.md` - прогон чекера конвенции с пересказом по MM-кодам
- `agents/commands/mm-new-button.md` - скаффолд кнопки из templates/НоваяКнопка.pushbutton
- `agents/commands/mm-new-compat.md` - добавление новой версии Revit в revit_compat.py (3 точки синхронизации)
- `agents/commands/mm-save-session.md` - сессионный коммит + Obsidian (D-16..D-18, D-21)
- `agents/commands/mm-update-repo.md` - безопасное обновление (fetch + ff-only)
- `agents/commands/mm-doctor.md` - self-check окружения и репозитория (read-only, 6 групп)

## Decisions Made

- Все 7 файлов несут полный interfaces-контракт из 5 секций (включая «Гейты и запреты» в файлах Task 2, где acceptance требовал минимум 3) — единый стиль каталога
- Эвристика авто-подсказки панели описана один раз в mm-adopt-script и переиспользована ссылкой в mm-new-button — D-10 действует одинаково
- mm-new-compat предписывает удалять `__pycache__` после `py_compile` — перенос зафиксированного pitfall verify-цепочек фазы в процедуру
- `requirements mark-complete` пропущен: CONV-* — условные метки из RESEARCH (§Phase Requirements), в REQUIREMENTS.md (генерируется из Карты релизов) их нет; ID зафиксированы в frontmatter этого SUMMARY

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness

- Каталог процедур полон — план 03-07 (адаптеры `.claude/commands/`, `.gemini/commands/*.toml`, `.kilo/commands/`) может ссылаться на все 7 канонических файлов
- Manual UAT фазового гейта (после 03-07): прогон `/mm-adopt-script` на IFC_Двери (чекер, вопрос панели, diff, отсутствие регистрации без «да») и `/mm-save-session` (коммит по шаблону, без push без подтверждения)
- Untracked в дереве: `IFC_Двери.pushbutton/`, `IFC_Окна.pushbutton/`, `docs/` — существовали до плана; IFC-кнопки намеренно не адаптированы (UAT-вход для /mm-adopt-script)

## Self-Check: PASSED

- Все 7 файлов `agents/commands/mm-*.md` существуют на диске (FOUND ×7)
- Коммиты `182a065`, `5f047b8`, `e16b36b` в истории (FOUND ×3)
- Все три automated-verify задач повторно выполнены после завершения — exit 0
- Плановая верификация: ровно 7 файлов; слаги совпадают с каталогом AGENTS.md; key_links (check_convention.py, templates/, .planning/quick/) на месте

---
*Phase: 03-convention*
*Completed: 2026-07-24*
