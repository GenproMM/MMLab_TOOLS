---
gsd_state_version: 1.0
milestone: v1.0
milestone_name: milestone
status: executing
stopped_at: Completed 03-05-PLAN.md
last_updated: "2026-07-24T12:10:48.818Z"
last_activity: 2026-07-24
progress:
  total_phases: 1
  completed_phases: 0
  total_plans: 7
  completed_plans: 5
  percent: 71
---

# Project State

## Project Reference

See: .planning/PROJECT.md (updated 2026-06-09)

**Core value:** Надежные и предсказуемые pyRevit-скрипты, которые сокращают ручной труд без риска повредить модель.
**Current focus:** Phase 03 — convention

## Current Position

Phase: 03 (convention) — EXECUTING
Plan: 6 of 7
Status: Ready to execute
Last activity: 2026-07-24

Progress: [██████░░░░] 67%

## Performance Metrics

**Velocity:**

- Total plans completed: 3
- Average duration: 1 day (phase-local)
- Total execution time: 1.0 hours

## Accumulated Context

### Decisions

- [Phase 1]: Общие helper-функции ИОС должны быть вынесены в верхнеуровневый shared-модуль
- [Phase 03-convention]: Чекер конвенции: bundle.yaml разбирается построчным stdlib-парсером ограниченной схемы (без PyYAML); ограничение задокументировано в docstring
- [Phase 03-convention]: Пути нарушений чекера — POSIX-relpath от --root: стабильные ключи для baseline grandfathering legacy-кнопок
- [Phase 03-convention]: revit_compat: детекция версии каскадом (аргумент -> builtins.__revit__ -> pyrevit.HOST_APP через .version); при неопределённой версии units-ветка 2022+, гейт версии обязателен до convert_*
- [Phase 03-convention]: Белый список импортов MM008 строится динамически из стемов MM LAB.extension/lib/*.py без чтения содержимого (T-03-08); root=None — сырой режим без first-party
- [Phase 03-convention]: MM014: единственная чистая форма lib-бутстрапа — sys.path.insert(0, _LIB_DIR); триггеры — имя EXTENSION_ROOT, 4+ '..' в os.path.join, любой иной sys.path-вызов (D-15)
- [Phase 03-convention]: Baseline конвенции генерируется только фактическим аудитом (--write-baseline): 21 юнит, 202 нарушения заморожены; записи снимаются при адаптации кнопки (/mm-adopt-script)
- [Phase 03-convention]: Verify-цепочки: py_compile создаёт __pycache__ в папке кнопки и роняет следующий strict-прогон (MM013) — после py_compile чистить __pycache__ (в .gitignore уже покрыт)
- [Phase 03-convention]: AGENTS.md: канонический бутстрап-блок размещён без отступа списка — дословная копипаста шаблона; запрет легаси-формы MM014 описан без литерального легаси-имени переменной (гейт acceptance)
- [Phase 03-convention]: Белый список импортов в AGENTS.md сверен с HOST_IMPORT_ROOTS чекера (включая Microsoft), а не с сокращённой формулировкой плана — стандарт совпадает с реализацией до буквы

### Pending Todos

- Выполнить Revit smoke UAT для IOS-03 на контрольной модели.

### Blockers/Concerns

- Functional UAT в Revit не выполняется в headless-среде и вынесен в follow-up.

### Quick Tasks Completed

| # | Description | Date | Commit | Directory |
|---|-------------|------|--------|-----------|
| 260709-jko | ИОС panel: AttributeError на RBS_DUCT_LOSS_METHOD_SERVER_PARAM + bool/string outcome mismatch в ensure_loss_method_undefined | 2026-07-09 | (pending) | [260709-jko-get-parameter-typeerror-builtinparameter](./quick/260709-jko-get-parameter-typeerror-builtinparameter/) |
| Phase 03-convention P01 | 17 min | 2 tasks | 10 files |
| Phase 03-convention P02 | 10 min | 2 tasks | 2 files |
| Phase 03-convention P03 | 12 min | 2 tasks | 3 files |
| Phase 03-convention P04 | 10 min | 2 tasks | 5 files |
| Phase 03-convention P05 | 9 min | 2 tasks | 4 files |

### Roadmap Evolution

- Phase 1 established at initialization: Дедупликация общих helper-функций ИОС
- Phase 2 added: Проанализировать все скрипты на наличие общих повторяющихся функций, вынести их в ./lib и добавить импорты в скриптах
- Phase 3 added (целевой релиз v260724): Конвенция правил скриптов MM LAB и Claude-команда проверки/адаптации сторонних скриптов при добавлении в MM LAB.tab

## Session Continuity

Last session: 2026-07-24T12:10:31.283Z
Stopped at: Completed 03-05-PLAN.md
Resume file: None
