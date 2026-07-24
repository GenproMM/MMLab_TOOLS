---
gsd_state_version: 1.0
milestone: v1.0
milestone_name: milestone
current_phase: 03
current_phase_name: convention
status: executing
stopped_at: Completed 03-03-PLAN.md
last_updated: "2026-07-24T11:42:16.801Z"
last_activity: 2026-07-24
last_activity_desc: Phase 03 execution started
progress:
  total_phases: 1
  completed_phases: 0
  total_plans: 7
  completed_plans: 3
  percent: 0
---

# Project State

## Project Reference

See: .planning/PROJECT.md (updated 2026-06-09)

**Core value:** Надежные и предсказуемые pyRevit-скрипты, которые сокращают ручной труд без риска повредить модель.
**Current focus:** Phase 03 — convention

## Current Position

Phase: 03 (convention) — EXECUTING
Plan: 4 of 7
Status: Ready to execute
Last activity: 2026-07-24 — Phase 03 execution started

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

### Roadmap Evolution

- Phase 1 established at initialization: Дедупликация общих helper-функций ИОС
- Phase 2 added: Проанализировать все скрипты на наличие общих повторяющихся функций, вынести их в ./lib и добавить импорты в скриптах
- Phase 3 added (целевой релиз v260724): Конвенция правил скриптов MM LAB и Claude-команда проверки/адаптации сторонних скриптов при добавлении в MM LAB.tab

## Session Continuity

Last session: 2026-07-24T11:40:52.612Z
Stopped at: Completed 03-03-PLAN.md
Resume file: None
