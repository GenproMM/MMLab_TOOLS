---
phase: 03-convention
plan: 07
subsystem: agent-commands
tags: [mm-commands, adapters, claude-code, gemini-cli, kilo-code, toml, unittest, convention]

# Dependency graph
requires:
  - phase: 03-convention (03-06)
    provides: "agents/commands/ — 7 канонических процедур mm-команд (единственный полный текст каждой команды)"
  - phase: 03-convention (03-05)
    provides: "AGENTS.md §Команды MM LAB — таблица назначений 7 команд (источник description адаптеров)"
  - phase: 03-convention (03-01/03-03)
    provides: "tools/check_convention.py + convention_baseline.json — компоненты финального full-suite фазы"
provides:
  - "21 тонкий адаптер: .claude/commands/mm-*.md (7), .gemini/commands/mm-*.toml (7), .kilo/commands/mm-*.md (7)"
  - "Каждая команда доступна под ЕДИНЫМ именем /mm-<слаг> во всех трёх агентах (D-19, D-20); в Claude Code — немедленно, в Gemini/Kilo — при установке CLI"
  - "tools/tests/test_mm_commands_catalog.py — каталожный тест (7 тест-методов, входит в общий discover-прогон)"
  - "Анти-дрейф закреплён машинно: адаптеры обязаны ссылаться на существующие канонические файлы, без shell-вставок, без команд-двойников"
affects: [verify-work (Manual UAT фазы через /mm-* в Claude Code), будущие правки каталога команд]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Тонкий адаптер: метаданные (frontmatter/TOML-ключи) + 2 строки тела «прочитай agents/commands/mm-<слаг>.md и выполни» + маркер аргументов"
    - "Gemini: плоские имена файлов mm-<слаг>.toml — подпапка mm/<слаг>.toml дала бы команду /mm:<слаг> (Pitfall 9)"

key-files:
  created:
    - .claude/commands/mm-adopt-script.md
    - .claude/commands/mm-new-button.md
    - .claude/commands/mm-check.md
    - .claude/commands/mm-save-session.md
    - .claude/commands/mm-update-repo.md
    - .claude/commands/mm-doctor.md
    - .claude/commands/mm-new-compat.md
    - .gemini/commands/mm-adopt-script.toml
    - .gemini/commands/mm-new-button.toml
    - .gemini/commands/mm-check.toml
    - .gemini/commands/mm-save-session.toml
    - .gemini/commands/mm-update-repo.toml
    - .gemini/commands/mm-doctor.toml
    - .gemini/commands/mm-new-compat.toml
    - .kilo/commands/mm-adopt-script.md
    - .kilo/commands/mm-new-button.md
    - .kilo/commands/mm-check.md
    - .kilo/commands/mm-save-session.md
    - .kilo/commands/mm-update-repo.md
    - .kilo/commands/mm-doctor.md
    - .kilo/commands/mm-new-compat.md
    - tools/tests/test_mm_commands_catalog.py
  modified: []

key-decisions:
  - "Маркер аргументов ($ARGUMENTS / {{args}} / «всё после /mm-<слаг>») включён во ВСЕ 21 адаптер, а не только в команды с argument-hint — единообразие каталога; acceptance это допускает"
  - "Каталожный тест — substring-ассерты без tomllib (сохранён контракт «CPython >= 3.10» инструментов фазы); TOML-валидность всех 7 адаптеров разово подтверждена tomllib при приёмке Task 1"
  - "Комбинированный коммит из acceptance Task 2 разделён на два per-task коммита (feat адаптеры / test каталог) — протокол атомарных коммитов GSD приоритетнее формулировки плана"

patterns-established:
  - "test_no_extra_mm_files: любой файл mm-* вне SLUGS в каталогах адаптеров — красный тест (защита от опечаток-двойников T-03-22); файлы БЕЗ префикса mm- (чужие инструменты) не трогаются"
  - "Инвариант T-03-21 (без shell-вставок !{ в адаптерах) закреплён тестом test_no_shell_injection_in_adapters"

requirements-completed: [CONV-ADAPT, CONV-CHECK]

# Coverage metadata (#1602)
coverage:
  - id: D1
    description: "21 тонкий адаптер (7 слагов × Claude/Gemini/Kilo): единое имя /mm-<слаг>, ссылка на СВОЙ канонический файл, description по назначению, без дублирования процедур"
    requirement: "CONV-ADAPT"
    verification:
      - kind: unit
        ref: "tools/tests/test_mm_commands_catalog.py#test_claude_adapters/test_gemini_adapters/test_kilo_adapters/test_canonical_procedures_exist"
        status: pass
      - kind: other
        ref: "Task 1 verify: py -3 -c (21 файл на месте) + tomllib-парсинг 7 TOML + проверка BOM/ссылок/$ARGUMENTS"
        status: pass
    human_judgment: false
  - id: D2
    description: "Каталожный тест согласованности в общем прогоне + финальный full-suite фазы зелёный"
    requirement: "CONV-CHECK"
    verification:
      - kind: unit
        ref: "py -3 -m unittest discover -s tools/tests -q (49 тестов OK, включая 7 каталожных)"
        status: pass
      - kind: other
        ref: "составная verify Task 2: unittest + чекер --all --baseline + шаблон --strict + py_compile → exit 0"
        status: pass
    human_judgment: false
  - id: D3
    description: "Фактическая вызываемость /mm-* живыми агентами: /mm-check --all, /mm-doctor (отчёт-таблица), /mm-adopt-script на IFC_Двери до ревью-гейта без одобрения"
    requirement: "CONV-ADAPT"
    verification: []
    human_judgment: true
    rationale: "Слэш-команда проверяется только интерактивной сессией агента; Gemini CLI и Kilo Code не установлены (RESEARCH §Environment Availability) — их адаптеры проверены форматно. Manual UAT фазы — сценарии из 03-07-PLAN §verification"

# Metrics
duration: 9min
completed: 2026-07-24
status: complete
---

# Phase 03 Plan 07: Адаптеры mm-команд для трёх агентов Summary

**21 тонкий адаптер mm-команд (7 слагов × Claude Code / Gemini CLI / Kilo Code) поверх канонических процедур agents/commands/ + каталожный unittest на согласованность; полный full-suite фазы зелёный**

## Performance

- **Duration:** 9 min
- **Started:** 2026-07-24T12:27:29Z
- **Completed:** 2026-07-24T12:36:30Z
- **Tasks:** 2
- **Files modified:** 22 (все — созданы)

## Accomplishments

- 7 команд `/mm-<слаг>` доступны под единым именем во всех трёх агентах: Claude Code — немедленно (`.claude/commands/`), Gemini CLI / Kilo Code — при установке CLI (адаптеры созданы «вперёд», проверены форматно)
- Анти-дрейф: тело каждого адаптера — 2 строки («прочитай канонический файл и выполни»); полный текст процедур нигде не дублируется
- Каталожный тест `test_mm_commands_catalog.py`: 7 тест-методов ловят битые ссылки, короткие/пустые процедуры, shell-вставки, команды-двойники и рассинхрон с AGENTS.md
- Финальный полный прогон фазы: 49 unittest OK + чекер `--all` с baseline (21 юнит, 0/0) + шаблон `--strict` (0/0) + `py_compile` — exit 0

## Task Commits

Each task was committed atomically:

1. **Task 1: Создать 21 адаптер (Claude / Gemini / Kilo)** - `6529c5a` (feat)
2. **Task 2: Каталожный тест согласованности + финальный полный прогон** - `cf7211f` (test)

**Plan metadata:** коммит этого SUMMARY (`docs(03-07)`)

## Files Created/Modified

- `.claude/commands/mm-*.md` (7) — адаптеры Claude Code: YAML-frontmatter (`description`, у 4 команд `argument-hint`), тело — ссылка на канонический файл + `$ARGUMENTS`
- `.gemini/commands/mm-*.toml` (7) — адаптеры Gemini CLI: плоские имена, ключи `description` и `prompt` с `{{args}}`
- `.kilo/commands/mm-*.md` (7) — адаптеры Kilo Code: frontmatter `description`, тело — ссылка + «всё после /mm-<слаг>»
- `tools/tests/test_mm_commands_catalog.py` — каталожный тест: `SLUGS` (7), `REPO`, 7 тест-методов

## Verification Log

- Task 1 verify (`py -3 -c` — 21 файл на месте, плоские имена): exit 0
- Приёмка Task 1: tomllib-парсинг 7 TOML валиден; все ссылки ведут на существующие `agents/commands/mm-*.md`; `$ARGUMENTS`/`{{args}}` на месте; ни одного `!{`; BOM отсутствует; `.claude/skills/` не тронут (git status пуст)
- Task 2 verify (составная): `unittest discover` (49 OK) && чекер `--all --baseline` (exit 0) && `templates/НоваяКнопка.pushbutton --strict` (exit 0) && `py_compile` (exit 0)
- Плановая верификация: `ls .claude/commands .gemini/commands .kilo/commands` — ровно по 7 файлов `mm-*` в каждом каталоге
- Контрольный strict-прогон шаблона после чистки `__pycache__`: exit 0

## Decisions Made

- Маркер аргументов включён во все 21 адаптер (не только в 4 с `argument-hint`) — единообразие; для Gemini `{{args}}` во всех 7 требовался acceptance-критерием
- Каталожный тест не использует `tomllib` (появился в 3.11), а проверяет TOML substring-ассертами — сохранён контракт «CPython >= 3.10» инструментов фазы; полная TOML-валидация выполнена разово при приёмке Task 1
- Плановый единый коммит разделён на два per-task коммита (см. Deviations)

## Deviations from Plan

**1. [Протокол GSD — гранулярность коммитов] Один коммит из acceptance Task 2 разделён на два per-task**
- **Found during:** Task 1/Task 2 (коммит-этапы)
- **Issue:** acceptance-критерий Task 2 описывал единый коммит «feat(03-07): адаптеры mm-команд для Claude/Gemini/Kilo + каталожный тест», что противоречит протоколу исполнителя «каждый таск — отдельный атомарный коммит»
- **Fix:** два коммита: `6529c5a` (feat — 21 адаптер) и `cf7211f` (test — каталожный тест); суммарно покрывают ровно тот же состав файлов и намерение сообщения
- **Files modified:** нет (только структура коммитов)
- **Verification:** `git log --oneline` — оба коммита с префиксом `(03-07)`
- **Committed in:** 6529c5a, cf7211f

---

**Total deviations:** 1 (протокольная, без изменения состава работ)
**Impact on plan:** нулевой для артефактов — все must_haves и acceptance-проверки содержимого выполнены дословно; отличие только в разбиении истории git.

## Issues Encountered

- `py_compile` (последний шаг составной verify) создал `__pycache__/` в `templates/НоваяКнопка.pushbutton` и `MM LAB.extension/lib` — удалены сразу после прогона (известный pitfall фазы: мусор роняет следующий `--strict` по MM013); контрольный strict-прогон после чистки — exit 0

## Known Stubs

Нет — все 22 файла функционально полные; «неактивность» Gemini/Kilo-адаптеров — свойство окружения (CLI не установлены), а не заглушка: файлы полностью рабочие и подхватятся установленным CLI без правок.

## User Setup Required

None - no external service configuration required. Для активации команд в Gemini CLI / Kilo Code достаточно установить соответствующий CLI — адаптеры уже на месте.

## Next Phase Readiness

- Фаза 03 полностью исполнена: 7/7 планов, все артефакты на месте (стандарт, чекер+baseline, compat, шаблон, процедуры, адаптеры, тесты)
- Готово к `/gsd-verify-work 03`; Manual UAT фазового гейта: `/mm-check --all`, `/mm-doctor`, `/mm-adopt-script` на IFC_Двери до ревью-гейта (сценарий CONV-ADAPT из 03-VALIDATION.md)
- IFC_Двери/IFC_Окна остаются неадаптированными (untracked) — намеренно: это живой материал для UAT-прогона `/mm-adopt-script` отдельным quick task

## Self-Check: PASSED

- 22/22 созданных файла на диске (`[ -f ]`)
- Коммиты `6529c5a`, `cf7211f` найдены в `git log`
- Плановая verify-цепочка воспроизведена с exit 0

---
*Phase: 03-convention*
*Completed: 2026-07-24*
