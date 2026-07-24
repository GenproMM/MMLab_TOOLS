---
phase: 03-convention
plan: 01
subsystem: tooling
tags: [pyrevit, convention, checker, cli, unittest, tdd, stdlib, baseline]

# Dependency graph
requires: []
provides:
  - "tools/check_convention.py — stdlib-only CLI-чекер конвенции MM LAB: правила MM000–MM007, MM013"
  - "Контракт CLI: PATHS/--all/--root/--strict/--json/--baseline/--write-baseline, exit-коды 0/1/2"
  - "Публичные функции для расширения: Violation, check_pushbutton, check_script, check_layouts, iter_pushbuttons, load_baseline, apply_baseline, write_baseline, main"
  - "Схема baseline JSON (generated/note/units) для grandfathered legacy-кнопок"
  - "Фикстуры tools/tests/fixtures/repo_ok и repo_bad (кириллица + пробелы в путях)"
affects: [03-03, 03-04, 03-06, mm-check, mm-adopt-script]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Построчный парсер bundle.yaml ограниченной схемы (без PyYAML)"
    - "Пути нарушений — POSIX-relpath от --root (стабильные ключи baseline)"
    - "Статический анализ без исполнения кода: только ast.parse + чтение байтов"
    - "Baseline grandfathering: legacy-нарушения замораживаются, --strict их игнорирует"

key-files:
  created:
    - tools/check_convention.py
    - tools/tests/test_check_convention.py
    - "tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/bundle.yaml"
    - "tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/Тестовая панель.panel/bundle.yaml"
    - "tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/Тестовая панель.panel/Хорошая кнопка.pushbutton/script.py"
    - "tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/Тестовая панель.panel/Хорошая кнопка.pushbutton/bundle.yaml"
    - "tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/Тестовая панель.panel/Хорошая кнопка.pushbutton/README.md"
    - "tools/tests/fixtures/repo_bad/MM LAB.extension/MM Lab.tab/bundle.yaml"
    - "tools/tests/fixtures/repo_bad/MM LAB.extension/MM Lab.tab/Плохая панель.panel/bundle.yaml"
    - "tools/tests/fixtures/repo_bad/MM LAB.extension/MM Lab.tab/Плохая панель.panel/Плохая кнопка.pushbutton/script.py"
  modified: []

key-decisions:
  - "В layout repo_bad добавлен разделитель «-----» и заасерчено, что орфаном он не считается (контракт interfaces: 3+ символов -/> — разделитель)"
  - "MM013 не спускается внутрь мусорных папок (__pycache__/) — одно нарушение на папку, без дублей на *.pyc внутри"
  - "--strict эскалирует warning только при подсчёте exit-кода; severity в выводе остаётся warning (дословно по плану)"
  - "При SyntaxError/ValueError от ast.parse — MM000 и остановка файловых проверок; строчные проверки MM001–MM003 выполняются до парсинга"

patterns-established:
  - "Юнит проверки = папка *.pushbutton; сырой *.py — только файловые правила MM000–MM004"
  - "Тесты создают мусор для MM013 во временных копиях (в фикстуры .gitignore его не пустит)"

requirements-completed: [CONV-CHECK, CONV-REG]

coverage:
  - id: D1
    description: "CLI-чекер конвенции: правила MM000–MM007, MM013, exit-коды 0/1/2, --json/--strict/--baseline/--write-baseline/--all/--root"
    requirement: CONV-CHECK
    verification:
      - kind: unit
        ref: "py -3 -m unittest discover -s tools/tests -p \"test_check_convention*.py\" -q (21 тест)"
        status: pass
      - kind: other
        ref: "py -3 tools/check_convention.py <плохая кнопка> --root <repo_bad> → exit 1, перечислены MM001–MM007"
        status: pass
      - kind: other
        ref: "py -3 tools/check_convention.py --all --root <repo_ok> → exit 0; --all --root <repo_bad> --json → exit 1, валидный JSON"
        status: pass
    human_judgment: false
  - id: D2
    description: "Правило MM007: регистрация кнопки в layout панели + орфаны записей layout в tab/panel bundle.yaml (включая хвостовые пробелы и разделители)"
    requirement: CONV-REG
    verification:
      - kind: unit
        ref: "tools/tests/test_check_convention.py#LayoutTests (button_not_in_layout, orphan_button_entry, orphan_panel_in_tab, skipped_outside_panel)"
        status: pass
    human_judgment: false
  - id: D3
    description: "Фикстуры repo_ok (эталонная кнопка) и repo_bad (кнопка-нарушитель с BOM) с кириллическими путями и пробелами"
    verification:
      - kind: unit
        ref: "tools/tests/test_check_convention.py#BadButtonTests.test_mm003_fixture_guard (первые 3 байта == EF BB BF)"
        status: pass
      - kind: unit
        ref: "tools/tests/test_check_convention.py#GoodButtonTests.test_good_button_clean"
        status: pass
    human_judgment: false

# Metrics
duration: 17min
completed: 2026-07-24
status: complete
---

# Phase 03 Plan 01: Чекер конвенции (структурные правила) Summary

**Stdlib-only CLI-чекер конвенции MM LAB (`tools/check_convention.py`): правила MM000–MM007 и MM013, baseline/strict/json-режимы — построен методом TDD (RED → GREEN), 21 unittest-тест зелёный**

## Performance

- **Duration:** 17 min
- **Started:** 2026-07-24T10:49:21Z
- **Completed:** 2026-07-24T11:06:00Z
- **Tasks:** 2 (RED + GREEN)
- **Files modified:** 10 created

## Accomplishments

- `tools/check_convention.py` (640 строк): dataclass `Violation`, словарь `RULES` с русскими сообщениями, публичные функции `check_pushbutton`/`check_script`/`check_layouts`/`iter_pushbuttons`/`load_baseline`/`apply_baseline`/`write_baseline`/`main` — сигнатуры дословно по контракту interfaces плана
- Структурные правила: MM000 (не парсится), MM001 (`#! python3`), MM002 (coding-строка), MM003 (UTF-8 BOM запрещён), MM004 (docstring «Совместимость:»/«Зависимости:», warning), MM005 (bundle.yaml с title/tooltip), MM006 (README.md), MM007 (регистрация в layout + орфаны записей), MM013 (мусор, warning)
- CLI по контракту: exit-коды 0/1/2; `--json` — ровно один JSON-объект; `--write-baseline` фиксирует все текущие нарушения и возвращает 0; `--baseline` фильтрует пары (путь, код); `--strict` игнорирует baseline и учитывает warning в exit-коде
- Построчный парсер bundle.yaml (PyYAML не используется): ключи `title:`/`tooltip:`/`layout:`, записи `- имя` со strip (реальный кейс «ВОР  » с хвостовыми пробелами), разделители из 3+ `-`/`>` пропускаются; ограничение задокументировано в docstring
- 21 unittest-тест: каждое правило поймано на плохой фикстуре и молчит на хорошей; roundtrip baseline; эскалация warning в `--strict`; режим сырого скрипта без структурных правил
- Фикстуры с кириллицей и пробелами в путях: repo_ok (эталон по канону «Мокрых зон») и repo_bad (BOM-скрипт, орфаны «Призрак» и «Нет папки»)
- Безопасность по threat model: проверяемый код НЕ исполняется (только `ast.parse` + байты, T-03-01); пути через `Path.resolve()`, `--all` не выходит за `--root` (T-03-02); битые файлы дают MM000 без traceback (T-03-03)

## Task Commits

Каждая TDD-фаза закоммичена атомарно:

1. **Task 1: RED — фикстуры и падающие тесты** - `413d24f` (test) — прогон падал только из-за отсутствия `tools/check_convention.py`
2. **Task 2: GREEN — реализация чекера** - `a334726` (feat) — все 21 тест зелёные с первого прогона

**REFACTOR:** не потребовался — реализация прошла тесты сразу, очевидных улучшений нет (коммит по TDD-протоколу опционален).

## Files Created/Modified

- `tools/check_convention.py` — CLI-чекер конвенции (структурные правила + baseline/strict/json)
- `tools/tests/test_check_convention.py` — исполняемая спецификация: 21 тест, хелперы `run_main`/`copy_button_to_tmp`
- `tools/tests/fixtures/repo_ok/**` — эталонное дерево: tab/panel bundle.yaml + «Хорошая кнопка.pushbutton» (script.py по канону, bundle.yaml, README.md)
- `tools/tests/fixtures/repo_bad/**` — дерево-нарушитель: script.py с BOM (EF BB BF), орфан «Призрак» с хвостовыми пробелами в tab-layout, орфан «Нет папки» в panel-layout

## Decisions Made

- **Разделитель в фикстуре:** в tab-layout repo_bad добавлена запись `- -----` (сверх двух записей из текста задачи) — тест закрепляет требование interfaces «записи из 3+ `-`/`>` — разделители, пропускать»; без этого контракт остался бы непокрытым
- **MM013 без дублей:** мусорная папка даёт одно нарушение, обход внутрь неё не идёт (иначе `__pycache__/*.pyc` давал бы два нарушения на один артефакт)
- **MM000 и ValueError:** `ast.parse` кидает ValueError на NUL-байтах — обработан вместе с SyntaxError (битый ввод не роняет чекер, T-03-03)
- **check_script и путь:** функция самодостаточна (путь как передан, POSIX); `main` дополнительно нормализует путь к relpath от `--root` для стабильности baseline

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None. Консольные кракозябры cp866 при создании фикстур подтвердили актуальность Pitfall 5 — чекер решает это через `sys.stdout.reconfigure(encoding="utf-8")`.

## Known Stubs

Отсутствуют. Заглушка `main()` в фикстуре «Хорошая кнопка» — намеренная часть тестовых данных (кнопка-эталон без Revit-вызовов), а не незавершённая функциональность.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness

- Чекер готов к расширению AST-правилами MM008–MM012, MM014 в плане 03-03 (точки расширения: `RULES`, `_check_script_file`, режим сырого скрипта)
- Контракт CLI зафиксирован для планов 03-03, 03-04 (baseline реального репо), 03-06 (`/mm-check`, `/mm-adopt-script`)
- Условные REQ-ID CONV-CHECK/CONV-REG отсутствуют в `.planning/REQUIREMENTS.md` (файл генерируется из «Карты релизов» и содержит только T-1xx..T-3xx) — отметка требований выполняется на уровне ROADMAP/фазы, шаг `requirements mark-complete` неприменим

## Self-Check: PASSED

- Все 10 созданных файлов и SUMMARY существуют на диске (проверено `[ -f ]`)
- Оба коммита в истории: `413d24f` (test), `a334726` (feat)
- Верификация плана перезапущена: unittest exit 0; `--all --root repo_bad --json` → exit 1, валидный JSON

---
*Phase: 03-convention*
*Completed: 2026-07-24*
