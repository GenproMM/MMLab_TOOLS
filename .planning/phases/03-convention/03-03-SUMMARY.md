---
phase: 03-convention
plan: "03"
subsystem: tooling
tags: [python, ast, linter, unittest, pyrevit, convention]

requires:
  - phase: 03-convention (план 03-01)
    provides: каркас tools/check_convention.py (MM000–MM007, MM013), CLI, baseline, фикстуры repo_ok/repo_bad
provides:
  - AST-правила MM008–MM012 в check_convention.py (белый список импортов, wildcard, LookupParameter-литерал, голый except, pyrevit.forms)
  - Правило MM014 — машинная проверка канонического lib-бутстрапа D-15 (_EXTENSION_DIR/_LIB_DIR)
  - Функции allowed_import_roots(root) и check_ast_rules(tree, allowed_roots)
  - check_script(path, root=None) — полный AST-анализ сырого .py до скаффолда
  - Расширенная плохая фикстура с реальными дефектами IFC_Двери (BOM сохранён)
affects: [03-04 baseline, 03-06 mm-check, mm-adopt-script, приёмка сторонних скриптов]

tech-stack:
  added: []
  patterns:
    - "AST-анализ недоверенного кода: только ast.walk по распарсенному дереву, без import/exec (T-03-01)"
    - "Динамический white list first-party: стемы lib/*.py без чтения содержимого (T-03-08)"

key-files:
  created: []
  modified:
    - tools/check_convention.py
    - tools/tests/test_check_convention.py
    - "tools/tests/fixtures/repo_bad/MM LAB.extension/MM Lab.tab/Плохая панель.panel/Плохая кнопка.pushbutton/script.py"

key-decisions:
  - "Белый список импортов MM008 строится динамически из стемов MM LAB.extension/lib/*.py без чтения содержимого; root=None — сырой режим без first-party"
  - "MM014: единственная чистая форма lib-бутстрапа — sys.path.insert(0, _LIB_DIR); триггеры — имя EXTENSION_ROOT, 4+ '..' в os.path.join, любой иной sys.path-вызов"
  - "check_script получил необязательный параметр root (обратная совместимость: 1-аргументный вызов работает как прежде) — CLI-контракт не менялся"

patterns-established:
  - "Каждое AST-правило: тест на плохом входе + тест-гард на чистом (канонический бутстрап, хорошая фикстура)"

requirements-completed: [CONV-CHECK]

coverage:
  - id: D1
    description: "AST-правила MM008–MM012: сторонний импорт, wildcard, LookupParameter-литерал, голый except, pyrevit.forms детектируются; хост/stdlib/first-party не дают ложных срабатываний"
    requirement: CONV-CHECK
    verification:
      - kind: unit
        ref: "tools/tests/test_check_convention.py#AstRuleTests.test_mm008_third_party_import / test_mm008_allows_host_and_stdlib / test_mm008_allows_first_party_lib / test_mm009_wildcard / test_mm010_lookup_literal / test_mm011_bare_except / test_mm012_pyrevit_forms"
        status: pass
    human_judgment: false
  - id: D2
    description: "MM014: неканонический lib-бутстрап (EXTENSION_ROOT, 4+ '..', неканонический sys.path) детектируется; канонический блок D-15 проходит чисто"
    requirement: CONV-CHECK
    verification:
      - kind: unit
        ref: "tools/tests/test_check_convention.py#AstRuleTests.test_mm014_extension_root_name / test_mm014_four_parent_hops / test_mm014_noncanonical_syspath / test_mm014_canonical_bootstrap_clean / test_mm014_on_bad_fixture"
        status: pass
    human_judgment: false
  - id: D3
    description: "AST-правила работают в режиме сырого .py (приёмка до скаффолда); хорошая фикстура по-прежнему без нарушений"
    requirement: CONV-CHECK
    verification:
      - kind: unit
        ref: "tools/tests/test_check_convention.py#AstRuleTests.test_raw_script_ast_rules / test_good_button_still_clean"
        status: pass
      - kind: other
        ref: "py -3 tools/check_convention.py --all --root tools/tests/fixtures/repo_ok (exit 0)"
        status: pass
      - kind: other
        ref: "py -3 tools/check_convention.py <плохая кнопка> --root tools/tests/fixtures/repo_bad --json (exit 1; JSON содержит MM008, MM009, MM010, MM011, MM012, MM014)"
        status: pass
    human_judgment: false

duration: 12 min
completed: 2026-07-24
status: complete
---

# Phase 03 Plan 03: AST-правила чекера конвенции Summary

**AST-правила MM008–MM012 и MM014 в tools/check_convention.py методом TDD: белый список импортов (stdlib + clr/System/Autodesk/pyrevit/Microsoft + first-party lib + vendored openpyxl/et_xmlfile), wildcard-импорт, LookupParameter-литерал, голый except, pyrevit.forms и машинная проверка канонического lib-бутстрапа D-15 — 15 новых тестов, 42/42 зелёные**

## Performance

- **Duration:** 12 min
- **Started:** 2026-07-24T11:28:07Z
- **Completed:** 2026-07-24T11:40:10Z
- **Tasks:** 2 (RED + GREEN; REFACTOR не потребовался)
- **Files modified:** 3

## Accomplishments

- Чекер теперь покрывает все правила конвенции MM000–MM014: именно MM008–MM012 отличают «сторонний скрипт» от скрипта по конвенции (реальные дефекты IFC_Двери: wildcard-импорт, LookupParameter("GP_23_Назначение"), голые except, pyrevit.forms)
- MM014 закрывает реальный класс дефекта бутстрапа («Мокрые зоны», «Экспорт ПСО»): EXTENSION_ROOT / 4×'..' в корень репо детектируются; единственная чистая форма — канонический блок _EXTENSION_DIR/_LIB_DIR (D-06/D-15)
- Формулировка конвенции «без сторонних импортов, кроме vendored lib/» стала машинно проверяемой: vendored-исключение (openpyxl, et_xmlfile) в белом списке
- Режим сырого .py даёт полный AST-анализ до скаффолда — гейт для /mm-adopt-script
- Плохая фикстура расширена живыми конструкциями дефектов с сохранением UTF-8 BOM (бинарная перезапись)

## TDD

- **RED:** 15 новых тестов AstRuleTests + расширение плохой фикстуры; прогон падал ровно по новым правилам (13 errors: отсутствующие allowed_import_roots/check_ast_rules и сигнатура check_script; 2 failures: MM014/MM008 не найдены в результатах), 21 тест плана 03-01 оставался зелёным — `ec1e27b`
- **GREEN:** allowed_import_roots + check_ast_rules (один обход ast.walk) + интеграция в _check_script_file → check_script/check_pushbutton; 36/36 тестов чекера, 42/42 всего tools/tests — `736176e`
- **REFACTOR:** не потребовался — хелперы атомарные, дублирования нет

## Task Commits

1. **Task 1: RED — падающие тесты MM008–MM012, MM014 + фикстура** - `ec1e27b` (test)
2. **Task 2: GREEN — AST-правила MM008–MM012, MM014** - `736176e` (feat)

## Files Created/Modified

- `tools/check_convention.py` - функции allowed_import_roots и check_ast_rules; правила MM008–MM012, MM014; check_script(path, root=None); docstring-таблица MM000–MM014
- `tools/tests/test_check_convention.py` - класс AstRuleTests (15 тестов), хелпер write_tmp_script
- `tools/tests/fixtures/repo_bad/.../Плохая кнопка.pushbutton/script.py` - requests, wildcard, pyrevit.forms, LookupParameter-литерал, голый except, legacy-бутстрап EXTENSION_ROOT (BOM в первых 3 байтах сохранён)

## Decisions Made

- Белый список MM008 строится ТОЛЬКО из стемов файлов `<root>/MM LAB.extension/lib/*.py` (root нормализован, содержимое не читается/не исполняется) — митигация T-03-08 из threat model
- Канон sys.path (MM014) — строго `insert(0, _LIB_DIR)`: Constant int 0 (bool отвергается проверкой типа) + Name ровно `_LIB_DIR`
- MM014(а) срабатывает на каждое вхождение имени EXTENSION_ROOT (Load и Store) — по букве контракта «ast.Name в любом контексте»; несколько записей одного кода не мешают baseline-механике (пары путь+код)
- check_script расширен необязательным root вместо новой функции — CLI-контракт и 1-аргументные вызовы из тестов 03-01 не тронуты

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

- `requirements mark-complete CONV-CHECK` пропущен: `.planning/REQUIREMENTS.md` генерируется из «Карты релизов» и содержит только задачи плагинов (T-1xx/T-2xx/T-3xx); условный ID CONV-CHECK в файле отсутствует (ожидаемо, см. 03-RESEARCH §Phase Requirements). В frontmatter SUMMARY requirements-completed заполнен по PLAN.

## Known Stubs

Нет — дефекты в плохой фикстуре являются целевыми тестовыми данными, а не заглушками.

## Threat Flags

Нет новой поверхности: правила анализируют недоверенный код только через ast.walk по уже распарсенному дереву (T-03-01 сохранена), first-party список — из имён файлов без исполнения (T-03-08 реализована).

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness

- Чекер покрывает все правила MM000–MM014 — готов к генерации baseline legacy-кнопок в плане 03-04
- Гейт приёмки --strict теперь ловит и legacy-бутстрап: /mm-adopt-script (план 03-06) может опираться на полный набор правил

## Self-Check: PASSED

- FOUND: tools/check_convention.py
- FOUND: tools/tests/test_check_convention.py
- FOUND: tools/tests/fixtures/repo_bad/.../Плохая кнопка.pushbutton/script.py
- FOUND: commit ec1e27b (test)
- FOUND: commit 736176e (feat)
- TESTS: `py -3 -m unittest discover -s tools/tests -q` → 42 tests OK (exit 0)

---
*Phase: 03-convention*
*Completed: 2026-07-24*
