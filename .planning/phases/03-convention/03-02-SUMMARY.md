---
phase: 03-convention
plan: 02
subsystem: lib
tags: [pyrevit, revit-api, pythonnet, cpython, compat, ast, unittest]

# Dependency graph
requires:
  - phase: 01-helper
    provides: "паттерн shared-модулей в MM_LAB.extension/lib (ios_common_helpers — стиль импортов, element_id_value, обход BIP->имя)"
provides:
  - "MM_LAB.extension/lib/revit_compat.py — единый compat-модуль Revit API: SUPPORTED_VERSIONS (2020, 2022, 2024) + 13 публичных функций (D-01..D-04)"
  - "Канонический каскад get_parameter из 5 ступеней — консолидация двух существующих обходов репо (третий вариант не создан)"
  - "tools/tests/test_revit_compat_contract.py — ast-контракт публичного API, работает без Revit"
affects: [03-04 (шаблон кнопки), 03-05 (AGENTS.md), 03-06 (mm-команды), все будущие кнопки]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Версионные ветвления Revit API живут только внутри revit_compat (D-01)"
    - "Ленивая инициализация версионных карт (_units_map): импорт модуля не зависит от версии Revit"
    - "Контрактный тест через ast.parse исходника — без импорта Revit-зависимого кода"

key-files:
  created:
    - "MM_LAB.extension/lib/revit_compat.py"
    - "tools/tests/test_revit_compat_contract.py"
  modified: []

key-decisions:
  - "Каскад детекции версии дополнен третьей попыткой .version — без неё ветка pyrevit.HOST_APP не могла вернуть версию (у HOST_APP нет .Application.VersionNumber/.VersionNumber)"
  - "При неопределённой версии _units_map берёт современную ветку 2022+ (гейт require_supported_version обязан отработать раньше по конвенции)"
  - "create_floor бросает ValueError на пустой curve_loops; convert_* — ValueError на неизвестный unit_key (понятная ошибка вместо невнятного падения API)"
  - "Коммит Task 2 использует формулировку из acceptance criteria плана: feat(03-02): revit_compat + контрактный тест API"

patterns-established:
  - "compat-каскад параметров: прямой get_Parameter -> __overloads__[BuiltInParameter] -> BIP->имя через ParameterElement (кеш _BIP_NAME_CACHE) -> LookupParameter по fallback-именам -> None"
  - "ast-контракт: список PUBLIC_API фиксируется в тесте дословно, дрейф публичного API ловится локально без Revit"

requirements-completed: [CONV-STD]

# Coverage metadata — one entry per shipped deliverable
coverage:
  - id: D1
    description: "Единый compat-модуль revit_compat.py: SUPPORTED_VERSIONS = (2020, 2022, 2024), 13 публичных функций, компилируется, версионные if только внутри compat"
    requirement: CONV-STD
    verification:
      - kind: other
        ref: "py -3 -m py_compile \"MM_LAB.extension/lib/revit_compat.py\""
        status: pass
      - kind: unit
        ref: "tools/tests/test_revit_compat_contract.py#TestRevitCompatContract"
        status: pass
    human_judgment: false
  - id: D2
    description: "Контрактный тест защищает шапку/BOM, докстринг, SUPPORTED_VERSIONS, полный публичный API (13 функций), запрет голых except и компилируемость"
    requirement: CONV-STD
    verification:
      - kind: unit
        ref: "py -3 -m unittest discover -s tools/tests -p \"test_revit_compat*.py\" -q (6 тестов, OK)"
        status: pass
    human_judgment: false
  - id: D3
    description: "Хелперы compat работают в реальном Revit 2020/2022/2024 (fail-fast TaskDialog на прочих версиях, convert_from_internal, element_id_value)"
    requirement: CONV-STD
    verification: []
    human_judgment: true
    rationale: "Revit API доступен только внутри Revit; headless-среда не позволяет runtime-проверку. UAT-чек-лист вынесен на фазовый гейт (прецедент Phase 1), см. 03-02-PLAN §verification Manual UAT"

# Metrics
duration: 10min
completed: 2026-07-24
status: complete
---

# Phase 3 Plan 2: revit_compat — единый compat-модуль Revit API Summary

**Единый compat-модуль `revit_compat.py` (детекция версии, fail-fast D-03, канонический каскад get_parameter, ElementId Int64 2024, Units 2020/2022+, Floor.Create, pythonnet-interop, ensure_vendor_lib) + ast-контрактный тест публичного API из 13 функций**

## Performance

- **Duration:** 10 min
- **Started:** 2026-07-24T11:10:23Z
- **Completed:** 2026-07-24T11:19:56Z
- **Tasks:** 2
- **Files modified:** 2 (созданы)

## Accomplishments

- Создан `MM_LAB.extension/lib/revit_compat.py` (452 строки): единственное место версионных ветвлений Revit API в репозитории (D-01..D-04) — скрипты кнопок теперь зовут стабильные хелперы, а не сырой версионный API
- Каскад get_parameter консолидировал оба существующих обхода pythonnet (`__overloads__` из «Мокрых зон» и BIP->имя->LookupParameter из ios_common_helpers) — третий вариант в репо не появился
- `require_supported_version` реализует fail-fast (D-03): TaskDialog со списком поддерживаемых версий + SystemExit; сообщение на русском («Обратись в GENPRO LAB»)
- Контрактный тест `tools/tests/test_revit_compat_contract.py` (6 тестов) фиксирует шапку, докстринг, SUPPORTED_VERSIONS и все 13 публичных функций через ast — без импорта Revit; полный прогон tools/tests (27 тестов вместе с планом 03-01) зелёный

## Task Commits

Each task was committed atomically:

1. **Task 1: Создать MM_LAB.extension/lib/revit_compat.py** - `930da43` (feat)
2. **Task 2: Контрактный тест API compat (ast, без Revit)** - `1c72ecb` (feat)

## Files Created/Modified

- `MM_LAB.extension/lib/revit_compat.py` - compat-модуль: SUPPORTED_VERSIONS, get_revit_version, require_supported_version, get_parameter, get_shared_parameter, element_id_value, make_element_id, convert_from_internal, convert_to_internal, create_floor, to_net_list, enum_from_int, iter_count, ensure_vendor_lib; приватные _version_number, _bip_to_lookup_name, _units_map, _unit_object, _BIP_NAME_CACHE
- `tools/tests/test_revit_compat_contract.py` - PUBLIC_API (13 имён) + TestRevitCompatContract: test_header, test_docstring, test_supported_versions, test_public_api, test_no_bare_except, test_compiles

## Decisions Made

- **Третья попытка `.version` в `_version_number`:** план предписывал извлекать версию через `.Application.VersionNumber` либо `.VersionNumber`, но у `pyrevit.HOST_APP` (обёртка _HostApplication) версия доступна как `.version` — без третьей попытки задекларированная планом ступень каскада была бы мёртвой
- **Неопределённая версия в `_units_map` → ветка 2022+:** покрывает 2 из 3 поддерживаемых версий; по конвенции скрипт обязан вызвать `require_supported_version` раньше любых convert_*
- **Защитные ValueError:** пустой `curve_loops` в create_floor (ветка 2020) и неизвестный `unit_key` в convert_* дают понятную русскую ошибку вместо невнятного падения Revit API
- **Транзакции:** compat не открывает транзакций и не пишет в модель — зафиксировано в докстрингах модуля и create_floor (threat T-03-05 mitigated)

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Мёртвая ступень каскада pyrevit.HOST_APP в get_revit_version**
- **Found during:** Task 1 (создание revit_compat.py)
- **Issue:** План задаёт извлечение версии только через `.Application.VersionNumber` / `.VersionNumber`, но у `pyrevit.HOST_APP` нет ни того, ни другого атрибута — третья ступень каскада никогда не вернула бы версию
- **Fix:** В `_version_number` добавлена третья попытка `int(obj.version)` (свойство _HostApplication), все попытки в try/except Exception
- **Files modified:** MM_LAB.extension/lib/revit_compat.py
- **Verification:** py_compile + контрактный тест зелёные; логика каскада покрыта докстрингом
- **Committed in:** 930da43 (Task 1 commit)

**2. [Rule 2 - Missing critical] Защитные ошибки на некорректный ввод**
- **Found during:** Task 1 (создание revit_compat.py)
- **Issue:** План не оговаривал поведение create_floor при пустом curve_loops и convert_* при неизвестном unit_key — без защиты падение было бы невнятным (NullReference/KeyError внутри API)
- **Fix:** `create_floor` бросает ValueError «пустой список контуров»; `_unit_object` бросает ValueError с перечнем доступных ключей
- **Files modified:** MM_LAB.extension/lib/revit_compat.py
- **Verification:** py_compile + контрактный тест зелёные
- **Committed in:** 930da43 (Task 1 commit)

---

**Total deviations:** 2 auto-fixed (1 bug, 1 missing critical)
**Impact on plan:** Оба исправления необходимы для корректности каскада детекции версии и предсказуемых ошибок. Публичный API и контракт не изменены — scope creep отсутствует.

## Issues Encountered

- Требование `CONV-STD` — условная метка фазы (из ROADMAP §Phase 3): в `.planning/REQUIREMENTS.md` (генерируется из «Карты релизов») такого ID нет, поэтому `requirements mark-complete` неприменим — зафиксировано в requirements-completed фронтматтера SUMMARY (ожидаемое поведение по 03-RESEARCH §Phase Requirements)

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness

- Публичный API compat зафиксирован контрактом — планы 03-04 (шаблон кнопки), 03-05 (AGENTS.md) и 03-06 (mm-команды) могут ссылаться на стабильные сигнатуры
- Manual UAT в Revit 2020/2022/2024 (fail-fast, convert, element_id_value) вынесен на фазовый гейт — headless-среда не позволяет runtime-проверку
- Wave 1 фазы завершена (03-01 чекер + 03-02 compat) — готовность к wave 2 (03-03, AST-правила чекера)

## Self-Check: PASSED

- FOUND: MM_LAB.extension/lib/revit_compat.py
- FOUND: tools/tests/test_revit_compat_contract.py
- FOUND: commit 930da43
- FOUND: commit 1c72ecb
- PASS: `py -3 -m py_compile "MM_LAB.extension/lib/revit_compat.py"` → exit 0
- PASS: `py -3 -m unittest discover -s tools/tests -p "test_revit_compat*.py" -q` → exit 0 (6 тестов)

---
*Phase: 03-convention*
*Completed: 2026-07-24*
