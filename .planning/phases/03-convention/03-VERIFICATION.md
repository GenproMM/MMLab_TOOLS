---
phase: 03-convention
verified: 2026-07-24T13:30:00Z
status: human_needed
score: 34/34 must-haves verified
human_verification:
  - test: "Compat-хелперы в реальном Revit 2020 / 2022 / 2024"
    expected: "Кнопка-пример из templates запускается; require_supported_version пропускает поддерживаемые версии; convert_from_internal(1000, \"mm\") возвращает число; element_id_value возвращает int (2024 — Int64 .Value); на иной версии — TaskDialog «поддерживает 2020 / 2022 / 2024» и мягкий выход. Также подтверждает семантику фикса WR-06 (_VALIDATED_VERSION)"
    why_human: "Revit API доступен только внутри Revit; headless-среда не позволяет runtime-проверку (прецедент Phase 1, 03-VALIDATION.md §Manual-Only)"
  - test: "/mm-adopt-script на IFC_Двери до ревью-гейта (сценарий CONV-ADAPT)"
    expected: "Чекер отработал (--json), панель спрошена с авто-подсказкой (D-10), diff/сводка правок показаны, без явного «да» регистрация в bundle.yaml НЕ выполняется (D-08)"
    why_human: "Соблюдение блокирующих гейтов проверяется только живым интерактивным прогоном slash-команды в Claude Code"
  - test: "Quick task GSD после приёмки (CONV-GSD)"
    expected: "После одобрения приёмки создана папка .planning/quick/<id>-adopt-<слаг>/ с <id>-PLAN.md и <id>-SUMMARY.md + строка в таблице «Quick Tasks Completed» в .planning/STATE.md"
    why_human: "Агентская команда исполняется интерактивно; артефакты появляются только при живом прогоне приёмки"
  - test: "Кнопка из шаблона видна на панели после pyRevit Reload"
    expected: "Скопированный в <Панель>.panel и зарегистрированный в layout шаблон появляется на панели MM LAB и запускается (диалоги, транзакция, отчёт «Стен в проекте: N»)"
    why_human: "pyRevit парсит расширение на старте; проверка видимости кнопки требует живого Revit + Reload"
  - test: "Вызываемость /mm-check --all и /mm-doctor как slash-команд"
    expected: "/mm-check --all пересказывает результат чекера по MM-кодам; /mm-doctor выдаёт отчёт-таблицу «Проверка / Статус / Что делать» без правок"
    why_human: "Slash-команда проверяется только интерактивной сессией агента (Gemini CLI / Kilo Code не установлены — их адаптеры проверены форматно)"
---

# Phase 3: Конвенция скриптов MM LAB — Verification Report

**Phase Goal:** Зафиксировать конвенцию написания скриптов MM LAB и автоматизировать приём сторонних скриптов в `MM LAB.tab` через Claude-команду, которая проверяет скрипт на соответствие конвенции, адаптирует его и заводит задачу GSD.
**Verified:** 2026-07-24
**Status:** human_needed (все автоматические проверки пройдены; 5 сценариев требуют живого Revit / интерактивной сессии)
**Re-verification:** No — initial verification

## Goal Achievement

Обе половины цели фазы подтверждены кодом:
1. **Конвенция зафиксирована** — AGENTS.md (323 строки, канонический стандарт) + машинный гейт `tools/check_convention.py` (15 правил MM000–MM014) + compat-модуль `revit_compat.py` + шаблон `templates/НоваяКнопка.pushbutton/`.
2. **Приём сторонних скриптов автоматизирован** — каноническая процедура `agents/commands/mm-adopt-script.md` (чекер → панель → адаптация → strict-гейт → ревью-гейт → регистрация в bundle.yaml → quick task GSD) + 21 тонкий адаптер (Claude/Gemini/Kilo) + каталожный тест.

### Observable Truths

Must-haves объединены из frontmatter всех 7 планов (ROADMAP фазы 3 не содержит отдельного массива success_criteria; контракт — Goal + «Объём работ», покрытые планами полностью).

| #  | Truth (план) | Status | Evidence |
|----|--------------|--------|----------|
| 1  | Чекер на плохой фикстуре → exit 1, MM001/MM002/MM003/MM005/MM006 (03-01) | ✓ VERIFIED | Прогон: exit 1, JSON содержит все 13 кодов MM001–MM012, MM014 |
| 2  | Хорошая фикстура → exit 0 (03-01) | ✓ VERIFIED | `--all --root repo_ok` → exit 0 |
| 3  | `--json` — валидный JSON path/code/severity/line/message (03-01) | ✓ VERIFIED | json.load OK; все 5 ключей в каждом нарушении; top-keys checked/errors/warnings/violations |
| 4  | `--write-baseline` roundtrip → повторный прогон exit 0 (03-01) | ✓ VERIFIED | write → 0; filter → 0 (живой прогон) |
| 5  | `--strict` игнорирует baseline, эскалирует warning (03-01) | ✓ VERIFIED | baseline + `--strict` → exit 1 (живой прогон) |
| 6  | revit_compat.py существует, компилируется (03-02) | ✓ VERIFIED | 478 строк; test_compiles в составе 56 зелёных тестов |
| 7  | SUPPORTED_VERSIONS == (2020, 2022, 2024) (03-02) | ✓ VERIFIED | revit_compat.py:55 дословно |
| 8  | require_supported_version: TaskDialog + SystemExit (03-02) | ✓ VERIFIED | Код проинспектирован: TaskDialog.Show + raise SystemExit, русское сообщение; runtime → human item 1 |
| 9  | Units 2021+/Floor.Create 2023/ElementId Int64 2024/pythonnet закрыты (03-02) | ✓ VERIFIED | _units_map (DisplayUnitType↔UnitTypeId), create_floor, .Value/.IntegerValue, to_net_list/enum_from_int/iter_count — все ветки в файле |
| 10 | Контрактный ast-тест: шапка, версии, 13 функций API (03-02) | ✓ VERIFIED | Все 13 имён PUBLIC_API — top-level FunctionDef; нет голых except; нет BOM; docstring полный |
| 11 | MM008–MM012 ловятся на плохой фикстуре (03-03) | ✓ VERIFIED | JSON-прогон: MM008, MM009, MM010, MM011, MM012 присутствуют |
| 12 | MM014: неканонический бутстрап ловится, канонический чист (03-03) | ✓ VERIFIED | MM014 на плохой фикстуре; шаблон с каноническим блоком проходит `--strict` exit 0 |
| 13 | Хорошая фикстура по-прежнему без нарушений (03-03) | ✓ VERIFIED | exit 0 |
| 14 | AST-правила в режиме сырого .py (03-03) | ✓ VERIFIED | tmp-скрипт `import requests` → MM008; MM005/006/007 не выдаются |
| 15 | Белый список: stdlib + clr/System/Autodesk/pyrevit + lib + vendored (03-03) | ✓ VERIFIED | allowed_import_roots с sys.stdlib_module_names; тесты allows_host_and_stdlib/first_party зелёные |
| 16 | templates/ вне MM Lab.tab, pyRevit не грузит (03-04) | ✓ VERIFIED | Папка в корне репо, вне MM_LAB.extension (D-14) |
| 17 | Шаблон — минимальная рабочая кнопка с TODO (03-04) | ✓ VERIFIED | 86 строк: шапка-канон → бутстрап D-15 → revit_compat → Transaction Commit/RollBack+raise → TaskDialog; 8 TODO; EXTENSION_ROOT отсутствует |
| 18 | Шаблон `--strict` → exit 0 (03-04) | ✓ VERIFIED | Живой прогон: exit 0 |
| 19 | `--all --baseline` → exit 0 (03-04) | ✓ VERIFIED | Живой прогон: 17 юнитов, 0 ошибок / 0 предупреждений; без baseline → exit 1 (baseline фильтрует реальные нарушения) |
| 20 | В tab bundle.yaml нет орфана «ВОР» (03-04) | ✓ VERIFIED | grep пуст; layout: АРХИТЕКТУРА/ИОС/КООРДИНАЦИЯ/«-----» сохранены; diff к HEAD пуст |
| 21 | AGENTS.md — единственный полный текст конвенции (03-05) | ✓ VERIFIED | 323 строки ≥ 180; все 14 обязательных маркеров; severity всех 15 правил совпадает со словарём RULES чекера (программная сверка) |
| 22 | CLAUDE.md начинается с @AGENTS.md, только Claude-специфика (03-05) | ✓ VERIFIED | Строка 1 `@AGENTS.md`; GSD Release Map сохранён; graphify query/MMLabs_OBSIDIAN отсутствуют |
| 23 | GEMINI.md + .kilocode/rules/00-mmlab.md — тонкие указатели (03-05) | ✓ VERIFIED | GEMINI.md 4 строки (@AGENTS.md строка 1); 00-mmlab.md 3 строки со ссылкой на AGENTS.md |
| 24 | graphify + Obsidian перенесены в AGENTS.md (03-05) | ✓ VERIFIED | Маркеры graphify/Obsidian в AGENTS.md; из CLAUDE.md удалены |
| 25 | userEmail/currentDate в AGENTS.md отсутствуют (03-05) | ✓ VERIFIED | Обе подстроки отсутствуют; шаблон коммита нового формата «## Сессия»/«- Агент/модель:» |
| 26 | 7 канонических процедур mm-* в agents/commands/ (03-06) | ✓ VERIFIED | Ровно 7 файлов; у всех разделы Аргументы/Процедура/Финал (+ Гейты и запреты) |
| 27 | mm-adopt-script: полный поток с гейтами D-08..D-11 (03-06) | ✓ VERIFIED | 170 строк; РЕВЬЮ-ГЕЙТ (шаг 7) ДО регистрации (шаг 8); эвристика ИОС/АРХИТЕКТУРА/КООРДИНАЦИЯ; «не исполнять»; снятие записи baseline; живой прогон → human item 2 |
| 28 | mm-save-session: шаблон коммита, пофайловый стейджинг, push с подтверждением (03-06) | ✓ VERIFIED | 107 строк; «сессия:», «## Сессия», «Агент/модель», core.quotepath=false, git add пофайлово, push после «да» |
| 29 | mm-update-repo: только fetch + ff-only при чистом дереве (03-06) | ✓ VERIFIED | --ff-only, status --porcelain, reset --hard в списке запретов |
| 30 | Процедуры самодостаточны для 3 агентов (03-06) | ✓ VERIFIED | Оба пути quick task: /gsd-quick + ручные артефакты по образцу .planning/quick/260709-jko-* (образец существует) |
| 31 | 7 команд × 3 агента под единым именем /mm-<слаг> (03-07) | ✓ VERIFIED | 21 файл: 7×.claude/commands/*.md + 7×.gemini/commands/*.toml + 7×.kilo/commands/*.md |
| 32 | Тело адаптера 2–5 строк, процедуры не дублируются (03-07) | ✓ VERIFIED | Инспекция: тело — 2 строки «прочитай agents/commands/… и выполни»; $ARGUMENTS/{{args}} во всех; каждый ссылается ровно на СВОЙ канонический файл (21/21) |
| 33 | Имена Gemini-файлов плоские (03-07) | ✓ VERIFIED | mm-<слаг>.toml без подпапок |
| 34 | Каталожный unittest подтверждает согласованность (03-07) | ✓ VERIFIED | 7 тест-методов (canonical/claude/gemini/kilo/shell-injection/extra-files/agents-md-slugs) в зелёном прогоне |

**Score:** 34/34 truths verified (программный уровень)

### Required Artifacts

| Artifact | Expected | Status | Details |
|----------|----------|--------|---------|
| `tools/check_convention.py` | ≥250 строк, MM007, stdlib_module_names | ✓ VERIFIED | 923 строки; 15 правил MM000–MM014; без yaml/exec/eval; wired: импортируется тестами + baseline + процедурами |
| `tools/tests/test_check_convention.py` | ≥150 строк, write_baseline | ✓ VERIFIED | 592 строки; import check_convention (строка 39); в составе 56 зелёных тестов |
| `tools/tests/fixtures/repo_ok/**`, `repo_bad/**` | эталон + нарушитель | ✓ VERIFIED | Все 5 файлов repo_ok; BOM-гард repo_bad (EF BB BF); плохая фикстура парсится ast |
| `MM_LAB.extension/lib/revit_compat.py` | ≥200 строк, SUPPORTED_VERSIONS = (2020, 2022, 2024) | ✓ VERIFIED | 478 строк; 13 публичных функций; getattr(builtins, "__revit__") каскад; шапка-канон без BOM |
| `tools/tests/test_revit_compat_contract.py` | ≥50 строк, require_supported_version, ast.parse | ✓ VERIFIED | 136 строк; ast-контракт без import revit_compat |
| `templates/НоваяКнопка.pushbutton/` (3 файла) | ≥60 строк script, tooltip, ## Описание | ✓ VERIFIED | script.py 86 строк + import revit_compat; bundle.yaml title/tooltip ru+en_us + author; README все разделы |
| `tools/convention_baseline.json` | units | ✓ VERIFIED | generated/note/units (15 юнитов, 66 кодов) + pending_adoption (2 IFC-кнопки — доработка ревью); templates/ не попал; .vs-мусора нет |
| `MM_LAB.extension/MM Lab.tab/bundle.yaml` | без «ВОР» | ✓ VERIFIED | Чистый layout, рабочее дерево без диффа |
| `AGENTS.md` | ≥180 строк, ## Обязательные правила кода | ✓ VERIFIED | 323 строки; таблица правил = RULES чекера до severity |
| `CLAUDE.md`, `GEMINI.md`, `.kilocode/rules/00-mmlab.md` | @AGENTS.md / AGENTS.md | ✓ VERIFIED | Все три указателя; .claude/CLAUDE.md не изменён (git status пуст) |
| `agents/commands/mm-*.md` (7) | ≥80/≥40 строк, маркеры | ✓ VERIFIED | Все 7; adopt-script 170, save-session 107 строк; все контент-гейты планов пройдены |
| `.claude/.gemini/.kilo` адаптеры (21) | ссылки + prompt | ✓ VERIFIED | 21/21; без `!{`; без BOM |
| `tools/tests/test_mm_commands_catalog.py` | ≥40 строк, SLUGS | ✓ VERIFIED | 162 строки; SLUGS ×7; 7 тест-методов |

### Key Link Verification

| From | To | Via | Status | Details |
|------|----|-----|--------|---------|
| test_check_convention.py | check_convention.py | import после sys.path.insert | ✓ WIRED | строка 39 |
| check_convention.py | panel bundle.yaml | построчный парсер layout | ✓ WIRED | 32 вхождения layout; MM007 срабатывает (в кодах плохой фикстуры) |
| test_revit_compat_contract.py | revit_compat.py | ast.parse исходника | ✓ WIRED | ast.parse строка 62; import revit_compat в тесте отсутствует |
| revit_compat.py | builtins.__revit__ | каскад детекции версии | ✓ WIRED | строка 115 getattr(builtins, "__revit__", None) |
| templates/script.py | revit_compat.py | import после бутстрапа | ✓ WIRED | строка 38; вызовы require_supported_version + iter_count реально используются |
| convention_baseline.json | check_convention.py | --baseline при --all | ✓ WIRED | Живой прогон: с baseline exit 0, без — exit 1 (данные реально фильтруют) |
| CLAUDE.md | AGENTS.md | @import строка 1 | ✓ WIRED | `@AGENTS.md` первой строкой |
| AGENTS.md | check_convention.py / agents/commands/ | разделы Проверка/Команды | ✓ WIRED | 7 + 7 вхождений; слаги ×7 |
| mm-adopt-script.md | check_convention.py / .planning/quick/ | гейт --strict / quick task | ✓ WIRED | 3 вызова чекера + оба пути quick task; образец 260709-jko существует |
| mm-new-button.md | templates/НоваяКнопка.pushbutton | копирование скелета | ✓ WIRED | 3 вхождения |
| адаптеры (21) | agents/commands/mm-*.md | «прочитай и выполни» | ✓ WIRED | Каждый → свой канонический файл (программная сверка 7×3) |
| test_mm_commands_catalog.py | 28 файлов каталога | SLUGS + перекрёстные ссылки | ✓ WIRED | 7 методов в зелёном прогоне |

### Data-Flow Trace (Level 4)

Не веб-приложение; эквивалентная трассировка «данные реально текут»:

| Artifact | Данные | Источник | Реальные данные | Status |
|----------|--------|----------|-----------------|--------|
| check_convention.py --json | violations | обход реальных файлов фикстур | 13 кодов с path/line из реальных script.py | ✓ FLOWING |
| convention_baseline.json | units → фильтр | фактический аудит репо | с baseline 0/0, без — exit 1 (66 замороженных кодов реально вычитаются) | ✓ FLOWING |
| AGENTS.md таблица правил | severity | словарь RULES чекера | программная сверка 15/15 совпадений | ✓ FLOWING |
| revit_compat._units_map | карта единиц | версия Revit (кеш _VALIDATED_VERSION) | обе ветки (DisplayUnitType / UnitTypeId) в коде; runtime → human | ✓ FLOWING (код) |

### Behavioral Spot-Checks

| Behavior | Command | Result | Status |
|----------|---------|--------|--------|
| Полный тест-набор | `py -3 -m unittest discover -s tools/tests -q` | Ran 56 tests — OK | ✓ PASS |
| Фазовый инвариант | `py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json` | exit 0 (17 юнитов, 0/0) | ✓ PASS |
| Baseline не пустой | `--all` без baseline | exit 1 | ✓ PASS |
| Гейт шаблона | `"templates/НоваяКнопка.pushbutton" --strict` | exit 0 | ✓ PASS |
| Детекция всех правил | плохая фикстура `--json` | exit 1; MM001–MM012, MM014 | ✓ PASS |
| Чистый эталон | `--all --root repo_ok` | exit 0 | ✓ PASS |
| Сырой режим | tmp .py c `import requests` | MM001/MM002/MM004/MM008; без MM005–MM007 | ✓ PASS |
| Baseline roundtrip | write → filter → strict | 0 → 0 → 1 | ✓ PASS |
| Slash-команды /mm-* вживую | — | интерактивная сессия | ? SKIP → human |

Примечание: `__pycache__`, создаваемый py_compile-тестом в `MM_LAB.extension/lib/`, удалён после прогона (задокументированный pitfall фазы; состояние до прогона восстановлено).

### Requirements Coverage

REQUIREMENTS.md генерируется из «Карты релизов» и содержит только T-1xx/2xx/3xx (задачи плагинов) — CONV-* ID определены строкой Requirements фазы 3 в ROADMAP.md. Это ожидаемое устройство (03-RESEARCH §Phase Requirements, повторено во всех 7 SUMMARY); осиротевших ID нет — все 5 заявлены планами.

| Requirement | Source Plans | Description (ROADMAP «Объём работ») | Status | Evidence |
|-------------|--------------|--------------------------------------|--------|----------|
| CONV-STD | 03-02, 03-04, 03-05 | Конвенция: шапка, BuiltInParameter/GUID вместо lookupParameter, общее — в lib, без сторонних импортов, README | ✓ SATISFIED | AGENTS.md (18 правил, таблица MM000–MM014) + revit_compat (get_parameter/get_shared_parameter) + шаблон; runtime compat → human |
| CONV-CHECK | 03-01, 03-03, 03-04, 03-06, 03-07 | Проверка скрипта на соответствие конвенции | ✓ SATISFIED | Чекер 15 правил; 56 тестов; гейты --strict/--baseline зелёные; mm-check |
| CONV-REG | 03-01, 03-04, 03-06 | Регистрация скрипта в bundle.yaml для отображения на панели | ✓ SATISFIED | MM007 (layout + орфаны); орфан «ВОР» устранён; mm-adopt-script шаг 8 / mm-new-button регистрация в layout; видимость на панели → human |
| CONV-ADAPT | 03-06, 03-07 | Адаптация скрипта под конвенцию | ✓ SATISFIED (текстово) | mm-adopt-script: полный поток с ревью-гейтом ДО регистрации + 21 адаптер + каталожный тест; живой прогон → human |
| CONV-GSD | 03-06 | Создание новой задачи GSD при адаптации | ✓ SATISFIED (текстово) | mm-adopt-script шаг 9: /gsd-quick (Claude) + ручные артефакты .planning/quick/ (Gemini/Kilo); живые артефакты → human |

### Anti-Patterns Found

| File | Line | Pattern | Severity | Impact |
|------|------|---------|----------|--------|
| templates/НоваяКнопка.pushbutton/* | — | 8 меток TODO | ℹ️ Info | Интенциональный дизайн D-13 (места правки для человека и /mm-new-button); кнопка при этом рабочая и strict-зелёная — НЕ заглушка |
| .planning/ROADMAP.md | 32 | «Plans: 3/7 plans executed» при 7/7 отмеченных чекбоксах | ℹ️ Info | Устаревший счётчик в авто-генерируемом заголовке; не влияет на цель фазы (метаданные planning) |
| 03-04-SUMMARY.md | — | заявляет baseline «21 юнит / 202 нарушения»; фактически 15 юнитов / 66 кодов + pending_adoption | ℹ️ Info | Расхождение объясняется фиксами ревью (11db648..3f021e1): убраны .vs-мусорные юниты, IFC-кнопки вынесены в pending_adoption; улучшение поверх плана, задокументировано в 03-REVIEW-FIX (status: all_fixed, 9/9) |

Блокирующих анти-паттернов нет: ни одного FIXME/HACK/placeholder в tools/ и lib/; голых except нет; заглушек-компонентов нет.

### Code Review Integration

Пост-исполнительное ревью с авто-фиксами учтено: 2 итерации, 9 Warning-находок исправлено (WR-01…WR-09, коммиты в диапазоне 11db648..3f021e1), финальный статус all_fixed (0 Critical / 0 Warning). Верификация выполнена по коду ПОСЛЕ фиксов: pending_adoption в baseline и чекере (валидация схемы, строка 179 BASELINE_SECTIONS), кеш _VALIDATED_VERSION в revit_compat, +7 тестов (49 → 56). WR-06 помечен «requires human verification» — включён в human item 1.

### Human Verification Required

#### 1. Compat-хелперы в реальном Revit 2020 / 2022 / 2024

**Test:** Скопировать шаблон в панель, зарегистрировать в layout, pyRevit Reload; запустить на Revit 2020, 2022, 2024; при наличии — на иной версии.
**Expected:** Кнопка запускается, `require_supported_version` пропускает; отчёт «Стен в проекте: N»; на неподдерживаемой версии — TaskDialog «поддерживает 2020 / 2022 / 2024» и мягкий выход. Заодно подтверждает семантику фикса WR-06.
**Why human:** Revit API доступен только в среде Revit; headless-прогон невозможен (03-VALIDATION.md §Manual-Only).

#### 2. /mm-adopt-script на IFC_Двери — ревью-гейт (CONV-ADAPT)

**Test:** В Claude Code выполнить `/mm-adopt-script` на `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Двери.pushbutton` (untracked, намеренно оставлен как UAT-вход).
**Expected:** Чекер отработал, панель спрошена с обоснованием (D-10), сводка правок и diff показаны; без явного «да» регистрации в bundle.yaml НЕТ (D-08). После «да» — регистрация + снятие записи из pending_adoption/baseline.
**Why human:** Блокирующий гейт одобрения проверяется только живой интерактивной сессией.

#### 3. Quick task GSD после приёмки (CONV-GSD)

**Test:** После одобренной приёмки проверить артефакты.
**Expected:** Папка `.planning/quick/<id>-adopt-<слаг>/` с `<id>-PLAN.md` и `<id>-SUMMARY.md`; строка в таблице «Quick Tasks Completed» в `.planning/STATE.md`.
**Why human:** Артефакты создаются агентской командой интерактивно.

#### 4. Кнопка из шаблона видна на панели

**Test:** Скопировать `templates/НоваяКнопка.pushbutton` в `<Панель>.panel/`, добавить имя в layout panel-bundle.yaml, pyRevit Reload.
**Expected:** Кнопка появилась на панели MM LAB и запускается.
**Why human:** pyRevit парсит расширение на старте — требуется живой Reload.

#### 5. /mm-check --all и /mm-doctor вызываются как slash-команды

**Test:** В Claude Code выполнить `/mm-check --all`, затем `/mm-doctor`.
**Expected:** /mm-check пересказывает результат чекера по кнопкам и MM-кодам; /mm-doctor выдаёт read-only отчёт-таблицу «Проверка / Статус / Что делать».
**Why human:** Вызываемость slash-команды проверяется только интерактивно; Gemini CLI/Kilo Code не установлены — их адаптеры проверены форматно (это норма по 03-RESEARCH §Environment Availability).

### Gaps Summary

Гэпов нет. Все 34 must-have истины подтверждены на уровне кода живыми прогонами (56 тестов, все фазовые гейты чекера, программные сверки контента). Пять пунктов вынесены на человеческую проверку — все они заранее классифицированы фазой как Manual-Only (headless-среда без Revit и интерактивных агентских сессий) и не являются недоделками кода.

---

_Verified: 2026-07-24T13:30:00Z_
_Verifier: Claude (gsd-verifier)_
