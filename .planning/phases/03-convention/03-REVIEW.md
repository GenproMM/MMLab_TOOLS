---
phase: 03-convention
reviewed: 2026-07-24T00:00:00Z
depth: standard
files_reviewed: 42
files_reviewed_list:
  - tools/check_convention.py
  - tools/tests/test_check_convention.py
  - tools/tests/test_revit_compat_contract.py
  - tools/tests/test_mm_commands_catalog.py
  - tools/convention_baseline.json
  - MM_LAB.extension/lib/revit_compat.py
  - MM_LAB.extension/MM Lab.tab/bundle.yaml
  - templates/НоваяКнопка.pushbutton/script.py
  - templates/НоваяКнопка.pushbutton/bundle.yaml
  - templates/НоваяКнопка.pushbutton/README.md
  - AGENTS.md
  - CLAUDE.md
  - GEMINI.md
  - .kilocode/rules/00-mmlab.md
  - agents/commands/mm-adopt-script.md
  - agents/commands/mm-check.md
  - agents/commands/mm-doctor.md
  - agents/commands/mm-new-button.md
  - agents/commands/mm-new-compat.md
  - agents/commands/mm-save-session.md
  - agents/commands/mm-update-repo.md
  - .claude/commands/mm-adopt-script.md
  - .claude/commands/mm-check.md
  - .claude/commands/mm-doctor.md
  - .claude/commands/mm-new-button.md
  - .claude/commands/mm-new-compat.md
  - .claude/commands/mm-save-session.md
  - .claude/commands/mm-update-repo.md
  - .gemini/commands/mm-adopt-script.toml
  - .gemini/commands/mm-check.toml
  - .gemini/commands/mm-doctor.toml
  - .gemini/commands/mm-new-button.toml
  - .gemini/commands/mm-new-compat.toml
  - .gemini/commands/mm-save-session.toml
  - .gemini/commands/mm-update-repo.toml
  - .kilo/commands/mm-adopt-script.md
  - .kilo/commands/mm-check.md
  - .kilo/commands/mm-doctor.md
  - .kilo/commands/mm-new-button.md
  - .kilo/commands/mm-new-compat.md
  - .kilo/commands/mm-save-session.md
  - .kilo/commands/mm-update-repo.md
findings:
  critical: 0
  warning: 0
  info: 11
  total: 11
status: clean
---

# Фаза 03: Отчёт code review (финальный, итерация 3)

**Проверено:** 2026-07-24
**Глубина:** standard
**Файлов проверено:** 42
**Статус:** clean (Critical: 0, Warning: 0; остаточные Info не блокируют)

## Summary

Финальный обзор после фикс-итерации 2 (коммиты f14bb89 — WR-08,
3f021e1 — WR-09). Оба исправления проверены по диффам и эмпирически —
**оба корректны и полны, регрессий нет**:

- **WR-08 — исправлено.** Введена секция `pending_adoption` в baseline
  (`tools/convention_baseline.json:102-127`): обе записи `IFC_Двери` /
  `IFC_Окна` перенесены из `units` в `pending_adoption` — грандфазеринг
  legacy-кнопок и временные допуски ещё не принятых кнопок теперь
  разделены и задокументированы (AGENTS.md:240-252, докстринг модуля
  `tools/check_convention.py:28-37`, процедура
  `agents/commands/mm-adopt-script.md:100-103`). Механика проверена:
  - `load_baseline` валидирует обе секции по общей схеме
    `{путь: [коды]}` через `BASELINE_SECTIONS` (строки 686-703);
  - `apply_baseline` учитывает обе секции (строки 706-722);
  - `write_baseline` сохраняет `pending_adoption` существующего файла
    и НЕ переносит его пути в `units` (строки 725-759) — временный
    допуск не превращается в грандфазеринг. Проверено запуском:
    регенерация baseline в scratch-копию идемпотентна (pending
    идентичен, IFC-путей в `units` нет, ключи `units` без изменений);
  - гейт приёмки не ослаблен: `--strict` игнорирует baseline целиком
    (`main`, строка 883) — проверено запуском: IFC-кнопка под
    `--strict --baseline` даёт exit 1 (6 ошибок MM005/MM006/MM007...).
  Ключевое возражение WR-08 («baseline заранее маскирует
  error-нарушения новых кнопок вопреки политике») снято: маскировка
  стала явной, временной и задокументированной политикой с
  сохранённым strict-гейтом приёмки, а `--write-baseline` не
  «отмывает» допуски в постоянные.
- **WR-09 — исправлено.** Добавлен класс `CheckerRegressionTests`
  (`tools/tests/test_check_convention.py:475-537`) — все четыре
  поведенческих фикса итерации 1 закреплены тестами:
  - `test_malformed_baseline_clean_exit_2` — три битых baseline →
    exit 2, русское сообщение в stderr, без traceback (WR-01);
  - `test_iter_pushbuttons_skips_junk_and_nested` — `*.pushbutton`
    внутри `.vs/`, `__pycache__/` и другой кнопки не считаются
    кнопками (WR-02); заодно оживлена ветка
    `copy_button_to_tmp(with_panel=True)` — IN-05 прошлого отчёта
    закрыт попутно;
  - `test_sibling_module_ast_rules` — MM008 ловится в `helpers.py`
    с префиксом имени файла, правила шапки MM001/MM002/MM004 на
    соседний модуль не распространяются (WR-03);
  - `test_write_baseline_json_prints_json_object` — `--json
    --write-baseline` печатает ровно один JSON-объект (WR-07).
  Плюс `PendingAdoptionTests` (строки 540-588, 3 теста) закрепляют
  контракт WR-08: фильтрация по pending, валидация схемы, сохранение
  секции при регенерации без переноса в units.

Регрессий нет: полный набор — **56 тестов, OK** (49 + 4 регрессионных
+ 3 pending_adoption; заявленное число совпадает);
`py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json`
— «Проверено: 17, ошибок: 0, предупреждений: 0», exit 0. Адаптеры
.claude/.gemini/.kilo — тонкие ссылки на канонические процедуры,
дрейфа после правки `agents/commands/mm-adopt-script.md` нет
(текст про `units`/`pending_adoption` живёт только в каноне).

Critical и Warning находок не осталось. Ниже — 11 Info: 9 актуальных
из прошлых итераций (IN-05 закрыт) и 2 новых мелких (IN-11, IN-12) по
краевым случаям кода WR-08; оба не эскалируются, потому что защитный
strict-гейт в этих сценариях не затрагивается (обоснование в тексте).

## Info

### IN-01: iter_count молча потребляет одноразовые Python-итераторы

**File:** `MM_LAB.extension/lib/revit_compat.py:437-457`
**Issue:** (Актуально, без изменений.) Финальный фолбэк считает элементы
итерацией: для генератора без `len`/`.Count` последовательность
вызывающего исчерпывается — `n = iter_count(gen); for x in gen:` не
итерирует ничего. .NET-коллекции безопасны.
**Fix:** Задокументировать ограничение в докстринге («не передавай
одноразовые итераторы/генераторы») или материализовать:
`return len(list(sequence))`.

### IN-02: MM014 даёт несколько дублирующихся нарушений на один legacy-блок

**File:** `tools/check_convention.py:399-421`
**Issue:** (Актуально.) Каждое `ast.Name` с `EXTENSION_ROOT`, каждый
неканонический `sys.path`-вызов и каждый 4-hop `join` — отдельная строка
MM014; один логический дефект даёт несколько строк в выводе `/mm-check`.
**Fix:** Дедуплицировать MM014 по файлу (оставить первую строку) перед
возвратом из `check_ast_rules`.

### IN-03: MM007 — регистрация и орфаны непоследовательно терпят записи layout с суффиксом

**File:** `tools/check_convention.py:273-286` vs `503-532`
**Issue:** (Актуально.) `_entry_has_folder` принимает запись с суффиксом
(`child.name == entry`), а `_check_layout_registration` сравнивает только
имя без суффикса: запись `- Кнопка.pushbutton` даёт «не зарегистрирована»
без соответствующей орфан-подсказки.
**Fix:** В `_check_layout_registration` дополнительно сверять
`button_dir.name` (с суффиксом) с `layout_names` и в сообщении указывать,
что запись должна быть без суффикса.

### IN-04: Белый список first-party импортов игнорирует пакеты-каталоги в lib/

**File:** `tools/check_convention.py:291-312`
**Issue:** (Актуально.) `lib_dir.glob("*.py")` собирает только модули.
Будущий пакет `MM_LAB.extension/lib/helpers/__init__.py` сделает
`import helpers` ложным MM008.
**Fix:** Добавлять также стемы подкаталогов с `__init__.py`.

### IN-06: test_no_extra_mm_files не сканирует agents/commands/

**File:** `tools/tests/test_mm_commands_catalog.py:133-151`
**Issue:** (Актуально.) Гард двойников покрывает только три каталога
адаптеров; опечатка в каноническом файле (`agents/commands/mm-chek.md`)
пройдёт незамеченной.
**Fix:** Добавить `CANONICAL_DIR: {slug + ".md" for slug in SLUGS}` в
словарь `expected`.

### IN-07: Gemini-адаптеры проверяются подстрокой, а не парсингом TOML

**File:** `tools/tests/test_mm_commands_catalog.py:94-108`
**Issue:** (Актуально.) Проверки `description =` / `prompt =` проходят и
на синтаксически битом TOML. На Python 3.11+ `tomllib` доступен бесплатно.
**Fix:** `try: import tomllib` и при наличии — `tomllib.loads(text)` с
проверкой ключей `description`/`prompt`.

### IN-08: Висячий разделитель в конце layout вкладки

**File:** `MM_LAB.extension/MM Lab.tab/bundle.yaml:7`
**Issue:** (Актуально.) `- -----` — последняя запись `layout:`,
вертикальный разделитель, после которого ничего нет.
**Fix:** Удалить хвостовую строку `- -----`.

### IN-09: Секция GSD в CLAUDE.md использует python3 при репо-стандарте py -3

**File:** `CLAUDE.md:12-17`
**Issue:** (Актуально.) Все вызовы чекера/тестов в AGENTS.md и процедурах
используют `py -3`; секция Release Map предписывает
`python3 RELEASE_MAP/gsd_release_sync.py ...`, что на Windows без алиаса
обычно не резолвится.
**Fix:** Использовать `py -3` (или задокументировать требуемый алиас).

### IN-10: Остаточный зазор WR-03 — *.py в подкаталогах папки кнопки по-прежнему вне проверки

**File:** `tools/check_convention.py:587-593` (также докстринг модуля)
**Issue:** (Актуально.) Обход соседних модулей нерекурсивный
(`glob("*.py")`). Файл `helpers/util.py` в подкаталоге кнопки не проходит
ни AST-правила, ни MM013 (мусорными считаются только
`__pycache__`/`.vs`): сторонний импорт всё ещё можно спрятать уровнем
ниже. Реализация соответствует предложенному в ревью фиксу (glob
верхнего уровня), поэтому severity — Info.
**Fix:** Либо перейти на `rglob("*.py")` с пропуском `JUNK_DIR_NAMES`,
либо явно указать в докстринге и AGENTS.md «только верхний уровень папки
кнопки».

### IN-11: write_baseline молча теряет pending_adoption при битом существующем baseline

**File:** `tools/check_convention.py:733-740`
**Issue:** (Новое, код WR-08.) `write_baseline` глотает и `OSError`
(нет файла — легитимный первый запуск), и `ValueError` (существующий
файл битый) одним `except (OSError, ValueError): existing = {}`.
Проверено запуском: `--write-baseline` поверх синтаксически битого
baseline завершается exit 0 без предупреждения, секция
`pending_adoption` исчезает, а нарушения её кнопок попадают в `units` —
временный допуск молча превращается в грандфазеринг. Не эскалируется
до Warning: сценарий требует уже испорченного файла (путь чтения
`--baseline` падает на нём громко с exit 2 — WR-01), фильтрующее
поведение обеих секций идентично, strict-гейт приёмки baseline
игнорирует в любом случае, а результат регенерации — коммитуемый файл,
перенос виден в diff.
**Fix:** Разделить ветки: `FileNotFoundError` → `existing = {}`
(первый запуск), прочие `OSError`/`ValueError` → громкий exit 2
(«существующий baseline битый — почини или удали перед регенерацией»)
либо хотя бы предупреждение в stderr о сбросе `pending_adoption`.

### IN-12: Help-строка --baseline не обновлена под pending_adoption

**File:** `tools/check_convention.py:788-790`
**Issue:** (Новое, косметика.) `--baseline` в argparse-описании всё ещё
«JSON с допущенными нарушениями legacy-кнопок», хотя докстринг модуля,
AGENTS.md и BASELINE_NOTE уже описывают две секции
(`units`/`pending_adoption`). `--help` вводит в заблуждение о составе
файла.
**Fix:** «JSON с допущенными нарушениями: units (legacy) и
pending_adoption (ещё не принятые кнопки)».

---

_Reviewed: 2026-07-24_
_Reviewer: Claude (gsd-code-reviewer)_
_Depth: standard_
_Iteration: 3 (финальная, после фикс-итерации 2: f14bb89, 3f021e1)_
