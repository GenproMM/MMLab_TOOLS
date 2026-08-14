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
  warning: 2
  info: 10
  total: 12
status: issues_found
---

# Фаза 03: Отчёт code review (повторный, итерация 2)

**Проверено:** 2026-07-24
**Глубина:** standard
**Файлов проверено:** 42
**Статус:** issues_found

## Summary

Повторный обзор после фикс-итерации 1 (коммиты 11db648..cdecbd1,
WR-01..WR-07 предыдущего отчёта). Все семь исправлений проверены по коду
и эмпирически — **все семь корректны и полны**:

- **WR-01 — исправлено.** `load_baseline` (`tools/check_convention.py:676-692`)
  валидирует схему `units` ({путь: [коды]}). Проверено запуском: baseline
  `{"units": []}`, `{"units": {"a": "MM001"}}` и не-JSON дают чистый exit 2
  с русским сообщением в stderr, traceback отсутствует.
- **WR-02 — исправлено.** `iter_pushbuttons` (строки 658-664) отбрасывает
  кандидатов, у которых среди предков (относительно tab-каталога) есть
  мусорный каталог или другая папка `*.pushbutton`. Легитимные контейнеры
  pyRevit (`*.pulldown`, `*.stack`) фильтром не задеваются. Из
  `tools/convention_baseline.json` удалены все 4 фиктивных `.vs`-юнита.
- **WR-03 — исправлено (с остаточным зазором, см. IN-10).**
  `check_pushbutton` (строки 577-583) прогоняет MM000/MM003 и AST-правила
  по всем `*.py` верхнего уровня папки кнопки; правила шапки
  MM001/MM002/MM004 корректно остаются только для script.py
  (`header_rules=False`). Сообщения соседних модулей префиксуются именем
  файла. Скоуп задокументирован в докстрингах модуля и функции.
- **WR-04 — исправлено.** AGENTS.md правило 10 (строки 90-94) и
  `agents/commands/mm-adopt-script.md:66-67` приведены к каркасу шаблона
  (`Start()` перед `try`), с объяснением, почему `Start()` внутри `try`
  запрещён. Grep по всем адаптерам и докам: старой формулировки нигде
  не осталось; адаптеры .claude/.gemini/.kilo — тонкие ссылки, текст
  транзакций не дублируют.
- **WR-05 — исправлено.** `test_compiles`
  (`tools/tests/test_revit_compat_contract.py:123-132`) компилирует через
  `cfile=` во временный каталог. Проверено: после прогона suite каталог
  `MM_LAB.extension/lib/__pycache__/` не создаётся.
- **WR-06 — исправлено.** `require_supported_version` кеширует
  валидированную версию в `_VALIDATED_VERSION`
  (`MM_LAB.extension/lib/revit_compat.py:137-150`); `_units_map()` и
  `create_floor()` берут версию через `_effective_version()`
  (строки 153-162, 319, 391) — путь с явным аргументом `revit=` больше
  не теряется. Фолбэк на повторную детекцию до первого вызова
  `require_supported_version` задокументирован и соответствует контракту
  (require — первым в `main()`).
- **WR-07 — исправлено.** `--json --write-baseline` печатает ровно один
  JSON-объект (`tools/check_convention.py:832-837`). Проверено запуском:
  `{"baseline_written": "...", "violations": 190}`, exit 0.

Регрессий фиксы не внесли: полный тест-набор проходит (49 тестов, OK);
`py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json`
даёт exit 0 (17 юнитов — 15 прежних + 2 новых IFC-кнопки из рабочего
дерева, 4 фиктивных `.vs`-юнита корректно исчезли).

Два новых предупреждения: baseline «грандфазерит» две НОВЫЕ (не legacy,
не закоммиченные) кнопки вопреки задокументированной политике
храповика (WR-08), и ни одно из четырёх поведенческих исправлений
чекера не покрыто регрессионным тестом — набор остался в 49 тестов
(WR-09). Все девять прежних Info-находок актуальны (перепроверены по
текущему коду), плюс одна новая (остаточный зазор WR-03).

## Warnings

### WR-08: Baseline грандфазерит новые незакоммиченные IFC-кнопки вопреки политике «новые нарушения baseline не покрывает»

**File:** `tools/convention_baseline.json:5-28`
**Issue:** Записи `АРХИТЕКТУРА.panel/IFC_Двери.pushbutton` и
`IFC_Окна.pushbutton` содержат по 10 кодов, включая error-правила
MM005 (нет bundle.yaml), MM006 (нет README.md), MM007 (не в layout
панели). Эти кнопки — не legacy: их папки отсутствуют в git (untracked,
только script.py + png в рабочем дереве). AGENTS.md:240-242 обещает:
«старые кнопки не "краснят" общий прогон, но и не растят долг — новые
нарушения baseline не покрывает». Закоммиченный baseline с записями под
локальные незакоммиченные папки (а) фиксирует машинно-локальное
состояние в общем файле, (б) заранее маскирует все error-нарушения этих
кнопок, когда их закоммитят: `/mm-check --all` никогда не покажет
отсутствие bundle.yaml/README/регистрации. Состояние существовало ещё до
фикс-итерации (коммит effab8b), но при регенерации baseline в 96480cb
записи сохранены.
**Fix:** Удалить обе записи `IFC_*` из `tools/convention_baseline.json`.
Если IFC-кнопки — незавершённая работа другой сессии, довести их до
конвенции (`/mm-adopt-script`) или не включать их допуски в общий
baseline до коммита самих кнопок.

### WR-09: Ни одно из четырёх поведенческих исправлений чекера не закреплено регрессионным тестом

**File:** `tools/tests/test_check_convention.py` (набор: 49 тестов до и после фиксов)
**Issue:** Фиксы WR-01, WR-02, WR-03 и WR-07 изменили поведение
`tools/check_convention.py`, но тест-модуль, объявленный «исполняемой
спецификацией» чекера (строки 4-7), не получил ни одного нового теста:
- битый baseline (`{"units": []}`, `{"units": {"a": "MM001"}}`) → exit 2
  без traceback — не проверяется (`load_baseline` тестируется только на
  валидном roundtrip, строки 459-472);
- пропуск `*.pushbutton` внутри `.vs/`/`__pycache__`/другой кнопки в
  `iter_pushbuttons` — не проверяется;
- AST-правила на соседних `*.py` папки кнопки (MM008 в helpers.py) и
  пропуск для них MM001/MM002/MM004 — не проверяется;
- `--json --write-baseline` печатает JSON-объект — не проверяется
  (`test_baseline_roundtrip` гоняет запись без `--json`).
Любой будущий рефакторинг чекера может молча вернуть все четыре бага.
**Fix:** Добавить по тесту на каждый случай, например:
```python
def test_baseline_units_not_dict_exit_2(self):
    path = Path(self._tmp()) / "bad.json"
    path.write_text('{"units": []}', encoding="utf-8")
    code, _out = run_main(["--all", "--root", str(REPO_OK),
                           "--baseline", str(path)])
    self.assertEqual(code, 2)

def test_sibling_module_ast_rules(self):
    root, button = copy_button_to_tmp(self, with_panel=False)
    (button / "helpers.py").write_text("import requests\n", encoding="utf-8")
    codes = {v.code for v in check_convention.check_pushbutton(button, root)}
    self.assertIn("MM008", codes)
    self.assertNotIn("MM001", codes)  # шапка соседей не требуется
```
и аналогичные для `iter_pushbuttons` (вложенный `.vs/X.pushbutton` в
tmp-репо) и `--json --write-baseline` (stdout парсится `json.loads`).

## Info

### IN-01: iter_count молча потребляет одноразовые Python-итераторы

**File:** `MM_LAB.extension/lib/revit_compat.py:437-457`
**Issue:** (Актуально, без изменений с прошлого обзора.) Финальный фолбэк
считает элементы итерацией: для генератора без `len`/`.Count`
последовательность вызывающего исчерпывается — `n = iter_count(gen);
for x in gen:` не итерирует ничего. .NET-коллекции безопасны.
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

### IN-05: Мёртвая ветка тест-хелпера — copy_button_to_tmp(with_panel=True) не используется

**File:** `tools/tests/test_check_convention.py:67-84`
**Issue:** (Актуально.) Все три вызова (строки 189, 209, 422) передают
`with_panel=False`; ветка `True` мертва.
**Fix:** Либо добавить тест с `with_panel=True` (регистрация в
tmp-панели — заодно закроет часть WR-09), либо убрать параметр.

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

**File:** `tools/check_convention.py:577-583` (также докстринг модуля, строки 11-13)
**Issue:** Обход соседних модулей нерекурсивный (`glob("*.py")`).
Файл `helpers/util.py` в подкаталоге кнопки не проходит ни AST-правила,
ни MM013 (мусорными считаются только `__pycache__`/`.vs`): сторонний
импорт всё ещё можно спрятать уровнем ниже. Докстринг «проверяются ВСЕ
*.py в папке кнопки» двусмыслен относительно вложенности. Реализация
соответствует предложенному в прошлом обзоре фиксу (glob верхнего
уровня), поэтому severity — Info, а не эскалация WR-03.
**Fix:** Либо перейти на `rglob("*.py")` с пропуском `JUNK_DIR_NAMES`,
либо явно указать в докстринге и AGENTS.md «только верхний уровень папки
кнопки».

---

_Reviewed: 2026-07-24_
_Reviewer: Claude (gsd-code-reviewer)_
_Depth: standard_
_Iteration: 2 (после фикс-итерации 1, коммиты 11db648..cdecbd1)_
