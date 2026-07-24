---
phase: 03-convention
fixed_at: 2026-07-24T00:00:00Z
review_path: .planning/phases/03-convention/03-REVIEW.md
iteration: 1
findings_in_scope: 7
fixed: 7
skipped: 0
status: all_fixed
---

# Phase 03: Code Review Fix Report

**Fixed at:** 2026-07-24
**Source review:** .planning/phases/03-convention/03-REVIEW.md
**Iteration:** 1

**Summary:**
- Findings in scope: 7 (Warning; Critical — 0, Info вне scope)
- Fixed: 7
- Skipped: 0

Верификация после каждого исправления и финально: `py -3 -m unittest discover
-s tools/tests -q` — 49 тестов OK; `py -3 tools/check_convention.py --all
--baseline tools/convention_baseline.json` — exit 0;
`py -3 tools/check_convention.py "templates/НоваяКнопка.pushbutton" --strict` — exit 0.

## Fixed Issues

### WR-01: Битый baseline давал raw traceback вместо чистого exit 2

**Files modified:** `tools/check_convention.py`
**Commit:** 11db648
**Applied fix:** `load_baseline` валидирует схему `units` (объект
`{путь: [коды]}`); битый baseline даёт `ValueError`, который `main()`
превращает в русское сообщение и exit 2. Проверено эмпирически на
`{"units": []}` и `{"units": {"a": "MM001"}}` — оба дают exit 2.

### WR-02: iter_pushbuttons считал кнопками `*.pushbutton` внутри мусорных папок

**Files modified:** `tools/check_convention.py`, `tools/convention_baseline.json`
**Commit:** 96480cb
**Applied fix:** Кандидаты с `.vs`/`__pycache__` или другой папкой
`*.pushbutton` среди предков (относительно tab-каталога) пропускаются.
Baseline регенерирован `--write-baseline` в том же коммите: удалены ровно
4 фиктивных `.vs`-юнита. Записи `IFC_Двери`/`IFC_Окна` из HEAD сохранены
вручную — эти кнопки существуют только как неотслеживаемые папки в рабочем
дереве и отсутствовали в checkout worktree; их удаление сломало бы гейт
основного дерева (итог: 21 → 17 юнитов).

### WR-03: AST-правила применялись только к script.py

**Files modified:** `tools/check_convention.py`
**Commit:** 5a5bfde
**Applied fix:** `check_pushbutton` проверяет все `*.py` папки кнопки:
script.py — полностью; соседние модули — MM000/MM003 + AST-правила
(MM008–MM012, MM014) через новый параметр `header_rules=False`
(шапка MM001/MM002/MM004 по таблице RULES обязательна только для
script.py — вариант, явно допущенный ревью). Сообщения соседних модулей
префиксуются именем файла. Поведение задокументировано в docstring модуля.
Эмпирически: `helpers.py` с `import requests` рядом со script.py теперь
даёт MM008 без ложного MM001. Полный гейт с baseline остался exit 0 —
у легаси-кнопок нет соседних `*.py` с нарушениями, регенерация baseline
не потребовалась.

### WR-04: Правило 10 AGENTS.md противоречило каркасу транзакций шаблона

**Files modified:** `AGENTS.md`, `agents/commands/mm-adopt-script.md`
**Commit:** 0d2cbe7
**Applied fix:** Формулировки приведены к каноническому шаблону
(`templates/НоваяКнопка.pushbutton/script.py`, он корректен и не менялся):
`Start()` — ПЕРЕД `try`; `Commit()` — в `try`; в `except` —
`RollBack()` и `raise`. В AGENTS.md добавлено объяснение, почему `Start()`
внутри `try` запрещён (RollBack у незапущенной транзакции маскирует
исходную ошибку). Дубли формулировки в прочих файлах не найдены
(упоминания в `.planning/` и «Регламенте» — вне scope ревью).

### WR-05: Контрактный тест писал bytecode-кеш в живое расширение

**Files modified:** `tools/tests/test_revit_compat_contract.py`
**Commit:** 58219c3
**Applied fix:** `test_compiles` компилирует через
`py_compile.compile(..., cfile=...)` во временный каталог
(`tempfile.TemporaryDirectory`). Эмпирически: после прогона suite каталог
`MM LAB.extension/lib/__pycache__/` не создаётся — `/mm-doctor` снова
действительно read-only.

### WR-06: revit_compat повторно детектировал версию вместо валидированной

**Files modified:** `MM LAB.extension/lib/revit_compat.py`
**Commit:** bf1614d
**Applied fix:** Добавлен модульный кеш `_VALIDATED_VERSION`
(заполняется в `require_supported_version` при успехе) и приватный хелпер
`_effective_version()` (кеш → повторная детекция). `_units_map()` и
`create_floor()` берут версию через него — путь с явным аргументом
`revit=` больше не теряется. Контрактный тест публичного API (13 функций,
запрет голых except) проходит.
**Статус:** fixed: requires human verification — семантика проверяема
только в живом Revit (2020-хост без `__revit__`/`HOST_APP`); статические
проверки и контрактный тест зелёные.

### WR-07: `--json --write-baseline` давал пустой stdout

**Files modified:** `tools/check_convention.py`
**Commit:** cdecbd1
**Applied fix:** В JSON-режиме печатается один JSON-объект
`{"baseline_written": <путь>, "violations": <N>}` (ensure_ascii=False) —
контракт «ровно один JSON-объект в stdout» соблюдается и при записи
baseline. Проверено эмпирически: вывод парсится `json.load`.

---

_Fixed: 2026-07-24_
_Fixer: Claude (gsd-code-fixer)_
_Iteration: 1_
