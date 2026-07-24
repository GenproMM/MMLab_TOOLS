---
phase: 03-convention
fixed_at: 2026-07-24T00:00:00Z
review_path: .planning/phases/03-convention/03-REVIEW.md
iteration: 2
findings_in_scope: 2
fixed: 2
skipped: 0
status: all_fixed
---

# Phase 03: Code Review Fix Report

**Fixed at:** 2026-07-24
**Source review:** .planning/phases/03-convention/03-REVIEW.md (итерация 2)
**Iteration:** 2

**Summary (итерация 2):**
- Findings in scope: 2 (Warning WR-08, WR-09; Critical — 0, Info вне scope)
- Fixed: 2
- Skipped: 0

**Кумулятивно по фазе (итерации 1–2):** 9 Warning-находок, 9 исправлено,
0 пропущено (WR-06 — «fixed: requires human verification», семантика
проверяема только в живом Revit).

Верификация после каждого исправления и финально в ОСНОВНОМ рабочем
дереве (с неотслеживаемыми IFC-кнопками):
`py -3 -m unittest discover -s tools/tests -q` — 56 тестов OK
(49 → 56: +3 теста контракта pending_adoption, +4 регрессионных);
`py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json`
— exit 0 (17 юнитов, 0 нарушений);
`py -3 tools/check_convention.py "templates/НоваяКнопка.pushbutton" --strict`
— exit 0.

## Fixed Issues (итерация 2)

### WR-08: Baseline грандфазерит новые незакоммиченные IFC-кнопки вопреки политике

**Files modified:** `tools/check_convention.py`, `tools/convention_baseline.json`,
`tools/tests/test_check_convention.py`, `AGENTS.md`, `agents/commands/mm-adopt-script.md`
**Commit:** f14bb89
**Applied fix:** Выбран вариант (а) из ревью — **выделенная секция
`pending_adoption`** в baseline, поскольку простое удаление записей
(вариант из **Fix** ревью) ломает обязательный гейт
`--all --baseline` в основном дереве, где IFC-кнопки физически
присутствуют (untracked, ожидают `/mm-adopt-script`), а автопропуск
незарегистрированных кнопок (вариант (б)) выхолостил бы MM007.
Что сделано:
- записи `IFC_Двери`/`IFC_Окна` перенесены из `units` в новую секцию
  `pending_adoption` — грандфазеринг legacy и временные допуски
  WIP-кнопок теперь машинно различимы;
- `load_baseline` валидирует обе секции по одной схеме `{путь: [коды]}`;
  `apply_baseline` учитывает обе;
- `write_baseline` СОХРАНЯЕТ `pending_adoption` существующего файла и
  исключает её пути из генерируемых `units` — регенерация baseline больше
  не превращает временный допуск в вечный (ровно та утечка, что произошла
  в 96480cb);
- политика зафиксирована: docstring чекера, `BASELINE_NOTE`, AGENTS.md
  (раздел Baseline: `pending_adoption` — не грандфазеринг, при приёмке
  кнопка обязана пройти `--strict`, который baseline игнорирует, и запись
  удаляется в том же коммите), шаг 8 `agents/commands/mm-adopt-script.md`
  (удалять запись из `units` ИЛИ `pending_adoption`);
- новый контракт закреплён 3 тестами (`PendingAdoptionTests`).
Интент политики соблюдён: допуски IFC-кнопок объявлены временными явно,
гейт приёмки они не ослабляют, а гейт основного дерева остаётся зелёным
(проверено эмпирически: exit 0, 17 юнитов).

### WR-09: Поведенческие исправления чекера не закреплены регрессионными тестами

**Files modified:** `tools/tests/test_check_convention.py`
**Commit:** 3f021e1
**Applied fix:** Новый класс `CheckerRegressionTests` — 4 теста, пиновые
для фиксов итерации 1:
- WR-01: битый baseline (`{"units": []}`, `{"units": {"a": "MM001"}}`,
  не-JSON) → exit 2, русское сообщение в stderr, без traceback
  (stderr перехватывается `contextlib.redirect_stderr`);
- WR-02: `iter_pushbuttons` пропускает `*.pushbutton` внутри `.vs/`,
  `__pycache__/` и внутри другой кнопки (использована живая ветка
  `copy_button_to_tmp(with_panel=True)` — попутно закрыт мёртвый код
  из IN-05);
- WR-03: сторонний импорт в соседнем `helpers.py` даёт MM008 с префиксом
  имени файла, правила шапки MM001/MM002/MM004 на соседний модуль
  не распространяются;
- WR-07: `--json --write-baseline` печатает ровно один JSON-объект
  (`json.loads` stdout, ключи `baseline_written`/`violations`).
Стиль — stdlib unittest, как в остальном модуле. Набор: 49 → 56 тестов, OK.

## Fixed Issues (итерация 1, для полноты)

Подробности — в git-истории и отчёте ревью итерации 2 (все семь фиксов
перепроверены ревьюером и признаны корректными):

- **WR-01** — `load_baseline` валидирует схему; битый baseline → чистый
  exit 2. Commit 11db648.
- **WR-02** — `iter_pushbuttons` фильтрует кандидатов в мусорных/вложенных
  папках; из baseline удалены 4 фиктивных `.vs`-юнита. Commit 96480cb.
- **WR-03** — AST-правила на всех `*.py` верхнего уровня папки кнопки;
  шапка — только для script.py. Commit 5a5bfde.
- **WR-04** — правило 10 AGENTS.md и mm-adopt-script.md приведены
  к каркасу шаблона (`Start()` перед `try`). Commit 0d2cbe7.
- **WR-05** — `test_compiles` пишет bytecode во временный каталог
  (`cfile=`). Commit 58219c3.
- **WR-06** — кеш `_VALIDATED_VERSION` + `_effective_version()` в
  revit_compat. Commit bf1614d. **Статус: fixed: requires human
  verification** (семантика проверяема только в живом Revit).
- **WR-07** — `--json --write-baseline` печатает JSON-объект статуса.
  Commit cdecbd1.

## Skipped Issues

Нет — обе находки итерации 2 исправлены; в итерации 1 пропусков не было.

---

_Fixed: 2026-07-24_
_Fixer: Claude (gsd-code-fixer)_
_Iteration: 2_
