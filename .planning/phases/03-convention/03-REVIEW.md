---
phase: 03-convention
reviewed: 2026-07-24T00:00:00Z
depth: standard
files_reviewed: 50
files_reviewed_list:
  - tools/check_convention.py
  - tools/tests/test_check_convention.py
  - tools/tests/test_revit_compat_contract.py
  - tools/tests/test_mm_commands_catalog.py
  - tools/convention_baseline.json
  - MM LAB.extension/lib/revit_compat.py
  - MM LAB.extension/MM Lab.tab/bundle.yaml
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
  - tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/bundle.yaml
  - tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/Тестовая панель.panel/bundle.yaml
  - tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/Тестовая панель.panel/Хорошая кнопка.pushbutton/script.py
  - tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/Тестовая панель.panel/Хорошая кнопка.pushbutton/bundle.yaml
  - tools/tests/fixtures/repo_ok/MM LAB.extension/MM Lab.tab/Тестовая панель.panel/Хорошая кнопка.pushbutton/README.md
  - tools/tests/fixtures/repo_bad/MM LAB.extension/MM Lab.tab/bundle.yaml
  - tools/tests/fixtures/repo_bad/MM LAB.extension/MM Lab.tab/Плохая панель.panel/bundle.yaml
  - tools/tests/fixtures/repo_bad/MM LAB.extension/MM Lab.tab/Плохая панель.panel/Плохая кнопка.pushbutton/script.py
findings:
  critical: 0
  warning: 7
  info: 9
  total: 16
status: issues_found
---

# Phase 03: Code Review Report

**Reviewed:** 2026-07-24
**Depth:** standard
**Files Reviewed:** 50
**Status:** issues_found

## Summary

Reviewed the convention-checker CLI (`tools/check_convention.py`), its test
suite and fixtures, the `revit_compat` module and its contract test, the
button template, the AGENTS.md standard, per-agent pointer files, and the
7×3 command adapter catalog.

Verification performed during review: the full test suite passes
(49 tests, OK) and `py -3 tools/check_convention.py --all --baseline
tools/convention_baseline.json` returns exit 0 (21 units checked). The
checker is genuinely static (only `open(..., "rb")` / `ast.parse`; no
import/exec of checked code), the adapters contain no shell inserts, and
the RULES dict severities match the AGENTS.md table exactly.

No critical (blocker) findings. Seven warnings were found — two of them
confirmed empirically during this review (malformed-baseline crash with a
raw traceback; bytecode cache written into the live pyRevit extension
directory by a "read-only" contract test), plus one real contradiction
between AGENTS.md rule 10 and the canonical template's transaction
skeleton, and a coverage hole where only `script.py` is AST-checked.

## Warnings

### WR-01: Malformed baseline crashes with raw traceback instead of clean exit 2

**File:** `tools/check_convention.py:655-663` (also `load_baseline`, 646-652)
**Issue:** `load_baseline` only validates that the top-level JSON is a dict.
If `units` is not a dict (e.g. `{"units": []}`), `apply_baseline` crashes
with an uncaught `AttributeError: 'list' object has no attribute 'items'` —
confirmed by running the CLI. If a `codes` value is a string instead of a
list, `for code in codes` silently iterates characters and the baseline
quietly stops matching. Both violate the documented exit-code contract
(2 = "ошибка использования/внутренняя") and print a Python traceback
instead of the Russian error message.
**Fix:** Validate schema in `load_baseline`:
```python
def load_baseline(path) -> dict:
    with open(path, "r", encoding="utf-8-sig") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("baseline должен быть JSON-объектом")
    units = data.get("units", {})
    if not isinstance(units, dict) or any(
            not isinstance(codes, list) for codes in units.values()):
        raise ValueError("baseline: units должен быть объектом "
                         "{путь: [коды]}")
    return data
```
The existing `except (OSError, ValueError)` in `main` then turns this into
a clean exit 2.

### WR-02: iter_pushbuttons treats junk-nested `*.pushbutton` dirs as real buttons

**File:** `tools/check_convention.py:622-641`
**Issue:** `tab_dir.rglob("*" + PUSHBUTTON_SUFFIX)` descends into junk
directories (`.vs/`) and into other pushbutton folders. Evidence: 4 of the
21 units currently in `tools/convention_baseline.json` are `.vs` artifacts
(`…/Высота этажа.pushbutton/.vs/Высота этажа.pushbutton` and three
siblings, lines 44-48, 57-61, 69-73, 88-92) — each grandfathered with
bogus MM000/MM005/MM006 entries. The checker simultaneously flags `.vs/`
as junk (MM013) on the parent and audits its contents as buttons. When a
user cleans the junk (as MM013 demands), the baseline entries go stale and
the "checked" count silently drops.
**Fix:** Skip candidates that have a junk dir or another `*.pushbutton`
among their ancestors relative to `tab_dir`:
```python
rel_parts = candidate.relative_to(tab_dir).parts[:-1]
if any(part in JUNK_DIR_NAMES or part.endswith(PUSHBUTTON_SUFFIX)
       for part in rel_parts):
    continue
```
Then regenerate the baseline (`--write-baseline`) to drop the 4 bogus units.

### WR-03: AST rules apply only to script.py — sibling modules in a button bypass all code rules

**File:** `tools/check_convention.py:552-567` (`check_pushbutton`)
**Issue:** `check_pushbutton` runs `_check_script_file` only on
`script.py`. Any other `*.py` in the pushbutton folder (a `helpers.py`
imported via `from . import helpers` or a runtime `sys.path` trick)
evades MM003 (BOM), MM008 (third-party imports), MM009 (wildcard), MM011
(bare except), MM014 (legacy bootstrap) entirely. A button can pass
`--strict` while smuggling `import requests` in a sibling module. Relative
imports are explicitly skipped by `_check_import_from_node`
(line 359), so this hole is reachable through the checker's own whitelist
logic.
**Fix:** Iterate all `*.py` in the button folder; apply file-level +
AST rules to each (header rules MM001/MM002/MM004 can stay script.py-only
if desired):
```python
for py_file in sorted(button_path.glob("*.py")):
    violations.extend(_check_script_file(py_file, unit_path, allowed_roots))
```
If script.py-only is a deliberate v1 scope decision, document it in the
module docstring and AGENTS.md — currently neither says so.

### WR-04: AGENTS.md rule 10 contradicts the canonical template's transaction skeleton

**File:** `AGENTS.md:90-91`; `templates/НоваяКнопка.pushbutton/script.py:68-75`; `agents/commands/mm-adopt-script.md:66-67`
**Issue:** AGENTS.md rule 10 (the single source of truth) prescribes:
"`transaction.Start()` и `transaction.Commit()` — в `try`". The canonical
template puts `transaction.Start()` OUTSIDE the try:
```python
transaction.Start()
try:
    transaction.Commit()
except Exception:
    transaction.RollBack()
    raise
```
The template's form is the correct one — if `Start()` itself throws inside
the try, the `except` would call `RollBack()` on a transaction that never
started, raising `InvalidOperationException` and masking the original
error. But an agent adapting a script per `mm-adopt-script.md` step 5
("`Start()`/`Commit()` в `try`") following the letter of the standard will
produce the broken form, and the checker has no MM rule to catch it.
**Fix:** Reword AGENTS.md rule 10 and mm-adopt-script step 5 to match the
template: "`transaction.Start()` перед `try`; `Commit()` — в `try`;
в `except` — `RollBack()` и `raise`."

### WR-05: Contract test writes bytecode cache into the live pyRevit extension directory

**File:** `tools/tests/test_revit_compat_contract.py:121-122`
**Issue:** `py_compile.compile(str(COMPAT_PATH), doraise=True)` writes
`MM LAB.extension/lib/__pycache__/revit_compat.cpython-*.pyc` into the
production extension directory. Confirmed on disk after this review's test
run. `.gitignore` covers it, but `/mm-doctor` (which runs this suite in
step 3) declares itself "Диагностика read-only: команда НИЧЕГО не правит"
(`agents/commands/mm-doctor.md:75-77`) — running it mutates the extension
tree. `mm-new-compat.md` step 5 even instructs deleting exactly this
`__pycache__` after a manual `py_compile`, yet the test recreates it on
every run.
**Fix:** Compile to a throwaway location:
```python
def test_compiles(self):
    with tempfile.TemporaryDirectory() as tmp:
        py_compile.compile(str(COMPAT_PATH),
                           cfile=os.path.join(tmp, "revit_compat.pyc"),
                           doraise=True)
```

### WR-06: revit_compat re-detects the Revit version instead of reusing the one validated by require_supported_version

**File:** `MM LAB.extension/lib/revit_compat.py:283-315` (`_units_map`), `352-384` (`create_floor`)
**Issue:** `require_supported_version(command_name, revit=app)` can succeed
via the explicit `revit` argument, but `_units_map()` and `create_floor()`
each call `get_revit_version()` with no argument. When
`builtins.__revit__` / `pyrevit.HOST_APP` are unavailable (the exact
scenario the explicit-argument path exists for), the second detection
returns None and both helpers silently take the 2022+ branch on a Revit
2020 host: `from Autodesk.Revit.DB import UnitTypeId` raises ImportError
(UnitTypeId does not exist in the 2020 API) and `Floor.Create` does not
exist before 2021. The user sees a cryptic import/attribute error instead
of the compat guarantee. `_UNITS_MAP` also caches the wrong branch for the
rest of the session.
**Fix:** Cache the version validated by `require_supported_version` in a
module-level variable and have `_units_map()`/`create_floor()` consult it
before re-detecting; or accept an optional `revit=`/`version=` parameter
on `convert_from_internal` / `convert_to_internal` / `create_floor`.

### WR-07: `--json --write-baseline` produces empty stdout, breaking the --json contract

**File:** `tools/check_convention.py:786-795`
**Issue:** The module docstring (line 30) promises for `--json`: "только
машинный вывод в stdout (ровно один JSON-объект)". With
`--write-baseline` + `--json`, the success message is suppressed
(`if not args.json`) and no JSON object is emitted — a machine consumer
gets zero bytes and exit 0. Nothing is printed at all.
**Fix:** Emit a status object in JSON mode:
```python
if args.json:
    print(json.dumps({"baseline_written": str(args.write_baseline),
                      "violations": len(violations)}, ensure_ascii=False))
```
(or reject the flag combination with exit 2).

## Info

### IN-01: iter_count silently consumes one-shot Python iterators

**File:** `MM LAB.extension/lib/revit_compat.py:411-431`
**Issue:** The final fallback counts by iterating. For a Python generator
(no `len`, no `.Count`) the caller's sequence is exhausted:
`n = iter_count(gen); for x in gen:` iterates nothing, silently. .NET
collections are safe; the hazard is only Python generators.
**Fix:** Document the constraint in the docstring ("не передавай
одноразовые итераторы/генераторы") or materialize: `return len(list(sequence))`.

### IN-02: MM014 fires multiple duplicate violations per legacy bootstrap block

**File:** `tools/check_convention.py:395-417`
**Issue:** Each `ast.Name` occurrence of `EXTENSION_ROOT`, each
non-canonical `sys.path` call and each 4-hop `join` produces a separate
MM014. The repo_bad fixture bootstrap alone yields several MM014 rows for
one logical defect — noisy output for the human reading `/mm-check`.
**Fix:** Deduplicate MM014 per file (keep the first line) before returning
from `check_ast_rules`.

### IN-03: MM007 registration and orphan checks tolerate suffixed layout entries inconsistently

**File:** `tools/check_convention.py:269-282` vs `493-522`
**Issue:** `_entry_has_folder` accepts a layout entry written with the
suffix (`child.name == entry` matches `Кнопка.pushbutton`), but
`_check_layout_registration` compares only the suffixless button name. A
layout entry `- Кнопка.pushbutton` yields a "не зарегистрирована"
violation with no corresponding orphan hint, sending the user in the wrong
direction.
**Fix:** In `_check_layout_registration`, also match `button_dir.name`
(with suffix) against `layout_names`, and word the message to say the
entry must be suffixless.

### IN-04: First-party import whitelist ignores package directories in lib/

**File:** `tools/check_convention.py:287-308`
**Issue:** `lib_dir.glob("*.py")` collects only module files. A future
first-party package `MM LAB.extension/lib/helpers/__init__.py` would make
`import helpers` a false MM008 error.
**Fix:** Also add stems of subdirectories containing `__init__.py`.

### IN-05: Dead test-helper branch — copy_button_to_tmp(with_panel=True) never used

**File:** `tools/tests/test_check_convention.py:67-84`
**Issue:** All three call sites pass `with_panel=False`; the `True` branch
(and the MM007-positive tmp scenario it enables) is dead code.
**Fix:** Either add a test exercising `with_panel=True` (registration
inside a tmp panel) or drop the parameter.

### IN-06: test_no_extra_mm_files does not scan agents/commands/

**File:** `tools/tests/test_mm_commands_catalog.py:133-151`
**Issue:** The doppelganger guard covers only the three adapter dirs. A
typo'd canonical file (`agents/commands/mm-chek.md`) would pass unnoticed
while `test_canonical_procedures_exist` still passes on the correct one.
**Fix:** Add `CANONICAL_DIR: {slug + ".md" for slug in SLUGS}` to the
`expected` map.

### IN-07: Gemini adapters validated by substring only, not parsed as TOML

**File:** `tools/tests/test_mm_commands_catalog.py:94-108`
**Issue:** `description =` / `prompt =` substring checks pass even if the
file is syntactically invalid TOML (unbalanced `"""`, stray key). The
declared floor for tools is Python ≥ 3.10, but when 3.11+ is available
`tomllib` could parse for free.
**Fix:** `try: import tomllib` and, when available, `tomllib.loads(text)`
asserting `description`/`prompt` keys.

### IN-08: Dangling separator at the end of the tab layout

**File:** `MM LAB.extension/MM Lab.tab/bundle.yaml:7`
**Issue:** The `- -----` entry is the last item in `layout:` — a vertical
separator with nothing after it. Harmless to the checker (skipped as
separator) but serves no purpose on the ribbon and looks like a leftover
from a removed panel entry.
**Fix:** Delete the trailing `- -----` line.

### IN-09: CLAUDE.md GSD section uses `python3` while the repo standardizes `py -3`

**File:** `CLAUDE.md:12-17`
**Issue:** All checker/test invocations in AGENTS.md and command
procedures use the Windows `py -3` launcher; the GSD Release Map section
prescribes `python3 RELEASE_MAP/gsd_release_sync.py ...`, which typically
does not resolve on Windows without an alias.
**Fix:** Use `py -3` for consistency (or note the required alias).

---

_Reviewed: 2026-07-24_
_Reviewer: Claude (gsd-code-reviewer)_
_Depth: standard_
