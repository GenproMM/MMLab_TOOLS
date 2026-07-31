---
name: VorScriptCreator
description: >
  Creates and edits ВОР Валидатор validation scripts — a run(doc, section, project, settings)
  function returning a ValidationResult, living at scripts/<Name>/<Name>.py with SCRIPT_ID
  metadata. Use when creating, editing, or debugging any such script — triggers:
  "создать скрипт валидации", "написать скрипт для ВОР", "добавить проверку",
  "create validation script", "add a check", or any work under scripts/<Name>/. Also use
  when wiring a show_results() results window, moving window-opening out of run(), or
  fixing a script whose run() opens UI or returns the wrong type.
version: 0.4.0
---

# VOR Script Creator

## Purpose

Generate validation scripts for the ВОР Валидатор pyRevit plugin. Scripts run inside Revit via `imp.load_module`. Each script validates or modifies Revit model elements and returns a structured result.

Every script lives in its own folder inside `scripts/`. The folder name and main `.py` file name must match. Companion modules (settings windows, helpers) go in an optional `lib/` subfolder.

> **Legacy**: Older scripts using `script.py` as the main file or single `.py` files in the `scripts/` root still work, but this skill generates only the new pattern.

## Canonical Folder Structure

```
scripts/
  Script Display Name/              # folder = SCRIPT_NAME
    Script Display Name.py           # main script file (name matches folder)
    lib/                             # optional — companion modules
      settings_window.py
      results_window.py
      helpers.py
```

**Rules**:
- Folder name = main `.py` file name (without `.py`) = value of `SCRIPT_NAME`
- `lib/` subfolder is optional — only created when the script has companion modules
- The engine adds the script's folder to `sys.path`. For `lib/` imports, inject it manually (see `show_settings` pattern below)

## Script Contract (Required)

Every script must contain a `run()` function and import `ValidationResult`:

```python
# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from core.validation_engine import ValidationResult

def run(doc, section, project, settings=None):
    # doc     — active Revit Document (Autodesk.Revit.DB.Document)
    # section — str, e.g. "АР"
    # project — str, e.g. "Ликино"
    # settings — dict | None, values from settings dialog
    return ValidationResult(
        check_name="Check name",
        passed=True,          # bool
        message="Description",
        elements=[]           # list[ElementId] of problem elements
    )
```

`settings` needs a `None` default because the engine dispatches via `try/except TypeError` against older scripts that don't accept it — without `None`, the call raises TypeError and the check silently fails. Always guard: `if settings: val = settings.get("key")`.

**Module reload — two lifetimes to know about**:
- **Between runs** (a NEW validation run starts): the engine `del`s the module from `sys.modules` and reloads it fresh. So module-level mutable state is RESET on each new run. Persisting anything across runs in module variables is futile — it gets wiped. (Module-level constants like `SCRIPT_NAME` are fine: they're re-evaluated on load.)
- **Within one session, between `run()` and `show_results()`**: the run window's "Открыть результат" button reuses the already-loaded module from `sys.modules` instead of reloading it. This is intentional — it lets `run()` cache `results_data` in a module-level `_last_results_data` that `show_results()` reads back instantly without re-running the check. So `_last_results_data` lives long enough for show_results, then is reset on the next run. See the "Separation of run() and show_results()" rule below.

## Metadata Variables (Top of File)

Place these **before** imports. They are extracted via regex without executing the file:

```python
SCRIPT_ID = "vor_a3f2c891"
# Needed for new scripts. Unique identifier in format: vor_ + 8 hex chars.
# Used to track the script if it's renamed or moved. The ID is stored in the
# central registry (scripts/script_registry.json) and in the user's config.
# Without it, the script still runs but config breaks if the file is moved.
# Generate a unique ID: random 8-hex, prefixed with "vor_".
# Verify uniqueness against scripts/script_registry.json before using.

SCRIPT_NAME = "Display Name"
# Overrides filename as the UI display name.
# Must match the folder name and main .py file name.

SCRIPT_DESCRIPTION = "What this script does"
# Shown below the script list when user clicks the script row.
# Supports single-line strings or triple-quoted multiline.

HAS_SETTINGS = True
# When True, a gear button appears next to the script in the UI.
```

## Two Approaches for Settings

### 1. Generic Window (Simple Scripts)

Define `SETTINGS_SCHEMA` — the plugin builds a settings dialog automatically:

```python
HAS_SETTINGS = True
SETTINGS_SCHEMA = [
    {"key": "selected_sheets", "type": "sheet_list", "label": "Select sheets",
     "sortable": True, "hide_unselected": True},
    {"key": "threshold",       "type": "text",    "label": "Threshold"},
    {"key": "max_count",       "type": "number",  "label": "Max elements"},
    {"key": "skip_hidden",     "type": "checkbox", "label": "Skip hidden elements",
     "checkbox_label": "Пропускать скрытые"},
    {"key": "sort_order",      "type": "select",  "label": "Sort order",
     "options": ["Ascending", "Descending"]},
]
```

#### Settings Types

| Type | UI Widget | Value in `settings` | Extra fields |
|---|---|---|---|
| `sheet_list` | Multi-select checklist of Revit sheets | `list[str]` — sheet numbers | `sortable` (bool), `hide_unselected` (bool) |
| `text` | Text input | `str` | — |
| `number` | Text input | `str` (convert manually) | — |
| `checkbox` | Checkbox | `bool` | `checkbox_label` (str) |
| `select` | Dropdown | `str` | `options` (list of str) |

When `sheet_list` has `sortable=True`, a sort dropdown appears above the list (populated with sheet parameter names). When `hide_unselected=True`, a "hide unselected" toggle appears.

### 2. Custom Window (Complex Scripts)

Define `show_settings(doc, current_settings)` — build your own WPF window. Takes priority over `SETTINGS_SCHEMA` if both exist.

**When a custom window is needed, invoke the VorUICreator skill** to generate the window class (`settings_window.py`, `results_window.py`, etc.). VorUICreator knows the design system, button templates, IronPython constraints, and layout patterns. Do not write WPF window code manually in this skill — delegate to VorUICreator for any file that subclasses `Window` or builds WPF controls.

```python
import os
import sys

def show_settings(doc, current_settings):
    """Show custom settings dialog. Returns settings dict or None."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.join(script_dir, "lib")
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)

    from settings_window import MySettingsWindow
    dlg = MySettingsWindow(doc, current_settings)
    return dlg.show_dialog()
```

**Why use custom window**: Complex UI logic (dynamic dependencies between settings, custom controls, conditional visibility beyond simple hide), or when the generic schema is insufficient.

**When to use generic window**: Simple selection of sheets, text/number inputs, checkboxes, dropdowns.

**Note**: The `lib/` subfolder must exist and contain the companion modules. The engine adds the script's own folder to `sys.path`, but not `lib/` — you must inject it in `show_settings()` or at the top of the script.

## Companion Modules in lib/

When a script needs helper modules, a custom settings window, or other companion files, place them in the `lib/` subfolder:

```
scripts/
  My Check/
    My Check.py
    lib/
      settings_window.py
      results_window.py
      helpers.py
```

To import from `lib/` anywhere in the main script (not just in `show_settings`):

```python
import os
import sys

_lib_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib")
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)
```

Place this block **after** the metadata variables and **before** other imports.

## Settings Storage

Settings auto-save to per-script files when the dialog closes:
- File: `%APPDATA%\pyRevit\ВОР_Валидатор\config_<script_name>.json`
- Structure: `{"Section|Project": {"key": "value", ...}}`
- Settings are specific to each Section+Project combination

## ValidationResult Fields

Constructor: `ValidationResult(check_name, passed, message, elements=None, skip_summary=False, script_id=None, duration_ms=0, has_results_window=False)`

- `check_name` (str) — identifier shown in the run window. Use `SCRIPT_NAME`.
- `passed` (bool) — `True` if check succeeded, `False` if problems found.
- `message` (str) — human-readable description (counts, details). Shown in the run window row.
- `elements` (list) — `ElementId` list of problematic elements. Empty if passed. **Serialized to the report** (`int(eid.IntegerValue)`), so it survives window reopen and powers the standard results window for scripts without `show_results`.
- `skip_summary` (bool) — set `True` on every `return`. The run window renders results; with the old default (`False`) the deprecated MessageBox summary can still fire alongside the dashboard.
- `script_id` (str|None) — the engine fills this from `SCRIPT_ID` automatically; you don't pass it.
- `duration_ms` (int) — the engine measures this automatically around `run()`.
- `has_results_window` (bool) — the engine sets this to `True` automatically when it detects a `show_results` function in your module. You don't pass it.

**You only set**: `check_name`, `passed`, `message`, `elements`, `skip_summary`. The engine populates the rest.

## IronPython 2.7 Constraints

- Use `codecs.open(path, "r", "utf-8")` instead of `open(..., encoding="utf-8")`.
- Use `imp` module, not `importlib`.
- String formatting: use `.format()`, not f-strings.
- No type hints, no `async`, no walrus operator.
- WPF `IsChecked` returns `Nullable<bool>` — use `bool(cb.IsChecked)` to normalize.
- WPF objects reject arbitrary Python attributes — `border._sort_combo = ...` raises `AttributeError`. Store references on `self` or in a dict.
- Lambda closures need default args to capture loop variables: `lambda s, e, n=name: handler(n)`.
- WPF CheckBox has `Checked` and `Unchecked` events (not `CheckedChanged`).
- Set `DialogResult` via `setattr(window, 'DialogResult', True)` — it's a .NET property.
- **Cyrillic encoding**: IronPython's `imp.load_module()` mangles UTF-8 Cyrillic in runtime string literals (both module-level and inside functions), producing garbage like `Ð—Ð°Ð¿Ð¾Ð»...`. Escape every Cyrillic rune in `u"..."` literals as `\uXXXX` — e.g. `u"\u0417\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u044c"` instead of `u"Заполненность"`. This applies everywhere a Cyrillic literal is used at runtime: SCRIPT_NAME, SCRIPT_DESCRIPTION, parameter names (PARAM_NAME etc.), UI labels, messages, format strings. Comments and docstrings are safe — they aren't executed. Keep ASCII portions readable and escape only Cyrillic, e.g. `u"GP_01_\u0417\u043e\u043d\u0430"`.
- **Separation of run() and show_results()**: `run(doc, section, project, settings)` computes and returns a `ValidationResult`. Opening a results window inside `run()` blocks the run window's dashboard and fires a window on every run, breaking the "Открыть результат" on-demand flow. Move window-opening into an optional `show_results(doc, section, project, settings)` function — the run window calls it when the user clicks "Открыть результат". Cache the computed `results_data` in a module-level variable (e.g. `_last_results_data`) so `show_results` can re-open the window instantly without re-running the check. Scripts without `show_results` get the engine's standard results window, which renders `ValidationResult.elements`.

  ```python
  _last_results_data = None

  def run(doc, section, project, settings=None):
      global _last_results_data
      # ... compute ...
      _last_results_data = results_data   # cache for show_results
      return ValidationResult(check_name=SCRIPT_NAME, ..., elements=problem_ids,
                              skip_summary=True)

  def show_results(doc, section, project, settings=None):
      """Open the custom results window. Called by the run window on demand."""
      global _last_results_data
      if not _last_results_data:
          return
      _show_results_window(doc, _last_results_data)
  ```

## Script Templates

### Minimal Validation

```
scripts/
  Стены без марки/
    Стены без марки.py
```

**`Стены без марки/Стены без марки.py`**:

```python
# -*- coding: utf-8 -*-
"""Docstring describing the script."""

SCRIPT_ID = "vor_XXXXXXXX"  # Auto-generated; do not reuse across scripts
SCRIPT_NAME = u"\u0421\u0442\u0435\u043d\u044b \u0431\u0435\u0437 \u043c\u0430\u0440\u043a\u0438"
SCRIPT_DESCRIPTION = u"\u041e\u043f\u0438\u0441\u0430\u043d\u0438\u0435 \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438."

from pyrevit import revit, DB
from core.validation_engine import ValidationResult


def run(doc, section, project, settings=None):
    """Run the validation check."""
    try:
        elements = (DB.FilteredElementCollector(doc)
                    .OfCategory(DB.BuiltInCategory.OST_Whatever)
                    .WhereElementIsNotElementType()
                    .ToElements())

        errors = []
        for elem in elements:
            pass  # validation logic

        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=len(errors) == 0,
            message=u"\u041f\u0440\u043e\u0432\u0435\u0440\u0435\u043d\u043e: {}. \u041f\u0440\u043e\u0431\u043b\u0435\u043c: {}".format(len(elements), len(errors)),
            elements=errors,
            skip_summary=True
        )

    except Exception as e:
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=False,
            message=u"\u041e\u0448\u0438\u0431\u043a\u0430: {}".format(str(e)),
            skip_summary=True
        )
```

### Script with Generic Settings

```
scripts/
  Проверка листов/
    Проверка листов.py
```

**`Проверка листов/Проверка листов.py`**:

```python
# -*- coding: utf-8 -*-
"""Script with generic settings dialog."""

SCRIPT_ID = "vor_XXXXXXXX"  # Auto-generated; do not reuse across scripts
SCRIPT_NAME = u"\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u043b\u0438\u0441\u0442\u043e\u0432"
SCRIPT_DESCRIPTION = u"\u041e\u043f\u0438\u0441\u0430\u043d\u0438\u0435."
HAS_SETTINGS = True
SETTINGS_SCHEMA = [
    {"key": "selected_sheets", "type": "sheet_list", "label": "Sheets to check",
     "sortable": True, "hide_unselected": True},
]

from pyrevit import revit, DB
from core.validation_engine import ValidationResult


def run(doc, section, project, settings=None):
    sheet_numbers = settings.get("selected_sheets", []) if settings else []
    if not sheet_numbers:
        return ValidationResult(
            check_name=SCRIPT_NAME, passed=False,
            message=u"\u041d\u0435 \u0432\u044b\u0431\u0440\u0430\u043d\u043e \u043d\u0438 \u043e\u0434\u043d\u043e\u0433\u043e \u043b\u0438\u0441\u0442\u0430.",
            skip_summary=True
        )
    # ... validation logic
```

### Script with Custom Settings Window

```
scripts/
  Проверка марки по шаблону/
    Проверка марки по шаблону.py
    lib/
      settings_window.py
```

**`Проверка марки по шаблону/Проверка марки по шаблону.py`**:

```python
# -*- coding: utf-8 -*-
"""Script with custom settings window."""

SCRIPT_ID = "vor_XXXXXXXX"  # Auto-generated; do not reuse across scripts
SCRIPT_NAME = u"\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u043c\u0430\u0440\u043a\u0438 \u043f\u043e \u0448\u0430\u0431\u043b\u043e\u043d\u0443"
SCRIPT_DESCRIPTION = u"\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u0441 \u043d\u0430\u0441\u0442\u0440\u0430\u0438\u0432\u0430\u0435\u043c\u044b\u043c \u043e\u043a\u043d\u043e\u043c."
HAS_SETTINGS = True

import os
import sys

from pyrevit import revit, DB
from core.validation_engine import ValidationResult


def show_settings(doc, current_settings):
    """Custom settings dialog."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.join(script_dir, "lib")
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)

    from settings_window import MySettingsWindow
    dlg = MySettingsWindow(doc, current_settings)
    return dlg.show_dialog()


def run(doc, section, project, settings=None):
    # ... validation logic using settings
    pass
```

**`lib/settings_window.py`** — generated by VorUICreator, not written by hand. Invoke the **VorUICreator** skill to produce it: it emits the correct design-system constants, styled button templates, explicit imports (no `import *`), and handles all IronPython WPF constraints. VorScriptCreator only writes the parent script's `show_settings(doc, current_settings)` wrapper that imports and calls it.

### Script with Custom Results Window (show_results)

This is the pattern for any script that needs its own results UI (rich tables, exclude/restore buttons, custom grouping). **The script must NOT open the window inside `run()`** — the run window's "Открыть результат" button calls `show_results()` instead. `run()` only computes and caches data; `show_results()` re-opens the window on demand without re-running the check.

```
scripts/
  Проверка с отчётом/
    Проверка с отчётом.py
    lib/
      results_window.py    # generated by VorUICreator (modeless ResultsWindow)
```

**`Проверка с отчётом/Проверка с отчётом.py`**:

```python
# -*- coding: utf-8 -*-
"""Script with a custom results window shown via show_results()."""

SCRIPT_ID = "vor_XXXXXXXX"
SCRIPT_NAME = u"\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u0441 \u043e\u0442\u0447\u0451\u0442\u043e\u043c"
SCRIPT_DESCRIPTION = u"\u041e\u043f\u0438\u0441\u0430\u043d\u0438\u0435."

import os
import sys

from pyrevit import revit, DB
from core.validation_engine import ValidationResult

# Кэш данных последнего прогона. Переживает между run() и show_results()
# в пределах сессии (run window переиспользует модуль из sys.modules).
# Сбрасывается в None при следующем новом прогоне.
_last_results_data = None


def run(doc, section, project, settings=None):
    """Compute validation and cache results. Opening a window here breaks the dashboard — use show_results()."""
    global _last_results_data
    try:
        problems = []   # list of {"id": ElementId, ...} dicts
        # ... collect elements, find problems ...

        # Кэшируем данные для show_results.
        _last_results_data = {
            "category_name": problems,
            "script_name": SCRIPT_NAME,
        }

        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=len(problems) == 0,
            message=u"\u041f\u0440\u043e\u0431\u043b\u0435\u043c: {}".format(len(problems)),
            elements=[p["id"] for p in problems],
            skip_summary=True,
        )
    except Exception as e:
        return ValidationResult(
            check_name=SCRIPT_NAME, passed=False,
            message=u"\u041e\u0448\u0438\u0431\u043a\u0430: {}".format(str(e)),
            skip_summary=True,
        )


def show_results(doc, section, project, settings=None):
    """Open the custom results window. Called by the run window on demand."""
    global _last_results_data
    if not _last_results_data:
        return
    _show_results_window(doc, _last_results_data)


def _show_results_window(doc, results_data):
    """Open the modeless results window (singleton via __main__)."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.join(script_dir, "lib")
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)

    import __main__
    attr = "_myscript_results_window"
    old = getattr(__main__, attr, None)
    if old:
        try:
            old.Close()
        except Exception:
            pass

    from results_window import ResultsWindow   # generated by VorUICreator
    win = ResultsWindow(doc, results_data)
    setattr(__main__, attr, win)
    win.Show()
```

**Key rules for this pattern**:
- `run()` returns a `ValidationResult` and caches data in `_last_results_data`. Calling `_show_results_window`/`Show()`/`MessageBox` from `run()` blocks the dashboard and fires a window on every run — leave display to `show_results()`. (See the run()/show_results() separation rule in IronPython Constraints.)
- `show_results()` signature must be exactly `show_results(doc, section, project, settings=None)`. The engine detects it via `hasattr(module, "show_results")` and sets `has_results_window=True` automatically.
- The `__main__` attribute name must be **unique per script** (e.g. `_myscript_results_window`) to avoid clashing with other scripts' result windows.
- For scripts WITHOUT a custom window, omit `show_results` entirely — the engine renders `elements` in the standard results window (`ui/standard_results_window.py`).

## Common Revit API Patterns

For element collection, parameter access, sheet operations, and transaction wrapping, see **`references/revit-api-patterns.md`** — it has the canonical code snippets (FilteredElementCollector, get_Parameter, AsString/AsDouble/AsInteger, DB.Transaction, etc.). The reference file is kept in sync with the real scripts.

## Validation-Only vs Modification Scripts

**Validation-only** (read model, report problems): No transaction needed. Return `passed=False` with problem element IDs.

**Modification** (change model data): Wrap all writes in a transaction. Always use `try/finally` or `try/except` with `.RollBack()` on failure.

## Additional Resources

### Reference Files

For detailed Revit API patterns and common validation scenarios:
- **`references/revit-api-patterns.md`** — Element collection, parameter access, transactions, sheet operations

### Example Scripts

Working scripts in `examples/`:
- **`examples/simple-check/simple-check.py`** — Minimal read-only validation (check element parameters)
- **`examples/check-with-settings/check-with-settings.py`** — Script with `SETTINGS_SCHEMA` generic settings
- **`examples/modification-script/modification-script.py`** — Script that modifies model elements in a transaction
- **`examples/folder-script-example/`** — Script with custom settings window, `lib/` subfolder for companion modules

Live scripts in `scripts/` (best real-world references):
- **`scripts/Заполненность GP_01_Зона/`** — full `show_results` pattern: `run()` caches `_last_results_data`, `show_results()` re-opens a custom modeless window (category expanders, click-to-select)
- **`scripts/Заполненность GP_01_Коды/`** — `show_results` with a rich custom window (2×2 grid, exclude/restore buttons)
- **`scripts/Загрузка параметров GP_01/`** — modification script WITHOUT `show_results`: returns `elements`, engine renders them in the standard results window
