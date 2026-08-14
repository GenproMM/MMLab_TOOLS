# Coding Conventions

**Analysis Date:** 2026-06-08

## Naming Patterns

**Files:**
- pyRevit command entry files use fixed filename `script.py` under button folders, e.g. `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Высота этажа.pushbutton/script.py` and `MM_LAB.extension/MM Lab.tab/КООРДИНАЦИЯ.panel/ВерсияШаблона.pushbutton/script.py`.
- Command metadata is colocated in `bundle.yaml`, e.g. `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/bundle.yaml`.

**Functions:**
- `snake_case` for helpers and domain logic: `feet_to_meters`, `get_floor_thickness`, `collect_elements`, `parse_revision_date` in `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Высота этажа.pushbutton/script.py`, `MM_LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py`, and `MM_LAB.extension/MM Lab.tab/КООРДИНАЦИЯ.panel/ВерсияШаблона.pushbutton/script.py`.

**Variables:**
- Local variables use `snake_case` (`levels_collector`, `closest_floor`, `target_params`).
- Constants use `UPPER_SNAKE_CASE` (`TARGET_PARAMETER_NAME`, `UNDEFINED_LABELS`, `SHEET_GROUPING_PARAMS`).

**Types:**
- Classes use `PascalCase` (`TreeNode`, `SupplyFlagDecision`, `SystemClassificationSelectionForm`).

## Code Style

**Formatting:**
- Tool used: Not detected (no `.editorconfig`, `pyproject.toml`, `setup.cfg`, `ruff.toml`, or formatter config at repository root).
- Style is hand-formatted with 4-space indentation and explicit section headers (`# === IMPORTS ===`, `# === MAIN ===`) in `MM_LAB.extension/MM Lab.tab/КООРДИНАЦИЯ.panel/ВерсияШаблона.pushbutton/script.py`.

**Linting:**
- Tool used: Not detected (no `flake8`, `pylint`, `ruff`, or mypy configuration files detected).
- Convention to follow: keep scripts self-validating with defensive checks and explicit `try/except` blocks, as seen in `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Высота этажа.pushbutton/script.py`.

## Import Organization

**Order:**
1. Standard library imports (`sys`, `os`, `re`, `datetime`).
2. CLR/Revit assembly loading (`import clr`, `clr.AddReference(...)`).
3. Revit API imports (`from Autodesk.Revit.DB import ...`, `from Autodesk.Revit.UI import ...`).
4. pyRevit/UI framework imports (`from pyrevit import script`, `from System.Windows.Forms import ...`).

**Path Aliases:**
- No Python package aliasing scheme detected.
- Vendor injection uses `sys.path.insert(0, VENDOR_DIR)` in `MM_LAB.extension/MM Lab.tab/КООРДИНАЦИЯ.panel/ВерсияШаблона.pushbutton/script.py` to load `lib`.

## Error Handling

**Patterns:**
- Guard for missing active document with `TaskDialog.Show(...)` and early return (`get_document` pattern in `MM_LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py`).
- Transaction safety pattern: `transaction.Start()` + `Commit()` in `try`, `RollBack()` in `except`.
- Top-level exception trap shows UI error instead of traceback (`show_error(COMMAND_NAME, ex)`).

## Logging

**Framework:**
- UI-first reporting through `Autodesk.Revit.UI.TaskDialog`.
- Secondary textual output via `pyrevit.script.get_output()` in `MM_LAB.extension/MM Lab.tab/КООРДИНАЦИЯ.panel/ВерсияШаблона.pushbutton/script.py`.

**Patterns:**
- User-facing completion/error summaries use dialog messages.
- Console `print(...)` appears only for local exception notes in utility functions (`get_materials_with_kod_edinicy` in `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py`).

## Comments

**When to Comment:**
- Comments are used to mark sections, Revit-specific assumptions, and business constants.
- Multiline module docstrings explain command purpose and compatibility.

**JSDoc/TSDoc:**
- Not applicable; repository command code is Python.
- Python docstrings are used on modules/functions/classes.

## Function Design

**Size:**
- Utility functions are small and single-purpose (`normalize_text`, `is_writable`).
- Entry scripts can be large and procedural; keep new complex logic split into helper functions before command execution blocks.

**Parameters:**
- Domain objects (`doc`, `element`, `parameter`) are passed explicitly.
- Localized string literals for parameter names are passed via fallback tuples/lists.

**Return Values:**
- Helper functions return primitives/status tokens (`True/False`, strings like `"updated"`, `"already"`, `"skipped"`).
- UI command handlers prefer side effects (parameter updates, dialogs) with simple success booleans.

## Module Design

**Exports:**
- No package export layer for command scripts; each `script.py` is an executable module entry point.
- Shared logic is currently copied across button scripts (same helper block in `MM_LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py` and `MM_LAB.extension/MM Lab.tab/ИОС.panel/Приточный по классификации.pushbutton/script.py`).

**Barrel Files:**
- Not used for extension command code.
- `lib/openpyxl` and `lib/et_xmlfile` are vendored third-party packages and should not define style for first-party scripts.

---

*Convention analysis: 2026-06-08*
