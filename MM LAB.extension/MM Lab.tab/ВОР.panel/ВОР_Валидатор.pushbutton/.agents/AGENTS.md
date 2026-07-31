# CLAUDE.md

pyRevit plugin for validating Bill of Quantities (ВОР) parameters in Revit. WPF + XAML UI on IronPython 2.7 / Revit 2024. No build/test commands — runs inside Revit via pyRevit extension loader.

## Architecture

`script.py` → migration → `ui/MainWindow` (XamlReader parses XAML at runtime). Three windows: main (scripts + description), settings (sections/projects CRUD + import), and script settings (dynamic from SETTINGS_SCHEMA).

**Data flow**: User picks Section+Project → config auto-loads from JSON → scripts populate → run validation → results in MessageBox.

**Per-user config**: `%APPDATA%\pyRevit\ВОР_Валидатор\config.json`. Both `config_manager.py` and `settings_store.py` read/write the same file (read full → modify → write full). Per-script settings in separate `config_<name>.json` files, keyed by `"Section|Project"`.

**Config entry**: `{"name": "...", "path": "...", "enabled": true, "id": "vor_xxxx"}` — `id` is optional (backward compatible).

**Validation pipeline**: `validation_engine.run_validation()` loads each script via `imp.load_module` and calls `run(doc, section, project, settings)`.

## Script Metadata (extracted via regex, no execution)

- `SCRIPT_ID = "vor_a3f2c891"` — REQUIRED for new scripts. Format: `vor_` + 8 hex chars. Stable ID across renames/moves.
- `SCRIPT_NAME = "Display Name"` — overrides filename as display name
- `SCRIPT_DESCRIPTION = "Description text"` — shown when script selected in UI
- `HAS_SETTINGS = True` — shows gear button next to script
- `SETTINGS_SCHEMA = [...]` — defines generic settings UI (types: `sheet_list`, `text`, `number`, `checkbox`, `select`)
- `show_settings(doc, current_settings)` — custom settings window, takes priority over SETTINGS_SCHEMA

## Script ID System

**Registry**: `scripts/script_registry.json` — shared file on network drive, maps IDs to paths/display names. Managed by `core/registry.py`.

**Startup**: `migrate_script_ids_startup()` adds `id` to config entries lacking it. `resolve_script_paths()` fixes broken paths via registry lookup.

**Missing scripts**: Gray italic "(не найден)", disabled checkbox, amber background. User can delete stale entry.

**Duplicate IDs**: `register_script()` rejects if ID registered for different file. Config loading skips duplicates with warning.

## Folder Scripts

`scripts/Script Name/Script Name.py` — every script in its own folder, main file name matches folder name. Companion modules (settings_window, helpers) in optional `lib/` subfolder. `lib/` must be added to `sys.path` manually for imports. Engine adds the script's own folder to `sys.path` automatically.

## IronPython Constraints

- `codecs.open(path, "r", "utf-8")` — no `encoding` param on `open()`
- `imp.find_module`/`imp.load_module` — no `importlib`
- WPF: `IsChecked` → `Nullable<bool>`, `DialogResult` via `setattr`, no arbitrary Python attrs on WPF objects
- Lambda closures: `lambda s, e, n=name: handler(n)` (default args capture loop vars)
- `GridLength`, `GridUnitType` from `System.Windows`

## Non-Obvious Details

- Deleting section/project preserves `rp_configs` entry; `visible: false` hides without removing
- `run()` accepts optional `settings` arg — old scripts without it work via `try/except TypeError`
- `migrate_inline_settings()` moves old inline settings to per-script files; `migrate_script_ids_startup()` adds IDs to config — both idempotent
- Scripts store absolute paths — SCRIPT_ID enables recovery if paths break
- `ValidationResult` in `core/validation_engine.py` (canonical, field: `elements`) — duplicate in `base_validation.py` is unused
- pyRevit folder hierarchy (`.extension/.tab/.panel/.pushbutton`) auto-builds the Revit ribbon UI
- Legacy scripts with `script.py` as main file or single `.py` in `scripts/` root still work (engine uses parent folder name for `script.py`, base filename otherwise). New scripts always use `<FolderName>.py`
