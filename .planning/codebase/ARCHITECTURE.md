# Architecture

**Analysis Date:** 2026-06-08

## Pattern Overview

**Overall:** Plugin command-pack architecture (pyRevit extension tree) plus repository-local workflow tooling (Node.js CLI for GSD planning operations).

**Key Characteristics:**
- UI composition is declarative via `bundle.yaml` files under `MM_LAB.extension/MM Lab.tab/...`.
- Business logic is command-local and script-centric (`script.py` per `*.pushbutton` directory).
- Workflow orchestration is centralized in `.github/get-shit-done/bin/gsd-tools.cjs` with feature modules in `.github/get-shit-done/bin/lib/*.cjs`.

## Layers

**Extension Metadata Layer:**
- Purpose: Register extension identity and top-level tab/panel layout.
- Location: `extension.json`, `MM_LAB.extension/MM Lab.tab/bundle.yaml`, `MM_LAB.extension/MM Lab.tab/*/bundle.yaml`
- Contains: Extension manifest (`type`, `name`) and panel ordering (`layout`).
- Depends on: pyRevit extension discovery conventions.
- Used by: Revit/pyRevit runtime when loading the extension UI.

**Command UI Metadata Layer:**
- Purpose: Define each clickable command and user-facing labels/tooltips.
- Location: `MM_LAB.extension/MM Lab.tab/*/*.pushbutton/bundle.yaml`
- Contains: Command title/tooltip metadata.
- Depends on: Panel-level `layout` references (for discoverability order).
- Used by: pyRevit ribbon rendering.

**Command Execution Layer (Python):**
- Purpose: Implement domain behavior for architecture/MEP/coordination operations.
- Location: `MM_LAB.extension/MM Lab.tab/*/*.pushbutton/script.py`
- Contains: Revit API calls, element collectors, transactions, dialogs, and command-specific algorithms.
- Depends on: `Autodesk.Revit.DB`, `Autodesk.Revit.UI`, `pyrevit` modules and `__revit__` context.
- Used by: Individual pushbutton activation at runtime.

**Bundled Dependency Layer (Python vendor libs):**
- Purpose: Provide Python packages not guaranteed in host environment.
- Location: `lib/openpyxl/**`, `lib/et_xmlfile/**`
- Contains: Vendored third-party packages for Excel export flows.
- Depends on: Script-side `sys.path` injection (example in `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`).
- Used by: Commands that need offline packaging-safe dependencies.

**Planning/Automation Layer (Node CLI):**
- Purpose: Manage `.planning` artifacts, phase routing, validation, and workflow state.
- Location: `.github/get-shit-done/bin/gsd-tools.cjs`, `.github/get-shit-done/bin/lib/*.cjs`
- Contains: Command router (`gsd-tools.cjs`) and domain modules (`phase.cjs`, `state.cjs`, `verify.cjs`, `roadmap.cjs`, etc.).
- Depends on: Node.js CommonJS runtime and filesystem/git access.
- Used by: Local CLI calls such as `node .github/get-shit-done/bin/gsd-tools.cjs init map-codebase`.

## Data Flow

**pyRevit Command Invocation Flow:**

1. pyRevit loads `extension.json` and tab/panel `bundle.yaml` metadata.
2. User clicks a command resolved from `*.pushbutton/bundle.yaml`.
3. Runtime executes corresponding `*.pushbutton/script.py` and injects `__revit__` context.
4. Script reads model state via `FilteredElementCollector`, computes updates, and writes through `Transaction`.
5. Script reports outcome via `TaskDialog`, `pyrevit.forms`, or output channel.

**GSD Workflow Execution Flow:**

1. User invokes `.github/get-shit-done/bin/gsd-tools.cjs`.
2. Router parses CLI args and resolves root/workstream.
3. Command dispatch calls module handlers in `.github/get-shit-done/bin/lib/*.cjs`.
4. Modules mutate/inspect `.planning` state and return machine-usable output.

**State Management:**
- Revit-side state is document-bound and transient per command invocation (`doc = __revit__.ActiveUIDocument.Document`).
- Workflow state is file-backed under `.planning/` and manipulated by GSD CLI modules.

## Key Abstractions

**Panel-to-Command Packaging:**
- Purpose: Keep UI composition and command logic loosely coupled.
- Examples: `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/bundle.yaml`, `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py`
- Pattern: Metadata (`bundle.yaml`) references command directory; command directory encapsulates code/assets/docs.

**Transaction-Bounded Mutation:**
- Purpose: Ensure model changes are atomic and rollback-capable.
- Examples: `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py`, `MM_LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py`
- Pattern: `Transaction.Start()` -> parameter updates -> `Commit()` or `RollBack()` on failure.

**Vendor Path Injection for Optional Dependencies:**
- Purpose: Make scripts self-sufficient in constrained pyRevit runtime.
- Examples: `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`
- Pattern: Compute extension root -> prepend `lib` to `sys.path` -> import vendor package.

## Entry Points

**pyRevit Extension Root:**
- Location: `extension.json`
- Triggers: pyRevit extension scan/load.
- Responsibilities: Declares extension identity and installation-level metadata.

**UI Tab Definition:**
- Location: `MM_LAB.extension/MM Lab.tab/bundle.yaml`
- Triggers: Ribbon tab rendering.
- Responsibilities: Defines tab title and panel order.

**Per-Command Executable Entrypoints:**
- Location: `MM_LAB.extension/MM Lab.tab/*/*.pushbutton/script.py`
- Triggers: User pushbutton click.
- Responsibilities: Execute one bounded domain operation against active Revit document/family.

**Planning CLI Entrypoint:**
- Location: `.github/get-shit-done/bin/gsd-tools.cjs`
- Triggers: Node CLI invocation.
- Responsibilities: Route commands, parse args, and invoke workflow modules.

## Error Handling

**Strategy:** Defensive command-local handling with user-facing dialogs plus transaction rollback for write paths.

**Patterns:**
- Guard clauses before write operations (example: family-only checks in `MM_LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/script.py`).
- `try/except` blocks around API calls with fallback behavior and message reporting.
- Explicit rollback semantics when mutation fails (`Transaction.RollBack()`).

## Cross-Cutting Concerns

**Logging:** Primarily `print`/script output and dialog reporting in command scripts.
**Validation:** Runtime checks on active document, phase, element presence, and parameter availability.
**Authentication:** Not applicable; scripts run in-process under the current Revit user/session.

---

*Architecture analysis: 2026-06-08*
