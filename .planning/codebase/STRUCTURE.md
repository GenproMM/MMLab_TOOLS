# Codebase Structure

**Analysis Date:** 2026-06-08

## Directory Layout

```text
MMLab_TOOLS/
|-- MM_LAB.extension/                 # pyRevit extension payload (tab/panel/command tree)
|   `-- MM Lab.tab/
|       |-- bundle.yaml               # Tab metadata and panel layout
|       |-- АРХИТЕКТУРА.panel/        # Architecture discipline commands
|       |-- ИОС.panel/                # MEP/IOS discipline commands
|       `-- КООРДИНАЦИЯ.panel/        # Coordination commands
|-- lib/                              # Vendored Python dependencies (openpyxl, et_xmlfile)
|-- .github/get-shit-done/            # GSD automation toolchain and templates
|-- .claude/                          # Claude-side mirrors/settings for GSD workflow
|-- .planning/codebase/               # Generated codebase mapping artifacts
|-- extension.json                    # Root pyRevit extension manifest
`-- README.md                         # Repository readme
```

## Directory Purposes

**`MM_LAB.extension/`:**
- Purpose: Runtime-consumed extension package for pyRevit.
- Contains: `*.tab`, `*.panel`, `*.pushbutton` folders, plus `bundle.yaml`, `script.py`, icons, command READMEs.
- Key files: `MM_LAB.extension/MM Lab.tab/bundle.yaml`, `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/bundle.yaml`, `MM_LAB.extension/MM Lab.tab/ИОС.panel/bundle.yaml`.

**`MM_LAB.extension/MM Lab.tab/*/*.pushbutton/`:**
- Purpose: Atomic command modules.
- Contains: Executable `script.py`, command `bundle.yaml`, optional `README.md`, `icon.png`.
- Key files: `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py`, `MM_LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`, `MM_LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/script.py`.

**`lib/`:**
- Purpose: Vendor directory for Python packages imported by commands.
- Contains: `openpyxl` and `et_xmlfile` package trees and dist-info metadata.
- Key files: `lib/openpyxl/__init__.py`, `lib/et_xmlfile/__init__.py`.

**`.github/get-shit-done/`:**
- Purpose: Repository automation and planning workflow implementation.
- Contains: CLI entrypoint, module library, templates, skills/workflows references.
- Key files: `.github/get-shit-done/bin/gsd-tools.cjs`, `.github/get-shit-done/bin/lib/state.cjs`, `.github/get-shit-done/bin/lib/phase.cjs`, `.github/get-shit-done/bin/lib/verify.cjs`.

**`.planning/codebase/`:**
- Purpose: Generated architecture/quality/stack/concerns summaries for downstream planning/execution tools.
- Contains: Mapping markdown outputs (current task output location).
- Key files: `.planning/codebase/ARCHITECTURE.md`, `.planning/codebase/STRUCTURE.md`.

## Key File Locations

**Entry Points:**
- `extension.json`: Declares extension identity and registration.
- `MM_LAB.extension/MM Lab.tab/bundle.yaml`: Declares tab title and panel layout.
- `.github/get-shit-done/bin/gsd-tools.cjs`: CLI router for workflow operations.

**Configuration:**
- `MM_LAB.extension/MM Lab.tab/*/bundle.yaml`: Panel and command ordering/labels.
- `.claude/settings.json`: Claude local settings for workflow behavior.

**Core Logic:**
- `MM_LAB.extension/MM Lab.tab/*/*.pushbutton/script.py`: Revit automation implementation.
- `.github/get-shit-done/bin/lib/*.cjs`: Planning workflow logic modules.

**Testing:**
- Not detected in repository root (`*.test.*`, `*.spec.*` not found in primary project code).

## Naming Conventions

**Files:**
- Command executables use fixed name `script.py` per pushbutton module.
- UI metadata uses `bundle.yaml` at tab, panel, and command scopes.
- Root extension manifest is `extension.json`.

**Directories:**
- pyRevit hierarchy follows `<Tab>.tab/<Panel>.panel/<Command>.pushbutton` pattern.
- GSD CLI modules use lowercase kebab/camel style filenames under `.github/get-shit-done/bin/lib/` (example: `profile-pipeline.cjs`, `schema-detect.cjs`).

## Where to Add New Code

**New Feature (Revit command):**
- Primary code: create new `*.pushbutton` directory under target panel in `MM_LAB.extension/MM Lab.tab/<DISCIPLINE>.panel/`.
- Tests: not currently standardized; if adding tests, introduce a dedicated test directory outside runtime extension tree to avoid pyRevit loading side effects.

**New Component/Module (workflow automation):**
- Implementation: add module in `.github/get-shit-done/bin/lib/` and wire command in `.github/get-shit-done/bin/gsd-tools.cjs`.

**Utilities:**
- Shared helpers for a single command should stay inside that command's `script.py`.
- Cross-command Python reuse is not currently centralized; prefer introducing a clearly named shared package folder only when at least two commands need identical logic.

## Special Directories

**`lib/`:**
- Purpose: vendored third-party Python dependencies.
- Generated: No (manually installed/copied into repo).
- Committed: Yes.

**`.planning/`:**
- Purpose: workflow state and generated planning artifacts.
- Generated: Yes (by GSD tooling and mapping commands).
- Committed: Yes (intended as process artifact in this repository).

**Nested `.vs/` inside some pushbutton folders:**
- Purpose: local IDE metadata.
- Generated: Yes.
- Committed: Currently present; treat as tooling byproduct, not runtime architecture.

---

*Structure analysis: 2026-06-08*
