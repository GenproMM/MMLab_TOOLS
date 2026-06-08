# Technology Stack

**Analysis Date:** 2026-06-08

## Languages

**Primary:**
- Python 3.x (pyRevit-hosted) - Revit automation scripts in `MM LAB.extension/MM Lab.tab/**/script.py`.
- Python (IronPython-compatible subset) - MEP scripts explicitly marked for IronPython in `MM LAB.extension/MM Lab.tab/ИОС.panel/Доп расход 0.pushbutton/script.py` and sibling tools.

**Secondary:**
- JavaScript (CommonJS, Node.js runtime) - GSD CLI and helpers in `.github/get-shit-done/bin/gsd-tools.cjs` and `.github/get-shit-done/bin/lib/*.cjs`.
- YAML/JSON/Markdown - pyRevit UI metadata and workflow content in `MM LAB.extension/MM Lab.tab/**/bundle.yaml`, `extension.json`, and `.github/get-shit-done/workflows/*.md`.

## Runtime

**Environment:**
- Autodesk Revit host process with pyRevit bridge (scripts access `__revit__` and `Autodesk.Revit.*` APIs), e.g. `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Мокрые зоны.pushbutton/script.py`.
- Mixed pyRevit engines: CPython3 tool detected via shebang `#! python3` in `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`, while several tools are labeled IronPython in docstrings under `MM LAB.extension/MM Lab.tab/ИОС.panel/*/script.py`.
- Node.js CLI runtime for planning automation (`#!/usr/bin/env node`) in `.github/get-shit-done/bin/gsd-tools.cjs`.

**Package Manager:**
- Python: `pip` usage documented for local vendor install (`pip install openpyxl --target ...`) in `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`.
- Node: no repository-level `package.json` for `.github/get-shit-done`; scripts are plain CommonJS modules.
- Lockfile: missing at repository root.

## Frameworks

**Core:**
- pyRevit extension model - extension metadata in `extension.json` and command metadata in `MM LAB.extension/MM Lab.tab/**/bundle.yaml`.
- Autodesk Revit .NET API (`RevitAPI`, `RevitAPIUI`) loaded via `clr.AddReference(...)` in multiple scripts such as `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py`.
- .NET UI interop (WinForms/WPF) via `System.Windows.Forms`, `System.Drawing`, and WPF assemblies in scripts like `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py`.

**Testing:**
- Not detected (no `pytest`, `unittest`, `jest`, `vitest`, or test directory conventions found).

**Build/Dev:**
- GSD CLI utilities (`.github/get-shit-done/bin/gsd-tools.cjs`) provide project workflow automation, roadmap/state parsing, and commit orchestration.

## Key Dependencies

**Critical:**
- `openpyxl` (vendored in repo) - Excel export functionality backed by `lib/openpyxl/**` and imported in `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`.
- `et_xmlfile` (vendored transitive dependency) - present in `lib/et_xmlfile/**`.
- Revit API assemblies (`RevitAPI`, `RevitAPIUI`, `RevitServices`) - loaded in scripts such as `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Высота этажа.pushbutton/script.py`.

**Infrastructure:**
- Node.js built-ins (`fs`, `path`, `child_process`) used by GSD internals in `.github/get-shit-done/bin/lib/core.cjs` and `.github/get-shit-done/bin/lib/commands.cjs`.
- GSD internal version marker `1.32.0` in `.github/get-shit-done/VERSION`.

## Configuration

**Environment:**
- pyRevit command discovery/config is file-based (`bundle.yaml`, `extension.json`) with no root `.env` detected.
- Optional API-key-driven capabilities in GSD CLI are read from environment variables and `~/.gsd/*` key files (see `.github/get-shit-done/bin/lib/config.cjs`).

**Build:**
- No dedicated build pipeline config detected (no root CI workflow files, no Dockerfiles, no bundler configs).
- Extension behavior is interpreted at runtime by pyRevit/Revit rather than prebuilt artifacts.

## Platform Requirements

**Development:**
- Windows workstation with Autodesk Revit and pyRevit available.
- Python runtime compatible with pyRevit execution modes (IronPython and CPython3 depending on tool).
- Node.js available to execute `.github/get-shit-done/bin/gsd-tools.cjs` commands.

**Production:**
- Deployment target is local desktop Revit environment (not server-hosted).
- Artifacts are source files loaded directly by pyRevit from the extension directory.

---

*Stack analysis: 2026-06-08*
