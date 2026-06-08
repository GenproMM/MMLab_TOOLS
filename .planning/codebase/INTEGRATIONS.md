# External Integrations

**Analysis Date:** 2026-06-08

## APIs & External Services

**Design/BIM Host APIs:**
- Autodesk Revit API - primary automation surface for element querying, transactions, parameters, and UI dialogs.
  - SDK/Client: .NET assemblies via `clr.AddReference("RevitAPI")` and `clr.AddReference("RevitAPIUI")` in files such as `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Мокрые зоны.pushbutton/script.py` and `MM LAB.extension/MM Lab.tab/ИОС.panel/Доп расход 0.pushbutton/script.py`.
  - Auth: Host-process trust model inside running Revit session (no explicit token flow in repository scripts).
- pyRevit runtime services - UI helpers and script output/forms wrappers.
  - SDK/Client: `from pyrevit import forms, script, revit` in `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py` and `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Проверка зон.pushbutton/script.py`.

**Developer Workflow Services (GSD CLI):**
- Brave Search API - optional web search provider for planning/research commands.
  - SDK/Client: native `fetch` call to `https://api.search.brave.com/res/v1/web/search` in `.github/get-shit-done/bin/lib/commands.cjs`.
  - Auth: `BRAVE_API_KEY` via `X-Subscription-Token` header in `.github/get-shit-done/bin/lib/commands.cjs`.
- Firecrawl / Exa capability detection - optional feature flags for tooling readiness.
  - SDK/Client: presence checks only in `.github/get-shit-done/bin/lib/config.cjs`.
  - Auth: `FIRECRAWL_API_KEY`, `EXA_API_KEY` or corresponding `~/.gsd/*` key files.

## Data Storage

**Databases:**
- Not detected.
  - Connection: Not applicable.
  - Client: Not applicable.

**File Storage:**
- Local filesystem only.
- Export outputs include CSV/XLSX writes from pyRevit scripts, e.g. CSV save path logic in `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Мокрые зоны.pushbutton/script.py` and workbook generation via `openpyxl` in `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`.
- Project planning state stored as Markdown/JSON under `.planning/**` and managed by `.github/get-shit-done/bin/gsd-tools.cjs`.

**Caching:**
- None detected in extension code.
- No Redis/memcached client imports in project source.

## Authentication & Identity

**Auth Provider:**
- Revit/pyRevit session context for extension commands.
  - Implementation: scripts access active document through `__revit__.ActiveUIDocument` in multiple tools such as `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Высота этажа.pushbutton/script.py`.
- For optional GSD web search, simple API-key authentication (`BRAVE_API_KEY`) in `.github/get-shit-done/bin/lib/commands.cjs`.

## Monitoring & Observability

**Error Tracking:**
- None detected (no Sentry/New Relic/AppInsights SDK usage).

**Logs:**
- User-facing runtime feedback is dialog-based through Revit `TaskDialog` or WinForms dialogs in scripts like `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`.
- CLI output is stdout/stderr JSON/text from `.github/get-shit-done/bin/gsd-tools.cjs`.

## CI/CD & Deployment

**Hosting:**
- Not server-hosted; deployed as a local pyRevit extension directory (`MM LAB.extension/**`).

**CI Pipeline:**
- None detected (`.github/workflows/**` absent in repository scan).

## Environment Configuration

**Required env vars:**
- `BRAVE_API_KEY` for GSD `websearch` command in `.github/get-shit-done/bin/lib/commands.cjs`.
- Optional integration toggles in `.github/get-shit-done/bin/lib/config.cjs`: `FIRECRAWL_API_KEY`, `EXA_API_KEY`.
- GSD operational context variables (non-secret) used by CLI: `GSD_WORKSTREAM`, `GSD_PROJECT`, `GSD_AGENTS_DIR` in `.github/get-shit-done/bin/gsd-tools.cjs` and `.github/get-shit-done/bin/lib/core.cjs`.

**Secrets location:**
- Environment variables at runtime.
- Optional local key files in user home under `~/.gsd/` as referenced by `.github/get-shit-done/bin/lib/config.cjs`.

## Webhooks & Callbacks

**Incoming:**
- None detected.

**Outgoing:**
- None in production extension scripts.
- Optional outbound HTTPS call to Brave Search API from `.github/get-shit-done/bin/lib/commands.cjs` when `websearch` is invoked.

---

*Integration audit: 2026-06-08*
