# Codebase Concerns

**Analysis Date:** 2026-06-08

## Tech Debt

**Silent exception swallowing across core tools:**
- Issue: Extensive `except: pass` and broad `except` blocks hide runtime failures and can leave partial state updates without operator visibility.
- Files: `MM LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/КООРДИНАЦИЯ.panel/ВерсияШаблона.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Приточный по классификации.pushbutton/script.py`
- Impact: Failed API calls are silently ignored; model cleanup and parameter writes can complete with hidden data gaps.
- Fix approach: Replace bare `except` with typed exceptions, log failure counts to pyRevit output, and fail transaction on critical mutation errors.

**High copy-paste duplication in IOS tool scripts:**
- Issue: Multiple large scripts share near-identical code blocks (same imports/constants/helpers around the first ~120 lines and recurring logic blocks at the same offsets).
- Files: `MM LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Приточный по классификации.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Конфузор-Диффузор.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Доп расход 0.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Доп расход 1.pushbutton/script.py`
- Impact: Bug fixes must be repeated in several files; behavior drift and regressions are likely.
- Fix approach: Extract shared helpers into a common module under the extension and keep each button script as thin command orchestration.

**Vendored dependency management drift risk:**
- Issue: Spreadsheet export depends on manually managed vendored packages in `lib/` and modifies `sys.path` at runtime.
- Files: `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Экспорт ПСО.pushbutton/script.py`, `lib/openpyxl-3.1.5.dist-info/top_level.txt`, `lib/et_xmlfile-2.0.0.dist-info/top_level.txt`
- Impact: Environment inconsistencies across developers and machines; hard-to-diagnose import/runtime mismatches.
- Fix approach: Pin and verify vendor package set with a reproducible bootstrap script and startup validation command.

## Known Bugs

**Partial cleanup can report success despite failed deletions:**
- Symptoms: CAD cleanup routines continue after delete failures and suppress exceptions; reported counters may not reflect true state.
- Files: `MM LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/script.py`
- Trigger: Any delete call failure inside loops around category/unused element cleanup paths with `except: pass`.
- Workaround: Re-run the tool and manually inspect resulting model state in Revit after each run.

**Template version write can silently skip target entities:**
- Symptoms: Version update executes but updates count may be zero when parameter is unavailable/read-only on candidates.
- Files: `MM LAB.extension/MM Lab.tab/КООРДИНАЦИЯ.panel/ВерсияШаблона.pushbutton/script.py`
- Trigger: Missing binding/read-only parameter scenarios in project info or sheet fallback paths.
- Workaround: Validate parameter binding and writable scope before command execution.

## Security Considerations

**High-impact destructive operations without confirmation gates:**
- Risk: Mass `doc.Delete(...)` actions can remove imports/types/elements in bulk.
- Files: `MM LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/script.py`
- Current mitigation: Transactions with rollback are present, but many inner failures are swallowed.
- Recommendations: Add explicit preflight summary + user confirmation, dry-run mode, and strict abort on critical delete failures.

## Performance Bottlenecks

**Repeated full-model scans in large Revit projects:**
- Problem: Multiple tools iterate full collectors and connector graphs; some scripts chain several collectors in one run.
- Files: `MM LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Приточный по классификации.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Конфузор-Диффузор.pushbutton/script.py`
- Cause: Repeated `FilteredElementCollector(...).WhereElementIsNotElementType()` loops and per-element connector traversal.
- Improvement path: Cache intermediate sets, narrow collectors by category/filter earlier, and avoid repeated recomputation across steps.

**Quadratic overlap checks during CAD replacement:**
- Problem: Bounding box pair checks perform nested iteration over imported geometries.
- Files: `MM LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/script.py`
- Cause: Nested loops over `bboxes` index pairs.
- Improvement path: Spatial indexing (grid/bucket/R-tree-like partitioning) before pairwise checks.

## Fragile Areas

**Large monolithic scripts with mixed UI, selection, and mutation logic:**
- Files: `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/script.py`, `MM LAB.extension/MM Lab.tab/ИОС.panel/Приточный по классификации.pushbutton/script.py`
- Why fragile: 498-717 line files combine data collection, business logic, dialogs, and transactions; small edits can affect unrelated behavior.
- Safe modification: Isolate helpers first, then change one transaction path at a time with explicit post-check output.
- Test coverage: No automated tests detected for these scripts.

## Scaling Limits

**Operational limit tied to model size and element count:**
- Current capacity: Not explicitly bounded in code; runtime scales with full-model scans and connector traversal.
- Limit: User-facing latency and timeout risk on large federated models.
- Scaling path: Introduce scoped execution (active view/selection/discipline filters) and phased processing with progress checkpoints.

## Dependencies at Risk

**Bundled `openpyxl` stack in repo:**
- Risk: Local vendored dependencies can drift from runtime interpreter constraints and lack automated verification.
- Impact: Export functionality breaks at import or workbook serialization time.
- Migration plan: Centralize dependency bootstrap/validation and add startup self-check in export command.

## Missing Critical Features

**No centralized logging/telemetry for mutation failures:**
- Problem: Commands rely mostly on dialogs and silent exception paths; failures are not consistently captured.
- Blocks: Reliable root-cause analysis and confidence in model-wide mutation tools.

## Test Coverage Gaps

**No automated test suite for extension scripts:**
- What's not tested: Parameter write flows, delete/purge behavior, and classification decisions in IOS/architecture tools.
- Files: `MM LAB.extension/**/*.py`
- Risk: Regressions in model mutation logic go undetected before user execution.
- Priority: High

---

*Concerns audit: 2026-06-08*
