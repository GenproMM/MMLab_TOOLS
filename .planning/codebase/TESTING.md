# Testing Patterns

**Analysis Date:** 2026-06-08

## Test Framework

**Runner:**
- Not detected for first-party extension code.
- Config: Not detected (`pytest.ini`, `pyproject.toml`, `tox.ini`, `setup.cfg`, `jest.config.*`, `vitest.config.*` are absent in repository root scan).

**Assertion Library:**
- Not detected.

**Run Commands:**
```bash
Not detected              # Run all tests
Not detected              # Watch mode
Not detected              # Coverage
```

## Test File Organization

**Location:**
- No automated test directories detected for first-party command code under `MM LAB.extension`.
- No files matching `*.test.py` or `*test*.py` were detected under `MM LAB.extension`.

**Naming:**
- Production command entry files use `script.py` naming in button folders, e.g. `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/script.py`.

**Structure:**
```
MM LAB.extension/
  MM Lab.tab/
    <Panel>.panel/
      <Command>.pushbutton/
        bundle.yaml
        script.py
        README.md (optional)
```

## Test Structure

**Suite Organization:**
```typescript
// Not applicable: no TypeScript/Jest-style suites in this repository's first-party code.
```

**Patterns:**
- Setup pattern: interactive Revit context setup (`__revit__.ActiveUIDocument`) before command execution in `MM LAB.extension/MM Lab.tab/КООРДИНАЦИЯ.panel/ВерсияШаблона.pushbutton/script.py`.
- Teardown pattern: transaction rollback on exception (`transaction.RollBack()`) in `MM LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py`.
- Assertion pattern: runtime guard conditions plus UI confirmation (`TaskDialog.Show(...)`) instead of unit assertions.

## Mocking

**Framework:**
- Not detected.

**Patterns:**
```typescript
// Not applicable: automated mocks are not present.
```

**What to Mock:**
- For future automated tests, mock Revit API wrappers around document/collector/transaction boundaries.

**What NOT to Mock:**
- Do not mock business rules that parse and transform parameter values (`parse_revision_date`, classification normalization); test those as pure functions.

## Fixtures and Factories

**Test Data:**
```typescript
// Not applicable in current repository state.
```

**Location:**
- No fixture/factory directories detected for first-party code.

## Coverage

**Requirements:**
- None enforced.

**View Coverage:**
```bash
Not applicable
```

## Test Types

**Unit Tests:**
- Not used in first-party extension scripts at this time.

**Integration Tests:**
- Informal runtime integration occurs inside Revit through command execution with live model data.
- Evidence: transaction-based execution and user dialogs in `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/Высота этажа.pushbutton/script.py` and `MM LAB.extension/MM Lab.tab/ИОС.panel/Сброс потерь.pushbutton/script.py`.

**E2E Tests:**
- No dedicated E2E framework detected.
- Practical E2E is manual UAT through pyRevit button invocation and report dialogs, documented in `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/ВОР_СсылкаНаЛист.pushbutton/README.md`.

## Common Patterns

**Async Testing:**
```typescript
// Not applicable: scripts are synchronous Revit command handlers.
```

**Error Testing:**
```typescript
// Current pattern is runtime try/except with rollback and TaskDialog:
// try: transaction.Start(); ...; transaction.Commit()
// except: transaction.RollBack(); TaskDialog.Show(...)
```

---

*Testing analysis: 2026-06-08*
