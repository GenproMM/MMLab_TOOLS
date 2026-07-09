---
quick_id: 260709-jko
status: complete
---

# Quick Task 260709-jko: ИОС panel button errors

## Task
Buttons on the ИОС panel throw errors on click:
1. "Доп. расход = 1" → `No method matches given arguments for get_Parameter: (<class 'int'>)`
2. "Сброс потерь" → `type object 'BuiltInParameter' has no attribute 'RBS_DUCT_LOSS_METHOD_SERVER_PARAM'`

## Root cause
- Error 1 was already fixed by prior commits (0acdac0, 7327da9, d207200) which moved
  `set_additional_flow_value` off `get_Parameter(BuiltInParameter)` onto
  `LookupParameter(string)`. Current code no longer reproduces this error.
- Error 2 is live: `get_loss_method_parameters()` in `lib/ios_common_helpers.py`
  referenced two BuiltInParameter names that do not exist in the Revit API —
  `RBS_DUCT_LOSS_METHOD_SERVER_PARAM` and `RBS_DUCT_TERMINAL_LOSS_METHOD_SERVER_PARAM`.
  Only `RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM` is real, and per Autodesk docs it
  covers both duct fittings and duct accessories — the two categories the
  "Сброс потерь" button actually targets.
- Secondary bug found during investigation: `ensure_loss_method_undefined` returned a
  bool, but the caller (`Сброс потерь/script.py`) compared the result against the
  strings `"updated"`/`"already"`. Every run silently reported everything as
  "Пропущено" regardless of actual outcome.

## Fix
`MM LAB.extension/lib/ios_common_helpers.py`:
- `get_loss_method_parameters`: drop the two nonexistent BuiltInParameter entries,
  keep only `RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM`.
- `ensure_loss_method_undefined`: return `"updated"` / `"already"` / `"skipped"`
  matching the caller's expected tri-state contract.

## Verify
- `python -c "import ast; ast.parse(open('MM LAB.extension/lib/ios_common_helpers.py', encoding='utf-8').read())"`
- Manual: click "Доп. расход = 1/0" and "Сброс потерь" in Revit — no error dialog,
  correct counts reported.
