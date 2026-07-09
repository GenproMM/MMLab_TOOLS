---
quick_id: 260709-jko
status: complete
---

# Summary: ИОС panel button errors

Fixed the crash on "Сброс потерь": `BuiltInParameter.RBS_DUCT_LOSS_METHOD_SERVER_PARAM`
and `RBS_DUCT_TERMINAL_LOSS_METHOD_SERVER_PARAM` do not exist in the Revit API — removed,
kept the one valid parameter `RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM` (covers both
duct fittings and accessories per Autodesk docs).

Also fixed `ensure_loss_method_undefined` returning a bool while its caller compared
against string outcomes ("updated"/"already") — every run was silently misreported as
"Пропущено". Now returns the correct tri-state string.

"Доп. расход = 1/0" button error was already fixed by prior commits (0acdac0, 7327da9,
d207200); current code no longer calls `get_Parameter(BuiltInParameter)`.

Files changed: `MM LAB.extension/lib/ios_common_helpers.py`
