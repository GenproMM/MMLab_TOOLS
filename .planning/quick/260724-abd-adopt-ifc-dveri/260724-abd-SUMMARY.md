---
quick_id: 260724-abd
status: complete
---

# Summary: приёмка кнопки IFC_Двери

Кнопка `IFC_Двери` (панель АРХИТЕКТУРА) адаптирована под конвенцию MM LAB и
зарегистрирована в layout. Классификация дверей по МССК не менялась по сути —
исправлен только доступ к параметрам (GUID общих параметров GP_01/GP_23 через
`revit_compat.get_shared_parameter`, встроенный «Модель» через `get_parameter`),
убраны `pyrevit.forms`, wildcard-импорт, голые `except:`, добавлен fail-fast по
версии Revit и канонический lib-бутстрап.

`--strict` на кнопке и полный прогон `--all --baseline` — оба exit 0, 0 ошибок,
0 предупреждений. Запись кнопки удалена из `pending_adoption` в
`tools/convention_baseline.json`.

Files changed:
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Двери.pushbutton/script.py`
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Двери.pushbutton/bundle.yaml`
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Двери.pushbutton/README.md`
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Двери.pushbutton/icon.png`
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/bundle.yaml`
- `tools/convention_baseline.json`
