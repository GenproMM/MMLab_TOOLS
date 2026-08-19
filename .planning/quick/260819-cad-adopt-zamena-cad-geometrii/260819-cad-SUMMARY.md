---
quick_id: 260819-cad
status: complete
---

# Summary: приёмка кнопки «Замена CAD-геометрии»

Адаптировали кнопку `Замена CAD-геометрии.pushbutton` (панель ИОС) под конвенцию
MM LAB: канонические шапка/докстрока, lib-бутстрап, `main(doc)`/`_entry()`
с `revit_compat.require_supported_version`, все голые `except:` → `except Exception:`,
чтение/запись параметров коннекторов переведены на `revit_compat.get_parameter`,
добавлен `README.md`. Геометрическая логика (боксы, tiny-экструзии, поиск грани,
purge unused) не менялась.

`--strict` по кнопке — 0 ошибок / 0 предупреждений. Запись кнопки удалена из
`units` в `tools/convention_baseline.json`.

Files changed:
- `MM_LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/script.py`
- `MM_LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton/README.md` (новый)
- `tools/convention_baseline.json`
