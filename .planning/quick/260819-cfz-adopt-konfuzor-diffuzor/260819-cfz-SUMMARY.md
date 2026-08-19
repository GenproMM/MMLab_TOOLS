---
quick_id: 260819-cfz
status: complete
---

# Summary: приёмка кнопки «Конфузор-Диффузор»

Кнопка `Конфузор-Диффузор` (панель ИОС) адаптирована под конвенцию MM LAB. Логика
классификации переходов (сравнение площадей входного/выходного HVAC-коннекторов,
запись в параметр «Конфузор») не менялась — исправлены только шапка файла
(BOM/docstring), lib-бутстрап (канонический блок вместо кастомного
`EXTENSION_ROOT`), доступ к `__revit__` (перенесён из lib-хелпера `get_document`
в локальный `_entry()` скрипта, правило 18), добавлен fail-fast гейт версии
Revit, `__author__` приведён к `"GENPRO LAB"`, добавлены `bundle.yaml` (en_us
локаль) и `README.md`.

`--strict` на кнопке — exit 0, 0 ошибок, 0 предупреждений. Запись кнопки удалена
из `units` в `tools/convention_baseline.json`. Полный прогон `--all --baseline`
по-прежнему возвращает ошибки, но только от несвязанных, ранее не принятых
кнопок (`СНиП`, `СНиП_ФОП25` на панели АРХИТЕКТУРА) — предсуществующий пробел вне
рамок этой приёмки.

Files changed:
- `MM_LAB.extension/MM Lab.tab/ИОС.panel/Конфузор-Диффузор.pushbutton/script.py`
- `MM_LAB.extension/MM Lab.tab/ИОС.panel/Конфузор-Диффузор.pushbutton/bundle.yaml`
- `MM_LAB.extension/MM Lab.tab/ИОС.panel/Конфузор-Диффузор.pushbutton/README.md`
- `tools/convention_baseline.json`
