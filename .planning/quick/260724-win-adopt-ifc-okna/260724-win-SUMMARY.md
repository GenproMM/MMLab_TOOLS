---
quick_id: 260724-win
status: complete
---

# Summary: приёмка кнопки IFC_Окна

Сторонний скрипт классификации окон по МССК принят и адаптирован под
конвенцию MM LAB на панель **АРХИТЕКТУРА**. Кнопка `IFC_Окна` заполняет
общие параметры `GP_01_КодКлассифМССК` / `GP_01_ИмяКлассифМССК`: балконные
блоки → `ЭЛ 30 18 09` «Балконный блок», остальные → `ЭЛ 30 18 40` «Окно».

Приведено к конвенции: шапка (MM001–MM004), явные импорты (MM009),
WinForms вместо `pyrevit.forms` (MM012), параметры через `revit_compat`
(MM010), `except Exception` (MM011), канонический lib-бутстрап (MM014),
`bundle.yaml` + `README.md` (MM005/MM006), fail-fast по версии (D-03),
транзакционный каркас (правило 10), `doc` параметром (правило 18),
`element_id_value` вместо `.IntegerValue`. Кнопка зарегистрирована в
layout панели (MM007), запись из `pending_adoption` снята.

Строгий чекер и общий прогон с baseline — exit 0.

Файлы:
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Окна.pushbutton/script.py`
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Окна.pushbutton/bundle.yaml`
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Окна.pushbutton/README.md`
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Окна.pushbutton/icon.png`
- `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/bundle.yaml`
- `tools/convention_baseline.json`
