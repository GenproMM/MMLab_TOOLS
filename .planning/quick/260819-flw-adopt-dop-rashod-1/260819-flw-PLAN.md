---
quick_id: 260819-flw
status: complete
---

# Quick Task 260819-flw: приёмка кнопки «Доп. расход 1»

## Task
Принять сторонний скрипт `Доп расход 1.pushbutton` (панель ИОС, уже присутствовал
в дереве и был зарегистрирован в layout) в конвенцию MM LAB через `/mm-adopt-script`.
Кнопка устанавливает значение 1 в общий параметр «Доп. расход» (`RBS_ADDITIONAL_FLOW`)
у элементов воздуховодной сети. Логика скрипта по сути не менялась — только приведение
к конвенции, зеркалит уже принятую сестринскую кнопку «Доп расход 0».

## Adaptation summary
- MM003: файл пересохранён в UTF-8 без BOM.
- MM004: добавлен docstring модуля с «Совместимость: Revit 2020 / 2022 / 2024»
  и «Зависимости: нет».
- MM014: неканонический lib-бутстрап (`EXTENSION_ROOT`, подъём через
  `dirname(dirname(dirname(...)))`, `sys.path.append`) → канонический блок
  `_SCRIPT_DIR`/`_EXTENSION_DIR`/`_LIB_DIR` с `sys.path.insert(0, _LIB_DIR)` (D-15).
- Правило 18 AGENTS.md: `ios_common_helpers.get_document`/`show_error` тянули
  `__revit__` внутри lib-модуля → заменено на `_entry()` в script.py, получающий
  `uidoc`/`doc` через `__revit__.ActiveUIDocument` и передающий `doc` в `main(doc)`
  параметром.
- Правило 9 (D-03): добавлен `revit_compat.require_supported_version(COMMAND_NAME)`
  в начале `main()`.
- `__author__` приведён к `"GENPRO LAB"` (как у остальных принятых кнопок панели).
- `bundle.yaml`: добавлены `en_us` title/tooltip, `author` → `"GENPRO LAB"`.
- MM006: создан `README.md` по разделам шаблона.
- Запись кнопки удалена из `units` в `tools/convention_baseline.json`.

## Verify
```bash
py -3 tools/check_convention.py "MM_LAB.extension/MM Lab.tab/ИОС.panel/Доп расход 1.pushbutton" --strict
py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json
```
Первая команда — exit 0, 0 ошибок, 0 предупреждений. Вторая по-прежнему возвращает
ошибки, но исключительно от несвязанных, ранее не принятых кнопок
(`СНиП.pushbutton`, `СНиП_ФОП25.pushbutton` на панели АРХИТЕКТУРА) — предсуществующий
пробел вне рамок этой приёмки, кнопка «Доп расход 1» в списке нарушений не участвует.
