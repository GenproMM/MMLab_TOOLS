---
quick_id: 260819-cad
status: complete
---

# Quick Task 260819-cad: приёмка кнопки «Замена CAD-геометрии»

## Task
Принять сторонний скрипт `Замена CAD-геометрии.pushbutton` (панель ИОС, уже
присутствовал в дереве и был зарегистрирован в layout) в конвенцию MM LAB через
`/mm-adopt-script`. Кнопка работает в редакторе семейств: заменяет DWG-импорты на
лёгкие нативные боксы, перенося коннекторы (тип, профиль, параметры потока и
классификации системы) на новую геометрию.

## Adaptation summary
- MM001/MM002: добавлены `#! python3` и `# -*- coding: utf-8 -*-` первыми строками.
- MM004: `__doc__`-строка заменена на модульный docstring с
  «Совместимость: Revit 2020 / 2022 / 2024» и «Зависимости: нет».
- MM014 (D-15): добавлен канонический lib-бутстрап
  (`_SCRIPT_DIR`/`_EXTENSION_DIR`/`_LIB_DIR`), которого не было вовсе.
- MM011: все 27 голых `except:` заменены на `except Exception:` без изменения
  логики (`except: pass` / `except: return ...` — поведение сохранено).
- Правило 7 AGENTS.md: 14 прямых вызовов `ce.get_Parameter(BuiltInParameter...)` /
  `ce_new.get_Parameter(...)` заменены на `revit_compat.get_parameter(ce, ...)` /
  `revit_compat.get_parameter(ce_new, ...)` — устраняет падения на pythonnet 3.x
  (Revit 2022/2024) при явном касте enum в get_Parameter.
- Правило 9/18 AGENTS.md (D-03, D-18): вся логика обёрнута в `main(doc)` с
  `revit_compat.require_supported_version(COMMAND_NAME)` первой строкой;
  `uidoc`/`doc` получаются в новой `_entry()` и передаются в `main` параметром
  (было: модульные глобалы `uidoc = __revit__.ActiveUIDocument`, `doc = uidoc.Document`).
- Правило 11 AGENTS.md: добавлена внешняя обёртка
  `try: _entry() except SystemExit: pass except Exception as ex: TaskDialog.Show(...)`.
- Транзакционный каркas (Start до try, Commit в try, RollBack+raise в except) уже
  был правильным в исходнике — сохранён без изменений.
- MM006: создан `README.md` по разделам шаблона.
- Запись кнопки удалена из `units` в `tools/convention_baseline.json`.

## Verify
```bash
py -3 tools/check_convention.py "MM_LAB.extension/MM Lab.tab/ИОС.panel/Замена CAD-геометрии.pushbutton" --strict
py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json
```
Первая команда — exit 0, 0 ошибок, 0 предупреждений. Вторая по-прежнему возвращает
ошибки, но исключительно от несвязанных, ранее не принятых кнопок
(`СНиП.pushbutton`, `СНиП_ФОП25.pushbutton` на панели АРХИТЕКТУРА, закоммичены
11.08.2026) — предсуществующий пробел вне рамок этой приёмки, кнопка «Замена
CAD-геометрии» в списке нарушений не участвует.
