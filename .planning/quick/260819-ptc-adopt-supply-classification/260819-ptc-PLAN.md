---
quick_id: 260819-ptc
status: complete
---

## Task

Приёмка кнопки «Приточный по классификации» из папки `MM_LAB.extension/MM Lab.tab/ИОС.panel/Приточный по классификации.pushbutton/` в соответствие с конвенцией MM LAB. Исходный скрипт имел разнобой имён (папка/bundle/docstring), устаревший lib-бутстрап, UTF-8 BOM, отсутствовал README.md. 

Процедура: статическая проверка чекером → адаптация → ревью-гейт → регистрация → коммит.

## Adaptation summary

**script.py (MM-коды нарушений):**
- MM003: Убран UTF-8 BOM с начала файла
- MM004: Добавлен настоящий docstring модуля с «Совместимость:» и «Зависимости:»
- MM014: Заменён нестандартный lib-бутстрап (`SCRIPT_DIR`/`EXTENSION_ROOT`/`sys.path.append`) на канонический (`_SCRIPT_DIR`/`_EXTENSION_DIR`/`sys.path.insert(0, ...)`)
- Синхронизирован `__title__` с именем папки и panel-подписью: было «Вытяжка» (чужое имя), стало «Приточный по классификации»
- `COMMAND_NAME` синхронизирован с `__title__`
- `__author__` → «GENPRO LAB»
- Логика вынесена в `main(doc)` + `_entry()`: doc/uidoc получаются в script.py и передаются параметром (правило 18) — раньше doc тянулся из lib
- В начало `main()` добавлен `revit_compat.require_supported_version()` (fail-fast по версии)
- Отмена диалога теперь `return` вместо `raise SystemExit`
- Убраны неиспользуемые импорты `get_document`/`show_error` из ios_common_helpers

**bundle.yaml:**
- Добавлены `en_us` локализации
- Русский title синхронизирован с новым именем
- author → «GENPRO LAB»

**README.md (MM006):** Создан с разделами по шаблону (Описание, Логика работы, Используемые параметры, Зависимости, Совместимость, Ограничения).

**tools/convention_baseline.json:**
- Удалена запись кнопки из `units` (legacy-допуск) — дальше кнопка держит `--strict` без поблажек

**Результат:** Strict check pass (0 ошибок, 0 предупреждений), полный репо-прогон с baseline pass.
