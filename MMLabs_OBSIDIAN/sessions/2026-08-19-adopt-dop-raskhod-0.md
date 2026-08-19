# Сессия 2026-08-19: Приёмка кнопки «Доп. расход 0»

## Что сделано

Адаптирована кнопка ИОС.panel/Доп расход 0 под конвенцию MM LAB. Кнопка уже физически лежала в репозитории с нарушениями MM003, MM004, MM006, MM014.

### Адаптационные правки

1. **MM003: UTF-8 BOM** — убран BOM в начале script.py.
2. **MM004: docstring** — добавлен русский docstring со строками «Совместимость: Revit 2020 / 2022 / 2024» и «Зависимости: нет».
3. **MM006: README.md** — создан README с разделами Описание, Логика работы, Используемые параметры, Зависимости, Совместимость, Ограничения.
4. **MM014: lib-бутстрап** — заменён на канонический блок `_SCRIPT_DIR`/`_EXTENSION_DIR`/`_LIB_DIR` с `sys.path.insert(0, ...)`.
5. **Правило 18** — `__revit__`/`uidoc`/`doc` теперь получаются в script.py и передаются в `main(doc)` параметром (вместо чтения `__revit__` в lib-модуле через `get_document`).
6. **Правило 9 / D-03** — добавлен `revit_compat.require_supported_version(COMMAND_NAME)` в начало `main()`.
7. **bundle.yaml** — добавлена en_us локаль (title/tooltip), `author` синхронизирован с «GENPRO LAB» (стандарт репозитория).

### Проверки

- Строгий чекер (`--strict`): ✅ exit 0, 0 ошибок, 0 предупреждений.
- Запись удалена из `tools/convention_baseline.json` (units).
- `.planning/STATE.md` обновлён (добавлена запись quick task).
- Quick task артефакты созданы в `.planning/quick/260819-adp-adopt-dop-raskhod-0/`.

## Затронутые файлы

- `MM_LAB.extension/MM Lab.tab/ИОС.panel/Доп расход 0.pushbutton/script.py` (изменён)
- `MM_LAB.extension/MM Lab.tab/ИОС.panel/Доп расход 0.pushbutton/bundle.yaml` (изменён)
- `MM_LAB.extension/MM Lab.tab/ИОС.panel/Доп расход 0.pushbutton/README.md` (создан)
- `tools/convention_baseline.json` (запись удалена из units)
- `.planning/STATE.md` (добавлена запись quick task)
- `.planning/quick/260819-adp-adopt-dop-raskhod-0/260819-adp-PLAN.md` (создан)
- `.planning/quick/260819-adp-adopt-dop-raskhod-0/260819-adp-SUMMARY.md` (создан)

## Ключевые выводы

1. **Процедура `/mm-adopt-script`** хорошо работает для адаптации уже существующих кнопок.
2. **Baseline grandfathering** — правильно удаляет записи legacy-кнопок при их приёмке.
3. **lib-бутстрап D-15** — канонический блок обязателен, легаси-форма больше не используется.
4. **Правило 18** — мода на передачу `doc`/`uidoc` параметрами вместо чтения `__revit__` в lib — требует перестройки скрипта, но даёт лучшую инжекцию.

## Следующие шаги

- pyRevit Reload в Revit.
- UAT в реальном Revit (проверка кнопки на панели ИОС).
