# /mm-new-compat — добавить ветку новой версии Revit в revit_compat.py

Расширение матрицы поддерживаемых версий Revit. Единственное место версионных
ветвлений в репозитории — `MM_LAB.extension/lib/revit_compat.py` (D-01);
процедура обновляет его, стандарт и контрактный тест синхронно.

## Аргументы

- `<версия>` — новая версия Revit, целое число (например `2025`).
  Если не передана — спроси: «Какую версию Revit добавляем в поддержку?»

## Процедура

1. **Прочитай** `MM_LAB.extension/lib/revit_compat.py` целиком — карта
   версионных веток и публичный API из 13 функций.

2. **Добавь версию в `SUPPORTED_VERSIONS`** — кортеж в начале модуля,
   по возрастанию (например `(2020, 2022, 2024, 2025)`).

3. **Проревизуй версионные ветки** на предмет ломающих изменений добавляемой
   версии — сейчас их три:
   - `_units_map()` — карта единиц: `Revit <= 2020` = `DisplayUnitType`,
     новее = `UnitTypeId` (ForgeTypeId);
   - `create_floor()` — `Revit <= 2020` = `doc.Create.NewFloor(CurveArray, ...)`,
     новее = `Floor.Create(doc, IList[CurveLoop], ...)`;
   - `make_element_id()` / `element_id_value()` — Int64/`.Value` (2024+)
     против `.IntegerValue` (старые версии).

   Выясни, ломает ли добавляемая версия что-то ещё: спроси пользователя
   и/или исследуй официальную документацию «What's New in the Revit <версия> API»
   (Autodesk). Для 2025+ обязательно учти переход Revit на .NET Core
   (поведение pythonnet/движка pyRevit может отличаться). Найденные ломающие
   изменения закрывай новыми ветками/хелперами ТОЛЬКО внутри `revit_compat.py`.

4. **Обнови три точки синхронно** (иначе контрактный тест упадёт):
   - `MM_LAB.extension/lib/revit_compat.py` — docstring модуля: строка
     «Совместимость: Revit … / <версия>»;
   - `AGENTS.md` — таблица «Совместимость Revit» и упоминания
     `SUPPORTED_VERSIONS` в тексте стандарта (включая строку «Совместимость:»
     в каноне шапки, шаблон `templates/НоваяКнопка.pushbutton` и маркер
     docstring MM004, если формат строки меняется);
   - `tools/tests/test_revit_compat_contract.py` — ожидаемый кортеж
     в `test_supported_versions` (и docstring теста).

5. **Прогони проверки:**

   ```bash
   py -3 -m py_compile "MM_LAB.extension/lib/revit_compat.py"
   py -3 -m unittest discover -s tools/tests -p "test_revit_compat*.py" -q
   ```

   После `py_compile` удали созданную папку `__pycache__` рядом с модулем
   (мусор; в `.gitignore` покрыт, но чистота дороже).

6. **Коммит пофайловый** (D-17), push — с подтверждения (D-18):

   ```bash
   git -c core.quotepath=false status --porcelain
   git add "MM_LAB.extension/lib/revit_compat.py"
   git add "AGENTS.md"
   git add "tools/tests/test_revit_compat_contract.py"
   git commit -m "compat: добавлена поддержка Revit <версия>"
   ```

## Гейты и запреты

- Версионные ветки — ТОЛЬКО внутри `revit_compat.py` (D-01): ветвления
  `if version >= ...` в скриптах кнопок запрещены конвенцией.
- Не выдумывай ломающие изменения по памяти — сверяйся с официальной
  документацией Autodesk и/или решением пользователя.
- Публичный API compat (13 функций) не сужать и не переименовывать —
  он зафиксирован контрактным тестом `test_revit_compat_contract.py`.
- `git add .` / `git add -A` запрещены (D-17).

## Финал

Сообщи пользователю:

- какая версия добавлена и какие ветки/хелперы изменены;
- все три точки обновлены: `revit_compat.py`, `AGENTS.md`,
  `tools/tests/test_revit_compat_contract.py`;
- `py_compile` и unittest — зелёные;
- ВАЖНО: чекер и тесты статические — изменения требуют Revit UAT
  на новой версии (запустить кнопки на реальной модели в Revit <версия>);
- коммит создан, push не выполнялся (или выполнен по подтверждению).
