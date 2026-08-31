---
quick_id: 260831-wal
status: complete
---

# Quick Task 260831-wal: приёмка кнопки IFC_Стены

## Task
Принять и адаптировать сторонний скрипт классификации стен по МССК под
конвенцию MM LAB (процедура `/mm-adopt-script`).

- Источник: `G:\Общие диски\10_BIM_LAB\8_DevOps\2_WIP\2_PyRevit\MMLab_TOOLS\LL.extension\MM Lab.tab\АРХИТЕКТУРА.panel\IFC_Стены.pushbutton\script.py`
  (байт-в-байт копия уже лежала в рабочем дереве untracked как
  `IFC_Стены.pushbutton/script.py` + `icon.png`, md5 `00511ed1436d1fa47c132b2fa1ea24fb`;
  записи в `convention_baseline.json` не было).
- Панель: **АРХИТЕКТУРА** (`Wall` / `WallType` / `WallKind.Curtain`,
  `GetRoomAtPoint`, `ROOM_NAME`, IFC-классификация) — подтверждено пользователем.
- Аналог уже принятых кнопок `IFC_Двери`, `IFC_Окна`, `IFC_Перекрытия`
  (те же общие параметры `GP_01_КодКлассифМССК` / `GP_01_ИмяКлассифМССК`,
  тот же WinForms-паттерн выбора типоразмеров).
- Отчёт чекера по исходнику: 2 error, 12 warning — MM001, MM002, MM004,
  MM012, MM010 ×2, MM011 ×8.

## Adaptation summary
- **MM001/MM002**: добавлены `#! python3` и `# -*- coding: utf-8 -*-`.
- **MM003**: файл записан UTF-8 без BOM, переводы строк LF (в исходнике CRLF).
- **MM004**: русский docstring со строками «Совместимость:» и «Зависимости:»;
  опечатка `__autor__` исправлена на `__author__ = "GENPRO LAB"`.
- **MM008/MM012**: `from pyrevit import revit, DB, forms` → явные импорты
  `Autodesk.Revit.DB/UI` и `System.Windows.Forms`; `forms.SelectFromList` →
  WinForms `CheckedListBox` (`select_wall_types`), `forms.alert` →
  `TaskDialog` (`revit_ui_helpers.alert` / локальный `confirm`).
- **MM009**: wildcard-импортов в исходнике не было.
- **MM010**: `LookupParameter("GP_11_Группирование")` и
  `LookupParameter("Наружная стена")` убраны (см. «Изменение поведения»);
  чтение встроенных параметров — `revit_compat.get_parameter`
  с `BuiltInParameter` (`SYMBOL_NAME_PARAM`, `ROOM_NAME`) вместо прямых
  `get_Parameter(bip)` (D-04, pythonnet 3.x).
- **MM011**: восемь голых `except:` → `except Exception`.
- **MM005/MM006**: созданы `bundle.yaml` (title/tooltip ru+en) и `README.md`.
- **MM007**: кнопка зарегистрирована в `АРХИТЕКТУРА.panel/bundle.yaml`
  (после `IFC_Перекрытия`).
- **MM014**: добавлен канонический lib-бутстрап
  `_SCRIPT_DIR`/`_EXTENSION_DIR`/`_LIB_DIR`.
- **D-03**: код верхнего уровня свёрнут в `main(doc)`, которая начинается с
  `revit_compat.require_supported_version(COMMAND_NAME)`.
- **Правило 10**: транзакция — `Start()` перед `try`, `Commit()` в `try`,
  `RollBack()`+`raise` в `except` (в исходнике `except` глотал ошибку и
  показывал `forms.alert`).
- **Правило 11/18**: верхний уровень обёрнут в `try/except Exception` с
  `TaskDialog`; глобальный `doc = revit.doc` убран — `doc` берётся в `_entry()`
  и передаётся аргументом во все функции классификации.
- **2020/2022/2024-совместимость**: `wall_type.Width * 304.8` →
  `revit_compat.convert_from_internal(..., "mm")`;
  `material_id.IntegerValue != -1` и все сравнения `Id`/`GetTypeId()` →
  `revit_compat.element_id_value` (Int64 в 2024).
- **pythonnet-устойчивость**: `set` из объектов `Element` заменены на словари
  `{id стены: стена}` — идентичность .NET-прокси ненадёжна.
- Иконка исходника оставлена как `icon.png`.

## Изменение поведения (решения пользователя)
1. Исходник писал код МССК в «Комментарии» (`ALL_MODEL_INSTANCE_COMMENTS`)
   и имя категории в `GP_11_Группирование`, при этом `__doc__` обещал GP_01.
   По решению пользователя приведено к обещанному и к соседним кнопкам:
   код → `GP_01_КодКлассифМССК`, имя → `GP_01_ИмяКлассифМССК` (GUID те же,
   что в `IFC_Двери`/`IFC_Окна`/`IFC_Перекрытия`). «Комментарии» и `GP_11`
   больше не трогаются; docstring, tooltip и README приведены к фактическому
   поведению.
2. Признак наружной стены: пользовательский yes/no-параметр
   `LookupParameter("Наружная стена")` заменён штатным признаком Revit —
   `WallType.Function == WallFunction.Exterior`. Остальные проверки
   (префикс `ВС_`/`нс_`, «наруж» в имени типа, помещение только с одной
   стороны) сохранены.

## Исправленные дефекты исходника
- **Этап 2 (цепное «заражение»)**: флаг `added_new` сбрасывался раз за итерацию,
  а не на каждой стене, поэтому после первой же добавленной зашивки все
  последующие стены в той же итерации проверялись только на зашивку, но не на
  перегородку и отделку. Теперь правила перебираются на каждой стене с выходом
  по первому совпадению — задуманное поведение; результат на реальной модели
  может немного отличаться от исходника.
- **`is_partition_wall`, ПРОВЕРКА 2**: после отсечения наружных шёл запрос
  помещений с двух сторон, а затем `if not is_exterior_wall(wall): return True`,
  то есть ветка всегда возвращала `True`. Оставлен `return True` без лишних
  запросов помещений — поведение то же, вызовов API меньше.
- Имя типоразмера читается единообразно (`SYMBOL_NAME_PARAM` с фолбэком
  на `.Name`): в исходнике Этап 1 брал `SYMBOL_NAME_PARAM`, а функции-предикаты
  `.Name`.
- Стены без общих параметров `GP_01_*` больше не игнорируются молча —
  выводятся в окно pyRevit как `❌ ElementId — причина` и считаются
  пропущенными.

## Verify
- `py -3 tools/check_convention.py "…/IFC_Стены.pushbutton/script.py" --strict`
  → exit 0 (0 ошибок, 0 предупреждений).
- `py -3 tools/check_convention.py "…/IFC_Стены.pushbutton" --strict`
  → exit 0 после регистрации в layout.
- `py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json`
  → exit 1, но **ни одного нарушения по IFC_Стены**: остаток — предсуществующий
  долг `СНиП`, `СНиП_ФОП25` (нет записей в baseline) и untracked-кнопки
  `ПубликацияШаблона` (ещё не принята). Приёмкой не затронут.
- Сторонний скрипт ни разу не исполнялся: только чтение и статический чекер.
- Runtime-истина (UAT) — прогон кнопки в живом Revit после pyRevit Reload.
