---
tags: [knowledge, debugging, reliability, ios, revit-api]
date: 2026-07-09
---

# Несуществующие имена BuiltInParameter ломают кнопки ИОС при вызове

Проблема:
- Обращение к `BuiltInParameter.<ИМЯ>`, которого нет в enum Revit API, падает с `AttributeError: type object 'BuiltInParameter' has no attribute '...'` в момент выполнения.
- В `get_loss_method_parameters` использовались `RBS_DUCT_LOSS_METHOD_SERVER_PARAM` и `RBS_DUCT_TERMINAL_LOSS_METHOD_SERVER_PARAM` — обоих в API нет. Валиден только `RBS_DUCT_FITTING_LOSS_METHOD_SERVER_PARAM`, и он покрывает и fitting, и accessory.

Симптомы:
- Кнопка «Сброс потерь» показывает диалог ошибки сразу при клике.
- Отдельно: «Доп. расход = 1» падал с `No method matches given arguments for get_Parameter: (<class 'int'>)` — перегрузки `get_Parameter(BuiltInParameter)` неоднозначны на Revit 2024+/pythonnet.

Базовое решение:
- Проверять существование имени BuiltInParameter по официальной документации Autodesk перед использованием.
- Уходить от `get_Parameter(BuiltInParameter)` к `LookupParameter(string)` (одна перегрузка) — имя резолвится через `ParameterElement.Definition.Name`.
- Следить за контрактом возврата функций: `ensure_loss_method_undefined` возвращала bool, а вызывающий скрипт сравнивал со строками `"updated"`/`"already"` — всё тихо репортилось как «Пропущено».

Связанные утверждения:
- [[Мультиверсионная совместимость Revit API через адаптер версии]]
- [[Тихие except блоки скрывают критические ошибки в ИОС-скриптах]]
- [[Отсутствие тестов повышает риск регрессий при правках]]
- [[Кнопка Сброс потерь централизованно сбрасывает расчетные потери в ИОС]]
