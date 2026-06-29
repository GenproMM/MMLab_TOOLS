---
tags: [knowledge, patterns, architecture, pyrevit, revit-api]
date: 2026-06-09
---

# Мультиверсионная совместимость Revit API через адаптер версии

Паттерн фиксирует единый способ поддержки Revit 2020, 2022 и 2024 без дублирования version-check по всем кнопкам.

Практическая структура:
- Все межверсионные различия Revit API инкапсулируются в shared helper/adapter-модуле
- `script.py` вызывает только универсальные функции адаптера и не содержит локальных if по версиям
- Версия активной сессии Revit определяется один раз и передается в ветвление внутри адаптера

Минимальный шаблон адаптера:

```python
#! python3
# -*- coding: utf-8 -*-

def get_revit_major_version(document):
    app = document.Application
    return int(app.VersionNumber)


def is_revit_2020(version):
    return version == 2020


def is_revit_2022(version):
    return version == 2022


def is_revit_2024(version):
    return version == 2024


def get_parameter_id_compat(document):
    version = get_revit_major_version(document)

    if is_revit_2020(version):
        return get_parameter_id_for_2020(document)
    if is_revit_2022(version):
        return get_parameter_id_for_2022(document)
    if is_revit_2024(version):
        return get_parameter_id_for_2024(document)

    raise NotImplementedError("Unsupported Revit version: {0}".format(version))
```

Чеклист ревью мультиверсий:
- Есть явная поддержка Revit 2020, 2022, 2024
- Межверсионные ветки находятся в одном adapter/helper, а не размазаны по командам
- В `script.py` нет прямых вызовов API, известных как version-sensitive, без обертки
- Для неизвестной версии есть контролируемый fail-fast с понятным сообщением
- Изменения в API покрыты smoke-проверкой минимум на 2020/2022/2024

Связанные утверждения:
- [[Архитектура построена вокруг pyRevit pushbutton-скриптов]]
- [[Кнопки pyRevit должны быть тонкими и выносить общую логику в модули]]
