---
tags: [knowledge, integrations, revit-api]
date: 2026-06-09
---

# Интеграция с Revit API выполняется в процессе Revit

Скрипты работают in-process внутри Revit и используют Autodesk.Revit.DB / Autodesk.Revit.UI через CLR.

Что это дает:
- Доступ к элементам модели через FilteredElementCollector
- Запись изменений через Transaction
- Прямую работу с параметрами и категориями

Связанные утверждения:
- [[Все изменения модели должны проходить через Transaction с rollback]]
- [[Проект не использует отдельную БД и хранит данные в модели Revit]]
- [[Коллекторы Revit нужно сужать по категориям для производительности]]
