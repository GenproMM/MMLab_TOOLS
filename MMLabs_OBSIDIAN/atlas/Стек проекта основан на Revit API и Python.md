---
tags: [atlas, stack, revit-api, python]
date: 2026-06-09
---

# Стек проекта основан на Revit API и Python

Ключевой стек:
- Autodesk Revit API через CLR-импорты
- pyRevit runtime для кнопок, форм и контекста документа
- Python-скрипты в командах extension
- openpyxl и et_xmlfile как вендорные зависимости для экспорта Excel

Связанные утверждения:
- [[Интеграция с Revit API выполняется в процессе Revit]]
- [[Интеграция с openpyxl нужна для экспорта ПСО в Excel]]
- [[Мы вендорим openpyxl внутри репозитория для предсказуемых поставок]]
