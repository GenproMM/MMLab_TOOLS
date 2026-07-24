---
status: testing
phase: 03-convention
source: [03-VERIFICATION.md]
started: 2026-07-24
updated: 2026-07-24
---

## Current Test

number: 1
name: Compat-хелперы в реальном Revit 2020 / 2022 / 2024
expected: |
  Кнопка-пример из шаблона запускается на Revit 2020, 2022, 2024 без ошибок API;
  на неподдерживаемой версии — fail-fast TaskDialog с перечислением поддержанных версий (D-03).
awaiting: user response

## Tests

### 1. Compat-хелперы в реальном Revit 2020 / 2022 / 2024
expected: Кнопка-пример работает на Revit 2020/2022/2024; на иной версии — fail-fast TaskDialog (закрывает и WR-06: версия берётся из кеша require_supported_version).
result: [pending]

### 2. /mm-adopt-script на IFC_Двери — ревью-гейт (CONV-ADAPT)
expected: В Claude Code `/mm-adopt-script` на `MM LAB.extension/MM Lab.tab/АРХИТЕКТУРА.panel/IFC_Двери.pushbutton`: чекер отрабатывает → панель спрашивается с авто-подсказкой (D-10) → показан diff → без явного «да» регистрации НЕТ (D-08) → после «да» кнопка появляется в bundle.yaml.
result: [pending]

### 3. Quick task GSD после приёмки (CONV-GSD)
expected: После одобрения приёмки создана папка `.planning/quick/<id>-adopt-<слаг>/` с `<id>-PLAN.md` и `<id>-SUMMARY.md` + строка в таблице «Quick Tasks Completed» в `.planning/STATE.md`.
result: [pending]

### 4. Кнопка из шаблона видна на панели после pyRevit Reload
expected: Скопированный в `<Панель>.panel` и зарегистрированный в layout шаблон появляется на панели MM LAB после pyRevit Reload и запускается (диалоги, транзакция, отчёт «Стен в проекте: N»).
result: [pending]

### 5. /mm-check --all и /mm-doctor — вызываемость slash-команд
expected: `/mm-check --all` прогоняет чекер по репо; `/mm-doctor` выполняет self-check (версия Revit vs поддерживаемые, целостность vendored lib, обязательные файлы кнопок) в read-only режиме.
result: [pending]

## Summary

total: 5
passed: 0
issues: 0
pending: 5
skipped: 0
blocked: 0

## Gaps
