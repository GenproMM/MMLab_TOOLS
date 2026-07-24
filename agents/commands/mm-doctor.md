# /mm-doctor — self-check окружения и репозитория MM LAB

Диагностика БЕЗ правок: команда только собирает факты и советует. Любые
исправления — отдельным явным запросом пользователя (по итогам отчёта).

## Аргументы

Нет. Полный прогон всех шести групп проверок.

## Процедура

Выполни шесть групп проверок; падение отдельной проверки — не причина
останавливать остальные (фиксируй статус и иди дальше).

1. **Версии Revit на машине vs поддерживаемые.**
   Найди установленные версии Revit:

   ```powershell
   Get-ChildItem "C:\Program Files\Autodesk" -Directory -Filter "Revit 20*"
   ```

   и/или (в try — ветки реестра может не быть):

   ```powershell
   reg query "HKLM\SOFTWARE\Autodesk\Revit"
   ```

   Сравни найденные версии с кортежем `SUPPORTED_VERSIONS` из
   `MM LAB.extension/lib/revit_compat.py` (сейчас `(2020, 2022, 2024)` —
   читай из файла, не по памяти). Предупреди о неподдерживаемых установках:
   на них кнопки завершатся fail-fast диалогом (D-03).

2. **Целостность vendored-библиотек.** Папки существуют и не пусты:
   - `lib/openpyxl/`
   - `lib/et_xmlfile/`

3. **Конвенция и тесты инструментов:**

   ```bash
   py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json
   py -3 -m unittest discover -s tools/tests -q
   ```

   Ожидание: exit 0 у обоих. Нарушения пересказать по кнопкам и MM-кодам
   (как в `/mm-check`).

4. **Обязательные файлы репозитория:**
   - `AGENTS.md` (канонический стандарт);
   - `templates/НоваяКнопка.pushbutton/` (script.py + bundle.yaml + README.md);
   - `agents/commands/` — все 7 процедур: mm-adopt-script.md, mm-new-button.md,
     mm-check.md, mm-new-compat.md, mm-save-session.md, mm-update-repo.md,
     mm-doctor.md;
   - `tools/check_convention.py` и `tools/convention_baseline.json`.

5. **Git-состояние:**
   - текущая ветка: `git branch --show-current`;
   - чистота дерева: `git -c core.quotepath=false status --porcelain`
     (не пусто — подсказать `/mm-save-session`);
   - доступность origin (в try — офлайн допустим и НЕ ошибка):
     `git fetch --dry-run` — если недоступен, статус «офлайн», не «сломано».

6. **Итоговый отчёт** — по-русски, таблицей:

   | Проверка | Статус | Что делать |
   |----------|--------|------------|
   | Версии Revit | OK / внимание | например: «Revit 2023 не поддерживается — кнопки покажут fail-fast; расширить матрицу через /mm-new-compat» |
   | Vendored lib (openpyxl, et_xmlfile) | OK / отсутствует | восстановить из origin (`/mm-update-repo`) |
   | Конвенция (чекер --all + baseline) | OK / нарушения | разобрать по MM-кодам, адаптация — /mm-adopt-script |
   | Тесты tools/tests | OK / падают | показать вывод unittest |
   | Обязательные файлы | OK / не хватает | перечислить недостающие |
   | Git | ветка, чисто/грязно, origin | /mm-save-session, /mm-update-repo |

## Гейты и запреты

- **Диагностика read-only:** команда НИЧЕГО не правит, не удаляет
  и не переключает — только читает файлы, запускает статический чекер/тесты
  и git-команды чтения (`status`, `branch`, `fetch --dry-run`).
- Правки по итогам отчёта — только по явному запросу пользователя
  и соответствующей командой (`/mm-adopt-script`, `/mm-new-compat`,
  `/mm-update-repo`, `/mm-save-session`).
- Отсутствие сети (origin недоступен) — допустимое состояние, а не ошибка.

## Финал

Покажи пользователю итоговую таблицу «Проверка / Статус / Что делать»
(шаг 6) и краткий вывод: «окружение готово» либо список проблем
в порядке важности с рекомендованной командой для каждой.
