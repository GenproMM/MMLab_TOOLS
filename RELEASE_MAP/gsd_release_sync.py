#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gsd_release_sync.py — синхронизация «Карты релизов» (CSV) с плановыми артефактами GSD.

Реализует поток, описанный в Регламенте, п. 8.2.1 и 11.3.

Команды:
  check                       валидация CSV-экспорта листа «Скрипты_Карта релизов»
  sync-docs                   генерация .planning/release-map.json, ROADMAP.md,
                              REQUIREMENTS.md, PROJECT.md (с сохранением прогресса)
  sync                        check + sync-docs (эквивалент «Синхронизируй gsd»)
  status <ID> "<статус>"      смена статуса задания в CSV + запись в журнал переходов

Зависимости: только стандартная библиотека Python 3.

Источник истины — CSV «Карта релизов» (экспорт Google-таблицы «Сводный Реестр
Плагинов», лист «Скрипты_Карта релизов»). Статусы задания меняются командой
``status`` (она же пишет CSV), а ``sync-docs`` детерминированно выводит из CSV
плановые документы. Жизненный цикл задания: Не начато → В работе → Готово → Релиз.

Структура листа: строка-заголовок релиза (заполнена только «Версия релиза»)
задаёт релиз для всех идущих ниже строк-заданий (у них «Версия релиза» пустая).

Правила синхронизации (решения пользователя, сессия 2026-08-13):
  * задание со статусом «Без статуса» (или пустым) НЕ синхронизируется;
  * задание с пустым «Название плагина» НЕ синхронизируется;
  * «MVP = TRUE» → приоритет «MVP», иначе «Обычный».
Пропущенные задания остаются в CSV и перечисляются в отчёте — это не потеря
данных, а явный сигнал, что строку в Google-таблице нужно дозаполнить.
"""

import argparse
import csv
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

# --- Пути -----------------------------------------------------------------

RELEASE_MAP_DIR = Path(__file__).resolve().parent
ROOT = RELEASE_MAP_DIR.parent
CSV_NAME = "Сводный Реестр Плагинов - Скрипты_Карта релизов.csv"
CSV_PATH = RELEASE_MAP_DIR / CSV_NAME
OUTPUT_DIR = RELEASE_MAP_DIR / "output"
JOURNAL_PATH = RELEASE_MAP_DIR / "status-journal.csv"
PLANNING_DIR = ROOT / ".planning"
LEGACY_DIR = PLANNING_DIR / "_legacy"
RELEASE_MAP_JSON = PLANNING_DIR / "release-map.json"

# --- Колонки листа «Скрипты_Карта релизов» --------------------------------

COL_RELEASE = "Версия релиза"
COL_PLAN = "План релиза"
COL_PLUGIN = "Название плагина"
COL_ID = "ID"
COL_MVP = "MVP"
COL_WEIGHT = "Вес"
COL_STATUS = "Статус"
COL_GROUP = "Группа задач"
COL_DESC = "Описание"
COL_AUTHOR = "Автор"
COL_COMMENT = "Комментарий"

REQUIRED_COLUMNS = [
    COL_RELEASE, COL_PLAN, COL_PLUGIN, COL_ID, COL_MVP, COL_WEIGHT,
    COL_STATUS, COL_GROUP, COL_DESC, COL_AUTHOR, COL_COMMENT,
]

# --- Доменные константы ---------------------------------------------------

STATUSES = ["Не начато", "В работе", "Готово", "Релиз"]
STATUS_NONE = "Без статуса"          # рабочий статус Google-таблицы: в GSD не переносится
PRIORITY_MVP = "MVP"
PRIORITY_REGULAR = "Обычный"
RELEASE_RE = re.compile(r"^v\d{6}$")  # версия релиза вида vYYMMDD

SKIP_NO_STATUS = "без статуса"
SKIP_NO_PLUGIN = "без названия плагина"

# Маркеры генерируемого блока в .planning/*.md: всё вне маркеров — ручной текст,
# он переживает регенерацию (в ROADMAP.md так живут фазы GSD).
MARK_BEGIN = "<!-- RELEASE-MAP:BEGIN — сгенерировано gsd_release_sync.py, не редактировать -->"
MARK_END = "<!-- RELEASE-MAP:END -->"

# Заголовки секций, которые генератор считает своими при миграции файла без маркеров.
GENERATED_SECTIONS = {
    "ROADMAP.md": re.compile(r"^## (Фаза |Релиз )"),
    "REQUIREMENTS.md": re.compile(r"^## "),
    "PROJECT.md": re.compile(r"^## (Сводка|Продукты|Плагины|Группы задач|Релизы)\s*$"),
}


# --- Утилиты --------------------------------------------------------------

def _now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")


def _fail(message):
    sys.stderr.write("ОШИБКА: %s\n" % message)
    sys.exit(1)


def _status_index(status):
    return STATUSES.index(status) if status in STATUSES else -1


def _is_mvp(value):
    return (value or "").strip().upper() in ("TRUE", "ДА", "1", "YES")


def read_rows(csv_path=None):
    """Читает CSV «Карта релизов» как есть. Возвращает (rows, fieldnames, delimiter).

    rows — сырые строки листа (включая строки-заголовки релизов): порядок и состав
    колонок сохраняются, чтобы команда ``status`` могла записать CSV без потерь.
    """
    path = Path(csv_path) if csv_path else CSV_PATH
    if not path.exists():
        _fail("CSV «Карта релизов» не найден: %s\n"
              "Выгрузите лист «Скрипты_Карта релизов» из Google-таблицы в этот файл." % path)
    with path.open(encoding="utf-8-sig", newline="") as fh:
        sample = fh.read(8192)
        fh.seek(0)
        delimiter = ";" if sample.count(";") > sample.count(",") else ","
        reader = csv.DictReader(fh, delimiter=delimiter)
        fieldnames = reader.fieldnames or []
        rows = []
        for raw in reader:
            row = {(k or "").strip(): (v or "").strip() for k, v in raw.items() if k is not None}
            if any(row.values()):
                rows.append(row)
    return rows, fieldnames, delimiter


def write_rows(rows, fieldnames, delimiter, csv_path=None):
    path = Path(csv_path) if csv_path else CSV_PATH
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, delimiter=delimiter, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fieldnames})


# --- Разбор листа: строки-заголовки релизов + строки-задания --------------

def _is_release_header(row):
    """Строка-заголовок релиза: заполнена «Версия релиза» и нет ID."""
    return bool(row.get(COL_RELEASE)) and not row.get(COL_ID)


def _skip_reason(row):
    """Причина, по которой задание не синхронизируется, либо None."""
    status = row.get(COL_STATUS, "")
    if not status or status == STATUS_NONE:
        return SKIP_NO_STATUS
    if not row.get(COL_PLUGIN):
        return SKIP_NO_PLUGIN
    return None


def _make_task(row, release, line_no):
    return {
        "ID": row.get(COL_ID, ""),
        "Плагин": row.get(COL_PLUGIN, ""),
        "Релиз": release,
        "Статус": row.get(COL_STATUS, ""),
        "Приоритет": PRIORITY_MVP if _is_mvp(row.get(COL_MVP)) else PRIORITY_REGULAR,
        "MVP": _is_mvp(row.get(COL_MVP)),
        "Вес": row.get(COL_WEIGHT, ""),
        "Группа задач": row.get(COL_GROUP, ""),
        "Задание": row.get(COL_DESC, ""),
        "Автор": row.get(COL_AUTHOR, ""),
        "Комментарий": row.get(COL_COMMENT, ""),
        "Строка": line_no,
        "_row": row,
    }


def parse_tasks(rows):
    """Разбирает сырые строки листа.

    Возвращает (tasks, skipped, orphans):
      tasks   — задания, попадающие в синхронизацию;
      skipped — задания, отсеянные правилами (с причиной);
      orphans — строки-задания до первого заголовка релиза (ошибка структуры).
    """
    tasks, skipped, orphans = [], [], []
    release = None
    for i, row in enumerate(rows, start=2):  # строка 1 — заголовок таблицы
        if _is_release_header(row):
            release = row.get(COL_RELEASE)
            continue
        if not row.get(COL_ID):
            continue  # служебная/пустая строка листа
        task = _make_task(row, release or "", i)
        if release is None:
            orphans.append(task)
            continue
        reason = _skip_reason(row)
        if reason:
            task["Причина пропуска"] = reason
            skipped.append(task)
        else:
            tasks.append(task)
    return tasks, skipped, orphans


def _public(task):
    """Задание без служебных полей — для release-map.json."""
    return {k: v for k, v in task.items() if not k.startswith("_")}


# --- Команда check --------------------------------------------------------

def validate(rows, fieldnames):
    """Возвращает (errors, warnings, notes)."""
    errors, warnings, notes = [], [], []

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in (fieldnames or [])]
    if missing_cols:
        errors.append("В CSV отсутствуют обязательные колонки: %s" % ", ".join(missing_cols))
        return errors, warnings, notes  # дальнейшая построчная проверка бессмысленна

    releases = [r for r in rows if _is_release_header(r)]
    if not releases:
        errors.append("В CSV нет ни одной строки-заголовка релиза (заполнена только «%s»)." % COL_RELEASE)
    for row in releases:
        version = row.get(COL_RELEASE, "")
        if not RELEASE_RE.match(version):
            errors.append("Версия релиза «%s» не соответствует шаблону vYYMMDD." % version)

    tasks, skipped, orphans = parse_tasks(rows)

    for task in orphans:
        errors.append("Строка %s (ID %s): задание идёт до первой строки-заголовка релиза."
                      % (task["Строка"], task["ID"] or "?"))

    seen_ids = {}
    for task in tasks + skipped + orphans:
        rid, line = task["ID"], task["Строка"]
        if rid in seen_ids:
            errors.append("Строка %s: дубль ID «%s» (уже в строке %s)." % (line, rid, seen_ids[rid]))
        else:
            seen_ids[rid] = line

        status = task["Статус"]
        if status and status not in STATUSES and status != STATUS_NONE:
            errors.append("Строка %s (ID %s): недопустимый Статус «%s» (ожидается один из: %s, %s)."
                          % (line, rid, status, ", ".join(STATUSES), STATUS_NONE))

        mvp_raw = (task["_row"].get(COL_MVP) or "").strip().upper()
        if mvp_raw and mvp_raw not in ("TRUE", "FALSE"):
            warnings.append("Строка %s (ID %s): нестандартное значение MVP «%s» (ожидается TRUE/FALSE)."
                            % (line, rid, task["_row"].get(COL_MVP)))

    for task in tasks:
        if not task["Задание"]:
            warnings.append("Строка %s (ID %s): пустое «%s» — в плановых документах задание будет без текста."
                            % (task["Строка"], task["ID"], COL_DESC))

    if not tasks:
        warnings.append("Ни одно задание не проходит правила синхронизации.")

    for task in skipped:
        notes.append("Строка %s (ID %s): не синхронизируется — %s."
                     % (task["Строка"], task["ID"], task["Причина пропуска"]))

    return errors, warnings, notes


def cmd_check(_args):
    rows, fieldnames, _ = read_rows()
    errors, warnings, notes = validate(rows, fieldnames)
    tasks, skipped, _orphans = parse_tasks(rows)

    for n in notes:
        print("ПРОПУСК: %s" % n)
    for w in warnings:
        print("ВНИМАНИЕ: %s" % w)
    if errors:
        for e in errors:
            print("ОШИБКА: %s" % e)
        print("\nИтог: НЕВАЛИДНО — синхронизируется %d, пропущено %d, ошибок %d, предупреждений %d"
              % (len(tasks), len(skipped), len(errors), len(warnings)))
        sys.exit(1)
    print("Итог: OK — синхронизируется %d, пропущено %d, ошибок 0, предупреждений %d"
          % (len(tasks), len(skipped), len(warnings)))
    return 0


# --- Команда sync-docs ----------------------------------------------------

def _group_by_release(tasks):
    releases = {}
    for task in tasks:
        releases.setdefault(task["Релиз"], []).append(task)
    return dict(sorted(releases.items()))


def _status_counts(tasks):
    counts = {s: 0 for s in STATUSES}
    for task in tasks:
        if task["Статус"] in counts:
            counts[task["Статус"]] += 1
    return counts


def _preserve_progress(tasks):
    """Сохраняет уже выполненный прогресс: если в прошлом release-map.json статус
    задания был более продвинутым, чем в текущем CSV, оставляет продвинутый."""
    if not RELEASE_MAP_JSON.exists():
        return []
    try:
        prev = json.loads(RELEASE_MAP_JSON.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return []
    prev_status = {}
    for rel in prev.get("releases", []):
        for task in rel.get("tasks", []):
            prev_status[task.get("ID")] = task.get("Статус")
    preserved = []
    for task in tasks:
        old = prev_status.get(task["ID"])
        if old and _status_index(old) > _status_index(task["Статус"]):
            preserved.append((task["ID"], task["Статус"], old))
            task["Статус"] = old
            task["_row"][COL_STATUS] = old
    return preserved


def _build_release_map(tasks, skipped):
    releases = []
    for version, items in _group_by_release(tasks).items():
        releases.append({
            "version": version,
            "summary": _status_counts(items),
            "tasks": [_public(t) for t in items],
        })
    return {
        "generated_at": _now(),
        "source": "RELEASE_MAP/%s" % CSV_PATH.name,
        "statuses": STATUSES,
        "rules": {
            "skip_statuses": ["", STATUS_NONE],
            "skip_without_plugin": True,
            "priority": {"MVP=TRUE": PRIORITY_MVP, "иначе": PRIORITY_REGULAR},
        },
        "summary": {
            "total": len(tasks),
            "skipped": len(skipped),
            "by_status": _status_counts(tasks),
        },
        "releases": releases,
        "skipped": [_public(t) for t in skipped],
    }


def _checkbox(status):
    return "[x]" if status in ("Готово", "Релиз") else "[ ]"


def _banner(data, title):
    return [
        "# %s" % title,
        "",
        "> Сгенерировано `RELEASE_MAP/gsd_release_sync.py sync-docs` — не редактировать вручную.",
        "> Источник: `%s`. Обновлено: %s." % (data["source"], data["generated_at"]),
        "",
    ]


def _render_roadmap(data):
    lines = _banner(data, "ROADMAP")
    for rel in data["releases"]:
        s = rel["summary"]
        lines.append("## Релиз %s" % rel["version"])
        lines.append("")
        lines.append("Статус: Релиз %d / Готово %d / В работе %d / Не начато %d (всего %d)."
                     % (s["Релиз"], s["Готово"], s["В работе"], s["Не начато"], len(rel["tasks"])))
        lines.append("")
        for t in rel["tasks"]:
            lines.append("- %s **%s** (`%s`, %s, %s) — %s — _%s_"
                         % (_checkbox(t["Статус"]), t["Плагин"], t["ID"], t["Группа задач"],
                            t["Приоритет"], t["Задание"] or "описание не заполнено", t["Статус"]))
        lines.append("")
    if data["skipped"]:
        lines.append("## Не синхронизировано")
        lines.append("")
        lines.append("Задания листа, отсеянные правилами (дозаполните строку в Google-таблице):")
        lines.append("")
        for t in data["skipped"]:
            lines.append("- `%s` (строка %s) — %s%s"
                         % (t["ID"], t["Строка"], t["Причина пропуска"],
                            (": %s" % t["Задание"]) if t["Задание"] else ""))
        lines.append("")
    return "\n".join(lines)


def _render_requirements(data):
    lines = _banner(data, "REQUIREMENTS")
    by_plugin = {}
    for rel in data["releases"]:
        for t in rel["tasks"]:
            by_plugin.setdefault(t["Плагин"], []).append((rel["version"], t))
    for plugin in sorted(by_plugin):
        lines.append("## %s" % plugin)
        lines.append("")
        for version, t in by_plugin[plugin]:
            lines.append("- `%s` [%s] %s — %s _(релиз %s, приоритет %s, группа %s)_"
                         % (t["ID"], t["Статус"], t["Задание"] or "—", t["Комментарий"] or "—",
                            version, t["Приоритет"], t["Группа задач"] or "—"))
        lines.append("")
    return "\n".join(lines)


def _render_project(data):
    s = data["summary"]["by_status"]
    plugins = sorted({t["Плагин"] for rel in data["releases"] for t in rel["tasks"]})
    groups = sorted({t["Группа задач"] for rel in data["releases"] for t in rel["tasks"] if t["Группа задач"]})
    lines = _banner(data, "PROJECT")
    lines += [
        "## Сводка",
        "",
        "| Метрика | Значение |",
        "| ----- | ----- |",
        "| Заданий синхронизировано | %d |" % data["summary"]["total"],
        "| Заданий пропущено | %d |" % data["summary"]["skipped"],
        "| Релизов | %d |" % len(data["releases"]),
        "| Плагинов | %d |" % len(plugins),
        "| Групп задач | %d |" % len(groups),
        "| Релиз | %d |" % s["Релиз"],
        "| Готово | %d |" % s["Готово"],
        "| В работе | %d |" % s["В работе"],
        "| Не начато | %d |" % s["Не начато"],
        "",
        "## Плагины",
        "",
    ]
    lines += ["- %s" % p for p in plugins]
    lines += ["", "## Релизы", ""]
    lines += ["- %s — заданий %d" % (rel["version"], len(rel["tasks"])) for rel in data["releases"]]
    lines.append("")
    return "\n".join(lines)


# --- Запись .planning с сохранением ручного текста -------------------------

def _manual_tail(text, name):
    """Ручной хвост файла без маркеров: всё, кроме заголовка, баннера и секций генератора.

    Нужен один раз — при миграции файла, сгенерированного старой версией скрипта.
    В ROADMAP.md так переживают регенерацию фазы GSD (`## Phase N: ...`).
    """
    pattern = GENERATED_SECTIONS.get(name)
    kept, drop_section = [], False
    for line in text.splitlines():
        if line.startswith("## "):
            drop_section = bool(pattern and pattern.match(line))
        if drop_section:
            continue
        if line.startswith("# ") or line.startswith("> "):
            continue  # заголовок документа и баннер генератора
        kept.append(line)
    return "\n".join(kept).strip("\n")


def _write_generated(path, text):
    """Пишет генерируемый блок в маркерах, сохраняя ручной текст вокруг него."""
    path.parent.mkdir(parents=True, exist_ok=True)
    block = "%s\n%s\n%s" % (MARK_BEGIN, text.rstrip("\n"), MARK_END)

    if path.exists():
        old = path.read_text(encoding="utf-8")
        if MARK_BEGIN in old and MARK_END in old:
            head = old.split(MARK_BEGIN, 1)[0]
            tail = old.split(MARK_END, 1)[1]
            new = head + block + tail
        else:
            LEGACY_DIR.mkdir(parents=True, exist_ok=True)
            backup = LEGACY_DIR / ("%s.%s.bak" % (path.name, _now().replace(":", "-").replace(" ", "_")))
            backup.write_text(old, encoding="utf-8")
            print("резервная копия: %s" % backup.relative_to(ROOT))
            tail = _manual_tail(old, path.name)
            new = block + ("\n\n" + tail if tail else "")
    else:
        new = block

    path.write_text(new if new.endswith("\n") else new + "\n", encoding="utf-8")
    print("обновлено: %s" % path.relative_to(ROOT))


def _write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("обновлено: %s" % path.relative_to(ROOT))


def cmd_sync_docs(_args):
    rows, fieldnames, delimiter = read_rows()
    errors, _warnings, _notes = validate(rows, fieldnames)
    if errors:
        _fail("CSV невалиден — сначала исправьте ошибки (`check`). Найдено ошибок: %d" % len(errors))

    tasks, skipped, _orphans = parse_tasks(rows)

    preserved = _preserve_progress(tasks)
    for rid, csv_status, kept in preserved:
        print("ПРОГРЕСС СОХРАНЁН: %s — в CSV «%s», оставлен более продвинутый «%s»." % (rid, csv_status, kept))
    if preserved:
        write_rows(rows, fieldnames, delimiter)  # фиксируем сохранённый прогресс обратно в CSV

    for task in skipped:
        print("ПРОПУСК: ID %s (строка %s) — %s." % (task["ID"], task["Строка"], task["Причина пропуска"]))

    data = _build_release_map(tasks, skipped)
    PLANNING_DIR.mkdir(parents=True, exist_ok=True)
    _write_json(RELEASE_MAP_JSON, data)
    _write_generated(PLANNING_DIR / "ROADMAP.md", _render_roadmap(data))
    _write_generated(PLANNING_DIR / "REQUIREMENTS.md", _render_requirements(data))
    _write_generated(PLANNING_DIR / "PROJECT.md", _render_project(data))

    # Копия release-map.json в output/ для выгрузки обратно на сетевой диск (п. 8.2.1, шаг 4)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    _write_json(OUTPUT_DIR / "release-map.json", data)
    print("sync-docs выполнен: заданий %d, релизов %d, пропущено %d."
          % (len(tasks), len(data["releases"]), len(skipped)))
    return 0


# --- Команда sync (check + sync-docs) ------------------------------------

def cmd_sync(args):
    print("== check ==")
    cmd_check(args)
    print("\n== sync-docs ==")
    cmd_sync_docs(args)
    return 0


# --- Команда status -------------------------------------------------------

def _append_journal(rid, old, new):
    new_file = not JOURNAL_PATH.exists()
    with JOURNAL_PATH.open("a", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        if new_file:
            writer.writerow(["Время", "ID", "Старый статус", "Новый статус"])
        writer.writerow([_now(), rid, old, new])


def cmd_status(args):
    rid, new_status = args.id, args.status
    if new_status not in STATUSES:
        _fail("Недопустимый статус «%s». Ожидается один из: %s" % (new_status, ", ".join(STATUSES)))

    rows, fieldnames, delimiter = read_rows()
    target = next((r for r in rows if r.get(COL_ID) == rid), None)
    if target is None:
        _fail("Задание с ID «%s» не найдено в Карте релизов." % rid)

    old_status = target.get(COL_STATUS, "")
    if old_status == new_status:
        print("Статус задания %s уже «%s» — изменений нет." % (rid, new_status))
        return 0
    if _status_index(new_status) < _status_index(old_status):
        print("ВНИМАНИЕ: откат статуса задания %s: «%s» → «%s»." % (rid, old_status, new_status))

    target[COL_STATUS] = new_status
    write_rows(rows, fieldnames, delimiter)
    _append_journal(rid, old_status, new_status)
    print("Статус задания %s: «%s» → «%s». Журнал: %s"
          % (rid, old_status or STATUS_NONE, new_status, JOURNAL_PATH.relative_to(ROOT)))
    if not target.get(COL_PLUGIN):
        print("ВНИМАНИЕ: у задания пустое «%s» — оно всё равно не попадёт в плановые документы."
              % COL_PLUGIN)
    print("Подсказка: запустите `sync-docs`, чтобы обновить плановые документы.")
    return 0


# --- CLI ------------------------------------------------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="gsd_release_sync.py",
        description="Синхронизация Карты релизов (CSV) с артефактами GSD (Регламент, п. 8.2.1 / 11.3).",
    )
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("check", help="валидация CSV «Карта релизов»")
    sub.add_parser("sync-docs", help="генерация плановых артефактов из CSV")
    sub.add_parser("sync", help="check + sync-docs («Синхронизируй gsd»)")
    p_status = sub.add_parser("status", help="смена статуса задания по ID")
    p_status.add_argument("id", help="ID задания из колонки ID")
    p_status.add_argument("status", help="новый статус: %s" % " / ".join(STATUSES))

    args = parser.parse_args(argv)
    handlers = {
        "check": cmd_check,
        "sync-docs": cmd_sync_docs,
        "sync": cmd_sync,
        "status": cmd_status,
    }
    return handlers[args.command](args)


if __name__ == "__main__":
    try:
        # консоль Windows может не кодировать часть символов — не роняем отчёт из-за этого
        sys.stdout.reconfigure(errors="replace")
        sys.stderr.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    try:
        sys.exit(main())
    except BrokenPipeError:
        # downstream получатель закрыл канал (например, `| head`) — выходим тихо
        import os
        os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
        sys.exit(0)
