#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gsd_release_sync.py — синхронизация «Карты релизов» (CSV) с плановыми артефактами GSD.

Реализует поток, описанный в Регламенте, п. 8.2.1 и 11.3.

Команды:
  check                       валидация CSV-экспорта «Карта релизов»
  sync-docs                   генерация .planning/release-map.json, ROADMAP.md,
                              REQUIREMENTS.md, PROJECT.md (с сохранением прогресса)
  sync                        check + sync-docs (эквивалент «Синхронизируй gsd»)
  status <ID> "<статус>"      смена статуса задания в CSV + запись в журнал переходов

Зависимости: только стандартная библиотека Python 3.

Источник истины — CSV «Карта релизов». Статусы задания меняются командой
``status`` (она же пишет CSV), а ``sync-docs`` детерминированно выводит из CSV
плановые документы. Жизненный цикл задания: Не начато → В работе → Готово → Релиз.
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
CSV_PATH = RELEASE_MAP_DIR / "Сводный Реестр Плагинов - Карта релизов.csv"
OUTPUT_DIR = RELEASE_MAP_DIR / "output"
JOURNAL_PATH = RELEASE_MAP_DIR / "status-journal.csv"
PLANNING_DIR = ROOT / ".planning"
RELEASE_MAP_JSON = PLANNING_DIR / "release-map.json"

# --- Доменные константы ---------------------------------------------------

STATUSES = ["Не начато", "В работе", "Готово", "Релиз"]
PRIORITIES = ["MVP", "Высокий", "Средний", "Низкий"]
REQUIRED_COLUMNS = ["ID", "Продукт", "PID", "Дисциплина", "Задание", "Релиз", "Приоритет", "Статус"]
RELEASE_RE = re.compile(r"^v\d{6}$")  # версия релиза вида vYYMMDD


# --- Утилиты --------------------------------------------------------------

def _now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")


def _fail(message):
    sys.stderr.write("ОШИБКА: %s\n" % message)
    sys.exit(1)


def _status_index(status):
    return STATUSES.index(status) if status in STATUSES else -1


def read_tasks():
    """Читает CSV «Карта релизов». Возвращает (rows, fieldnames, delimiter)."""
    if not CSV_PATH.exists():
        _fail("CSV «Карта релизов» не найден: %s" % CSV_PATH)
    with CSV_PATH.open(encoding="utf-8-sig", newline="") as fh:
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


def write_tasks(rows, fieldnames, delimiter):
    with CSV_PATH.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, delimiter=delimiter)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


# --- Команда check --------------------------------------------------------

def validate(rows, fieldnames):
    """Возвращает (errors, warnings)."""
    errors, warnings = [], []

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in (fieldnames or [])]
    if missing_cols:
        errors.append("В CSV отсутствуют обязательные колонки: %s" % ", ".join(missing_cols))
        return errors, warnings  # дальнейшая построчная проверка бессмысленна

    if not rows:
        warnings.append("CSV не содержит ни одного задания.")

    seen_ids = {}
    for i, row in enumerate(rows, start=2):  # строка 1 — заголовок
        rid = row.get("ID", "")
        if not rid:
            errors.append("Строка %d: пустой ID." % i)
        elif rid in seen_ids:
            errors.append("Строка %d: дубль ID '%s' (уже в строке %d)." % (i, rid, seen_ids[rid]))
        else:
            seen_ids[rid] = i

        if not row.get("Продукт"):
            errors.append("Строка %d (ID %s): пустой Продукт." % (i, rid or "?"))

        status = row.get("Статус", "")
        if status not in STATUSES:
            errors.append("Строка %d (ID %s): недопустимый Статус '%s' (ожидается один из: %s)."
                          % (i, rid or "?", status, ", ".join(STATUSES)))

        release = row.get("Релиз", "")
        if not RELEASE_RE.match(release):
            errors.append("Строка %d (ID %s): Релиз '%s' не соответствует шаблону vYYMMDD."
                          % (i, rid or "?", release))

        priority = row.get("Приоритет", "")
        if not priority:
            errors.append("Строка %d (ID %s): пустой Приоритет." % (i, rid or "?"))
        elif priority not in PRIORITIES:
            warnings.append("Строка %d (ID %s): нестандартный Приоритет '%s' (рекомендуется: %s)."
                            % (i, rid or "?", priority, ", ".join(PRIORITIES)))

    return errors, warnings


def cmd_check(_args):
    rows, fieldnames, _ = read_tasks()
    errors, warnings = validate(rows, fieldnames)
    for w in warnings:
        print("ВНИМАНИЕ: %s" % w)
    if errors:
        for e in errors:
            print("ОШИБКА: %s" % e)
        print("\nИтог: НЕВАЛИДНО — заданий: %d, ошибок: %d, предупреждений: %d"
              % (len(rows), len(errors), len(warnings)))
        sys.exit(1)
    print("Итог: OK — заданий: %d, ошибок: 0, предупреждений: %d" % (len(rows), len(warnings)))
    return 0


# --- Команда sync-docs ----------------------------------------------------

def _group_by_release(rows):
    releases = {}
    for row in rows:
        releases.setdefault(row.get("Релиз", "—"), []).append(row)
    return dict(sorted(releases.items()))


def _status_counts(rows):
    counts = {s: 0 for s in STATUSES}
    for row in rows:
        if row.get("Статус") in counts:
            counts[row["Статус"]] += 1
    return counts


def _preserve_progress(rows):
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
    for row in rows:
        old = prev_status.get(row.get("ID"))
        if old and _status_index(old) > _status_index(row.get("Статус", "")):
            preserved.append((row["ID"], row.get("Статус"), old))
            row["Статус"] = old
    return preserved


def _build_release_map(rows):
    releases = []
    for version, items in _group_by_release(rows).items():
        releases.append({
            "version": version,
            "summary": _status_counts(items),
            "tasks": [
                {k: row.get(k, "") for k in REQUIRED_COLUMNS + ["Описание"]}
                for row in items
            ],
        })
    return {
        "generated_at": _now(),
        "source": "RELEASE_MAP/%s" % CSV_PATH.name,
        "statuses": STATUSES,
        "summary": {"total": len(rows), "by_status": _status_counts(rows)},
        "releases": releases,
    }


def _checkbox(status):
    return "[x]" if status in ("Готово", "Релиз") else "[ ]"


def _render_roadmap(data):
    lines = [
        "# ROADMAP",
        "",
        "> Сгенерировано `RELEASE_MAP/gsd_release_sync.py sync-docs` — не редактировать вручную.",
        "> Источник: `%s`. Обновлено: %s." % (data["source"], data["generated_at"]),
        "",
    ]
    for rel in data["releases"]:
        s = rel["summary"]
        lines.append("## Фаза %s" % rel["version"])
        lines.append("")
        lines.append("Статус: Релиз %d / Готово %d / В работе %d / Не начато %d (всего %d)."
                     % (s["Релиз"], s["Готово"], s["В работе"], s["Не начато"], len(rel["tasks"])))
        lines.append("")
        for t in rel["tasks"]:
            lines.append("- %s **%s** (`%s`, %s, %s) — %s — _%s_"
                         % (_checkbox(t["Статус"]), t["Задание"], t["ID"], t["Продукт"],
                            t["Дисциплина"], t.get("Описание", ""), t["Статус"]))
        lines.append("")
    return "\n".join(lines)


def _render_requirements(data):
    lines = [
        "# REQUIREMENTS",
        "",
        "> Сгенерировано `RELEASE_MAP/gsd_release_sync.py sync-docs` — не редактировать вручную.",
        "> Источник: `%s`. Обновлено: %s." % (data["source"], data["generated_at"]),
        "",
    ]
    by_product = {}
    for rel in data["releases"]:
        for t in rel["tasks"]:
            by_product.setdefault(t["Продукт"], []).append((rel["version"], t))
    for product in sorted(by_product):
        lines.append("## %s" % product)
        lines.append("")
        for version, t in by_product[product]:
            lines.append("- `%s` [%s] %s — %s _(релиз %s, приоритет %s)_"
                         % (t["ID"], t["Статус"], t["Задание"], t.get("Описание", ""),
                            version, t["Приоритет"]))
        lines.append("")
    return "\n".join(lines)


def _render_project(data):
    s = data["summary"]["by_status"]
    products = sorted({t["Продукт"] for rel in data["releases"] for t in rel["tasks"]})
    disciplines = sorted({t["Дисциплина"] for rel in data["releases"] for t in rel["tasks"]})
    lines = [
        "# PROJECT",
        "",
        "> Сгенерировано `RELEASE_MAP/gsd_release_sync.py sync-docs` — не редактировать вручную.",
        "> Источник: `%s`. Обновлено: %s." % (data["source"], data["generated_at"]),
        "",
        "## Сводка",
        "",
        "| Метрика | Значение |",
        "| ----- | ----- |",
        "| Заданий всего | %d |" % data["summary"]["total"],
        "| Релизов (фаз) | %d |" % len(data["releases"]),
        "| Продуктов | %d |" % len(products),
        "| Дисциплин | %d |" % len(disciplines),
        "| Релиз | %d |" % s["Релиз"],
        "| Готово | %d |" % s["Готово"],
        "| В работе | %d |" % s["В работе"],
        "| Не начато | %d |" % s["Не начато"],
        "",
        "## Продукты",
        "",
    ]
    lines += ["- %s" % p for p in products]
    lines += ["", "## Релизы", ""]
    lines += ["- %s — заданий %d" % (rel["version"], len(rel["tasks"])) for rel in data["releases"]]
    lines.append("")
    return "\n".join(lines)


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text if text.endswith("\n") else text + "\n", encoding="utf-8")
    print("обновлено: %s" % path.relative_to(ROOT))


def cmd_sync_docs(_args):
    rows, fieldnames, delimiter = read_tasks()
    errors, _ = validate(rows, fieldnames)
    if errors:
        _fail("CSV невалиден — сначала исправьте ошибки (`check`). Найдено ошибок: %d" % len(errors))

    preserved = _preserve_progress(rows)
    for rid, csv_status, kept in preserved:
        print("ПРОГРЕСС СОХРАНЁН: %s — в CSV '%s', оставлен более продвинутый '%s'." % (rid, csv_status, kept))
    if preserved:
        write_tasks(rows, fieldnames, delimiter)  # фиксируем сохранённый прогресс обратно в CSV

    data = _build_release_map(rows)
    PLANNING_DIR.mkdir(parents=True, exist_ok=True)
    _write(RELEASE_MAP_JSON, json.dumps(data, ensure_ascii=False, indent=2))
    _write(PLANNING_DIR / "ROADMAP.md", _render_roadmap(data))
    _write(PLANNING_DIR / "REQUIREMENTS.md", _render_requirements(data))
    _write(PLANNING_DIR / "PROJECT.md", _render_project(data))

    # Копия release-map.json в output/ для выгрузки обратно на сетевой диск (п. 8.2.1, шаг 4)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    _write(OUTPUT_DIR / "release-map.json", json.dumps(data, ensure_ascii=False, indent=2))
    print("sync-docs выполнен: заданий %d, релизов %d." % (len(rows), len(data["releases"])))
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
        _fail("Недопустимый статус '%s'. Ожидается один из: %s" % (new_status, ", ".join(STATUSES)))

    rows, fieldnames, delimiter = read_tasks()
    target = next((r for r in rows if r.get("ID") == rid), None)
    if target is None:
        _fail("Задание с ID '%s' не найдено в Карте релизов." % rid)

    old_status = target.get("Статус", "")
    if old_status == new_status:
        print("Статус задания %s уже '%s' — изменений нет." % (rid, new_status))
        return 0
    if _status_index(new_status) < _status_index(old_status):
        print("ВНИМАНИЕ: откат статуса задания %s: '%s' → '%s'." % (rid, old_status, new_status))

    target["Статус"] = new_status
    write_tasks(rows, fieldnames, delimiter)
    _append_journal(rid, old_status, new_status)
    print("Статус задания %s: '%s' → '%s'. Журнал: %s"
          % (rid, old_status, new_status, JOURNAL_PATH.relative_to(ROOT)))
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
        sys.exit(main())
    except BrokenPipeError:
        # downstream получатель закрыл канал (например, `| head`) — выходим тихо
        import os
        os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
        sys.exit(0)
