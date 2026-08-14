#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gsd_release_sync.py — синхронизация «Карты релизов» (CSV) с плановыми артефактами GSD.

Реализует поток, описанный в Регламенте, п. 8.2.1 и 11.3.

Команды:
  fetch                       забрать свежий экспорт листа из «Загрузок»
                              (%USERPROFILE%\\Downloads) в RELEASE_MAP/
  check                       валидация CSV-экспорта листа «Скрипты_Карта релизов»
  sync-docs                   генерация .planning/release-map.json, ROADMAP.md,
                              REQUIREMENTS.md, PROJECT.md (с сохранением прогресса)
  sync                        fetch + check + sync-docs (эквивалент «Синхронизируй gsd»)
  reset                       очистка .planning от Карты релизов под новый набор
                              заданий (фазы GSD и STATE.md не трогаются)
  status <ID> "<статус>"      смена статуса задания в CSV + запись в журнал переходов

Зависимости: только стандартная библиотека Python 3.

Источник истины — CSV «Карта релизов» (экспорт Google-таблицы «Сводный Реестр
Плагинов», лист «Скрипты_Карта релизов»). Google-таблица выгружается браузером
в каталог «Загрузки», откуда команда ``fetch`` копирует экспорт в RELEASE_MAP/;
дальше работа идёт только с копией в репозитории (её же правит ``status``).
Статусы задания меняются командой ``status`` (она же пишет CSV), а ``sync-docs``
детерминированно выводит из CSV плановые документы.
Жизненный цикл задания: Не начато → В работе → Готово → Релиз.

Структура листа: строка-заголовок релиза (заполнена только «Версия релиза»)
задаёт релиз для всех идущих ниже строк-заданий (у них «Версия релиза» пустая).

Соответствие «Карта релизов» ↔ GSD:
  * версия релиза vYYMMDD  → фаза GSD «Phase vYYMMDD»;
  * ID задания             → требование GSD «REQ-<ID>»;
  * дата фазы              → ТОЛЬКО из версии релиза (v260814 → 2026-08-14);
    колонка «План релиза» — рудимент листа и не импортируется;
  * статус «Релиз» в CSV   → подтверждение закрытия требования в GSD; фаза
    закрыта, когда закрыты все её требования (статус «Готово» не закрывает).

Правила синхронизации (решения пользователя, сессии 2026-08-13 / 2026-08-14):
  * задание со статусом «Без статуса» (или пустым) НЕ синхронизируется;
  * задание с пустым «Название плагина» НЕ синхронизируется;
  * «MVP = TRUE» → приоритет «MVP», иначе «Обычный»;
  * НОВОЕ задание (которого не было в прошлом release-map.json) заводится в GSD
    ТОЛЬКО со статусом «Не начато»: пришедшее сразу «В работе»/«Готово»/«Релиз» —
    это работа мимо GSD, в план она не попадает;
  * уже заведённые задания при каждом синке обновляются из текущего CSV
    ЦЕЛИКОМ — статус, описание, комментарий, даты, автор, вес, группа, MVP;
    прошлый release-map.json значения не переопределяет.
Пропущенные задания остаются в CSV и перечисляются в отчёте — это не потеря
данных, а явный сигнал, что строку в Google-таблице нужно дозаполнить.
"""

import argparse
import csv
import json
import os
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

# --- Пути -----------------------------------------------------------------

RELEASE_MAP_DIR = Path(__file__).resolve().parent
ROOT = RELEASE_MAP_DIR.parent
CSV_NAME = "Сводный Реестр Плагинов - Скрипты_Карта релизов.csv"
CSV_PATH = RELEASE_MAP_DIR / CSV_NAME
OUTPUT_DIR = RELEASE_MAP_DIR / "output"
# Каталог «Загрузки», куда браузер кладёт экспорт листа. Переопределяется
# переменной окружения (нужна тестам и нестандартным профилям Windows).
DOWNLOADS_ENV = "MMLAB_DOWNLOADS_DIR"
JOURNAL_PATH = RELEASE_MAP_DIR / "status-journal.csv"
PLANNING_DIR = ROOT / ".planning"
LEGACY_DIR = PLANNING_DIR / "_legacy"
RELEASE_MAP_JSON = PLANNING_DIR / "release-map.json"

# --- Колонки листа «Скрипты_Карта релизов» --------------------------------

COL_RELEASE = "Версия релиза"
# «План релиза» — рудимент листа: НЕ импортируется. Дата релиза берётся
# исключительно из версии vYYMMDD (release_date), поэтому колонка не обязательна.
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
    COL_RELEASE, COL_PLUGIN, COL_ID, COL_MVP, COL_WEIGHT,
    COL_STATUS, COL_GROUP, COL_DESC, COL_AUTHOR, COL_COMMENT,
]

# --- Доменные константы ---------------------------------------------------

STATUSES = ["Не начато", "В работе", "Готово", "Релиз"]
STATUS_NEW = STATUSES[0]             # единственный статус, с которым задание заводится в GSD
STATUS_NONE = "Без статуса"          # рабочий статус Google-таблицы: в GSD не переносится
PRIORITY_MVP = "MVP"
PRIORITY_REGULAR = "Обычный"
RELEASE_RE = re.compile(r"^v\d{6}$")  # версия релиза вида vYYMMDD

SKIP_NO_STATUS = "без статуса"
SKIP_NO_PLUGIN = "без названия плагина"
SKIP_NEW_NOT_STARTED = "новое задание не в статусе «%s»" % STATUS_NEW

# Поля задания, изменения которых показываются в отчёте sync-docs. Значения
# всегда берутся из текущего CSV — отчёт лишь показывает, что именно приехало.
TRACKED_FIELDS = [
    "Статус", "Плагин", "Релиз", "Приоритет", "Вес",
    "Группа задач", "Задание", "Автор", "Комментарий",
]

# --- Соответствие «Карта релизов» ↔ GSD -----------------------------------
# Версия релиза = фаза GSD (v260814 → «Phase v260814»);
# задание листа  = требование GSD (ID 2830001 → «REQ-2830001»);
# закрытие требования в GSD подтверждается статусом «Релиз» в CSV.

REQ_PREFIX = "REQ-"
PHASE_PREFIX = "Phase "
STATUS_CLOSED = STATUSES[-1]          # «Релиз» — единственный статус закрытия в GSD

PHASE_NOT_STARTED = "Не начата"
PHASE_IN_PROGRESS = "В работе"
PHASE_READY = "Готова к релизу"
PHASE_CLOSED = "Закрыта релизом"

# Маркеры генерируемого блока в .planning/*.md: всё вне маркеров — ручной текст,
# он переживает регенерацию (в ROADMAP.md так живут фазы GSD).
MARK_BEGIN = "<!-- RELEASE-MAP:BEGIN — сгенерировано gsd_release_sync.py, не редактировать -->"
MARK_END = "<!-- RELEASE-MAP:END -->"

# Заголовки секций, которые генератор считает своими при миграции файла без маркеров.
GENERATED_SECTIONS = {
    # «## Phase v260814» — фаза из Карты релизов (генерируется);
    # «## Phase 3: ...» — ручная фаза GSD, регенерацию переживает.
    "ROADMAP.md": re.compile(r"^## (Фаза |Релиз |Phase v\d{6})"),
    "REQUIREMENTS.md": re.compile(r"^## "),
    "PROJECT.md": re.compile(r"^## (Сводка|Продукты|Плагины|Группы задач|Релизы|Фазы \(релизы\))\s*$"),
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


def req_id(task_id):
    """Требование GSD из ID задания: 2830001 → REQ-2830001."""
    task_id = (task_id or "").strip()
    if not task_id:
        return ""
    return task_id if task_id.startswith(REQ_PREFIX) else REQ_PREFIX + task_id


def phase_name(version):
    """Фаза GSD из версии релиза: v260814 → «Phase v260814»."""
    return (PHASE_PREFIX + version) if version else ""


def release_date(version):
    """Дата релиза из версии vYYMMDD: v260814 → «2026-08-14».

    Единственный источник даты релиза (колонка «План релиза» — рудимент листа
    и не импортируется). Некорректная версия → пустая строка.
    """
    if not RELEASE_RE.match(version or ""):
        return ""
    try:
        return datetime.strptime(version[1:], "%y%m%d").strftime("%Y-%m-%d")
    except ValueError:
        return ""


def phase_status(counts, total):
    """Статус фазы GSD, выведенный из статусов её требований.

    Закрытие подтверждает только «Релиз»: пока хоть одно требование не в нём,
    фаза не закрыта.
    """
    if not total:
        return PHASE_NOT_STARTED
    if counts[STATUS_CLOSED] == total:
        return PHASE_CLOSED
    if counts[STATUS_CLOSED] + counts["Готово"] == total:
        return PHASE_READY
    if counts["Не начато"] == total:
        return PHASE_NOT_STARTED
    return PHASE_IN_PROGRESS


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


# --- Команда fetch: экспорт листа из «Загрузок» ---------------------------

def downloads_dir():
    """Каталог «Загрузки»: %USERPROFILE%\\Downloads (вне Windows — ~/Downloads).

    Переопределяется переменной окружения MMLAB_DOWNLOADS_DIR.
    """
    override = os.environ.get(DOWNLOADS_ENV)
    if override:
        return Path(override)
    home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return Path(home) / "Downloads"


def find_downloaded_csv(directory=None):
    """Самый свежий экспорт листа в «Загрузках» либо None.

    Браузер нумерует повторные скачивания («… (1).csv»), поэтому берётся не
    только точное имя файла, но и варианты с суффиксом — из них самый новый
    по времени изменения.
    """
    directory = Path(directory) if directory else downloads_dir()
    if not directory.is_dir():
        return None
    stem = CSV_NAME[:-len(".csv")]
    candidates = [p for p in directory.glob("%s*.csv" % stem) if p.is_file()]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def cmd_fetch(_args=None):
    """Копирует свежий экспорт листа из «Загрузок» в RELEASE_MAP/.

    Если экспорта в «Загрузках» нет — предупреждает и оставляет текущий CSV
    репозитория (синк пройдёт по нему); ошибка только если CSV нет вообще.
    """
    directory = downloads_dir()
    source = find_downloaded_csv(directory)

    if source is None:
        if CSV_PATH.exists():
            print("ВНИМАНИЕ: в «Загрузках» (%s) нет экспорта «%s» — "
                  "синк пройдёт по текущему %s." % (directory, CSV_NAME, CSV_PATH.name))
            return 0
        _fail("В «Загрузках» (%s) нет экспорта «%s», и в репозитории CSV тоже нет.\n"
              "Выгрузите лист «Скрипты_Карта релизов» из Google-таблицы в «Загрузки»."
              % (directory, CSV_NAME))

    if CSV_PATH.exists() and source.read_bytes() == CSV_PATH.read_bytes():
        print("экспорт из «Загрузок» совпадает с текущим CSV — копирование не требуется: %s" % source)
        return 0

    shutil.copyfile(source, CSV_PATH)
    print("загружено: %s → %s" % (source, CSV_PATH.relative_to(ROOT)))
    return 0


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
        "REQ": req_id(row.get(COL_ID, "")),      # требование GSD
        "Плагин": row.get(COL_PLUGIN, ""),
        "Релиз": release,                        # = фаза GSD
        "Фаза": phase_name(release),
        "Дата релиза": release_date(release),    # дата только из версии релиза
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

    Все поля задания читаются из текущего CSV — прошлые прогоны на разбор
    не влияют (лист «Карта релизов» — единственный источник правды).
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
        elif not release_date(version):
            # версия — единственный источник даты фазы, поэтому дата обязана быть реальной
            errors.append("Версия релиза «%s»: несуществующая дата (vYYMMDD)." % version)

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


def previous_tasks():
    """Задания прошлого синка (из .planning/release-map.json) по ID.

    Это база «что уже есть в GSD». None — базы ещё нет (первый прогон или
    файл нечитаем): гейт новых заданий тогда не применяется, иначе первый же
    синк выкинул бы весь уже выполненный состав.
    """
    if not RELEASE_MAP_JSON.exists():
        return None
    try:
        prev = json.loads(RELEASE_MAP_JSON.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return None
    return {task.get("ID"): task
            for rel in prev.get("releases", [])
            for task in rel.get("tasks", [])}


def gate_new_tasks(tasks, skipped, previous):
    """Новое задание попадает в GSD только со статусом «Не начато».

    Задание, которого не было в прошлом release-map.json и которое приехало из
    листа сразу «В работе»/«Готово»/«Релиз», — это работа, сделанная мимо GSD:
    в план она не заводится. Уже известные GSD задания обновляются из CSV
    в любом статусе.
    """
    if previous is None:
        return tasks, skipped
    kept = []
    for task in tasks:
        if task["ID"] in previous or task["Статус"] == STATUS_NEW:
            kept.append(task)
        else:
            task["Причина пропуска"] = SKIP_NEW_NOT_STARTED
            skipped.append(task)
    return kept, sorted(skipped, key=lambda t: t["Строка"])


def _changed_since_last_sync(tasks, skipped, previous):
    """Что изменилось в заданиях по сравнению с прошлым release-map.json.

    Чисто информационный отчёт: значения НЕ переносятся из прошлого прогона —
    все поля берутся из текущего CSV. Возвращает список строк отчёта.
    """
    if previous is None:
        return []

    report = []
    for task in tasks:
        old = previous.get(task["ID"])
        if old is None:
            report.append("новое требование %s в %s (%s)"
                          % (task["REQ"], task["Фаза"], task["Статус"]))
            continue
        diffs = ["%s «%s» → «%s»" % (field, old.get(field, ""), task[field])
                 for field in TRACKED_FIELDS
                 if (old.get(field, "") or "") != (task[field] or "")]
        if diffs:
            report.append("требование %s: %s" % (task["REQ"], "; ".join(diffs)))
    for task in skipped:
        if task["ID"] in previous:
            report.append("требование %s выпало из синхронизации — %s"
                          % (task["REQ"], task["Причина пропуска"]))
    current_ids = {t["ID"] for t in tasks + skipped}
    for rid in previous:
        if rid not in current_ids:
            report.append("требование %s удалено из листа — убрано из плановых документов" % req_id(rid))
    return report


def _closures(tasks, previous):
    """Закрытия, подтверждённые новым CSV: требования и фазы в статусе «Релиз».

    Статус «Релиз» в Карте релизов — единственное подтверждение закрытия
    требования в GSD; фаза закрывается, когда закрыты все её требования.
    """
    report = []
    for task in tasks:
        if task["Статус"] != STATUS_CLOSED:
            continue
        was = (previous or {}).get(task["ID"], {}).get("Статус")
        if was != STATUS_CLOSED:
            report.append("требование %s закрыто релизом %s (%s) — было «%s»"
                          % (task["REQ"], task["Релиз"], task["Дата релиза"], was or "нет в GSD"))
    for version, items in _group_by_release(tasks).items():
        counts = _status_counts(items)
        if phase_status(counts, len(items)) != PHASE_CLOSED:
            continue
        was_closed = previous is not None and all(
            (previous.get(t["ID"], {}).get("Статус")) == STATUS_CLOSED for t in items)
        if not was_closed:
            report.append("фаза %s закрыта: все %d требований в статусе «%s»"
                          % (phase_name(version), len(items), STATUS_CLOSED))
    return report


def _build_release_map(tasks, skipped):
    releases = []
    for version, items in _group_by_release(tasks).items():
        counts = _status_counts(items)
        releases.append({
            "version": version,
            "phase": phase_name(version),
            "date": release_date(version),
            "phase_status": phase_status(counts, len(items)),
            "closed": counts[STATUS_CLOSED],
            "summary": counts,
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
            "new_task_status": STATUS_NEW,
            "closing_status": STATUS_CLOSED,
            "mapping": {
                "Версия релиза": "фаза GSD (%s<версия>)" % PHASE_PREFIX,
                "ID задания": "требование GSD (%s<ID>)" % REQ_PREFIX,
                "Дата релиза": "из версии vYYMMDD",
            },
            # прошлый release-map.json на значения не влияет: все поля — из CSV
            "source_of_truth": "CSV «Карта релизов»",
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
    """Закрытым требование считается только со статусом «Релиз» (D-21)."""
    return "[x]" if status == STATUS_CLOSED else "[ ]"


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
        total = len(rel["tasks"])
        lines.append("## %s%s" % (rel["phase"], (" — %s" % rel["date"]) if rel["date"] else ""))
        lines.append("")
        lines.append("**Статус фазы:** %s — закрыто %d из %d требований (статус «%s»)."
                     % (rel["phase_status"], rel["closed"], total, STATUS_CLOSED))
        lines.append("")
        lines.append("**Требования:** %s" % ", ".join(t["REQ"] for t in rel["tasks"]))
        lines.append("")
        lines.append("Статусы: Релиз %d / Готово %d / В работе %d / Не начато %d (всего %d)."
                     % (s["Релиз"], s["Готово"], s["В работе"], s["Не начато"], total))
        lines.append("")
        for t in rel["tasks"]:
            lines.append("- %s **%s** — %s (%s, %s) — %s — _%s_"
                         % (_checkbox(t["Статус"]), t["REQ"], t["Плагин"], t["Группа задач"],
                            t["Приоритет"], t["Задание"] or "описание не заполнено", t["Статус"]))
            if t["Комментарий"]:
                lines.append("  - _комментарий:_ %s" % t["Комментарий"])
        lines.append("")
    if data["skipped"]:
        lines.append("## Не синхронизировано")
        lines.append("")
        lines.append("Задания листа, отсеянные правилами (дозаполните строку в Google-таблице):")
        lines.append("")
        for t in data["skipped"]:
            lines.append("- `%s` (строка %s) — %s%s"
                         % (t["REQ"], t["Строка"], t["Причина пропуска"],
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
            lines.append("- **%s** [%s] %s — %s _(%s%s, приоритет %s, группа %s)_"
                         % (t["REQ"], t["Статус"], t["Задание"] or "—", t["Комментарий"] or "—",
                            phase_name(version), (", %s" % t["Дата релиза"]) if t["Дата релиза"] else "",
                            t["Приоритет"], t["Группа задач"] or "—"))
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
        "| Требований синхронизировано | %d |" % data["summary"]["total"],
        "| Требований пропущено | %d |" % data["summary"]["skipped"],
        "| Фаз (релизов) | %d |" % len(data["releases"]),
        "| Фаз закрыто релизом | %d |" % sum(1 for rel in data["releases"]
                                             if rel["phase_status"] == PHASE_CLOSED),
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
    lines += ["", "## Фазы (релизы)", ""]
    lines += ["- **%s**%s — %s, требований %d (закрыто %d)"
              % (rel["phase"], (" — %s" % rel["date"]) if rel["date"] else "",
                 rel["phase_status"], len(rel["tasks"]), rel["closed"])
              for rel in data["releases"]]
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


def cmd_sync_docs(_args=None):
    rows, fieldnames, _delimiter = read_rows()
    errors, _warnings, _notes = validate(rows, fieldnames)
    if errors:
        _fail("CSV невалиден — сначала исправьте ошибки (`check`). Найдено ошибок: %d" % len(errors))

    tasks, skipped, _orphans = parse_tasks(rows)

    # Плановые документы полностью перестраиваются из текущего CSV: статусы,
    # описания, комментарии, даты и состав заданий берутся только оттуда.
    # Прошлый release-map.json нужен лишь чтобы отличить новое задание от
    # известного (гейт «новое → только Не начато») и показать, что изменилось.
    previous = previous_tasks()
    tasks, skipped = gate_new_tasks(tasks, skipped, previous)
    for line in _changed_since_last_sync(tasks, skipped, previous):
        print("ОБНОВЛЕНО: %s" % line)
    for line in _closures(tasks, previous):
        print("ЗАКРЫТО: %s" % line)

    for task in skipped:
        print("ПРОПУСК: %s (строка %s) — %s." % (task["REQ"], task["Строка"], task["Причина пропуска"]))

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


# --- Команда reset --------------------------------------------------------

def cmd_reset(_args=None):
    """Зачищает артефакты Карты релизов в .planning под новый набор заданий.

    Генерируемые блоки трёх документов переписываются пустой картой, baseline
    `.planning/release-map.json` обнуляется. Ручной текст вне маркеров (фазы GSD
    `## Phase N: ...`), `.planning/STATE.md` и `.planning/phases/` НЕ трогаются —
    текущие фазы GSD переживают зачистку.

    Baseline обнуляется, а не удаляется: пустая база — это «в GSD нет ни одного
    требования», поэтому гейт «новое требование только со статусом «Не начато»»
    работает уже на первом синке нового CSV. Удалённый файл включил бы режим
    первого запуска и втянул бы весь состав листа в любых статусах.
    """
    data = _build_release_map([], [])
    PLANNING_DIR.mkdir(parents=True, exist_ok=True)
    _write_json(RELEASE_MAP_JSON, data)
    _write_generated(PLANNING_DIR / "ROADMAP.md", _render_roadmap(data))
    _write_generated(PLANNING_DIR / "REQUIREMENTS.md", _render_requirements(data))
    _write_generated(PLANNING_DIR / "PROJECT.md", _render_project(data))
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    _write_json(OUTPUT_DIR / "release-map.json", data)

    backups = sorted(LEGACY_DIR.glob("*.bak")) if LEGACY_DIR.is_dir() else []
    if backups:
        print("ВНИМАНИЕ: в %s остались резервные копии прошлых миграций (%d) — "
              "удалите вручную, если они больше не нужны."
              % (LEGACY_DIR.relative_to(ROOT), len(backups)))
    print("reset выполнен: плановые документы очищены от Карты релизов, "
          "baseline пуст. Фазы GSD и STATE.md не тронуты.")
    print("Дальше: fetch → check → sync-docs по новому CSV.")
    return 0


# --- Команда sync (fetch + check + sync-docs) -----------------------------

def cmd_sync(args):
    print("== fetch ==")
    cmd_fetch(args)
    print("\n== check ==")
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
    print("ВНИМАНИЕ: правка сделана в локальной копии CSV. Перенесите статус в Google-таблицу — "
          "следующий `fetch` заменит локальный CSV свежей выгрузкой.")
    print("Подсказка: запустите `sync-docs`, чтобы обновить плановые документы.")
    return 0


# --- CLI ------------------------------------------------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="gsd_release_sync.py",
        description="Синхронизация Карты релизов (CSV) с артефактами GSD (Регламент, п. 8.2.1 / 11.3).",
    )
    sub = parser.add_subparsers(dest="command", required=True)
    # %% — argparse прогоняет help через %-форматирование
    sub.add_parser("fetch", help="забрать экспорт листа из «Загрузок» (%%USERPROFILE%%\\Downloads)")
    sub.add_parser("check", help="валидация CSV «Карта релизов»")
    sub.add_parser("sync-docs", help="генерация плановых артефактов из CSV")
    sub.add_parser("sync", help="fetch + check + sync-docs («Синхронизируй gsd»)")
    sub.add_parser("reset", help="очистить .planning от Карты релизов под новый набор заданий")
    p_status = sub.add_parser("status", help="смена статуса задания по ID")
    p_status.add_argument("id", help="ID задания из колонки ID")
    p_status.add_argument("status", help="новый статус: %s" % " / ".join(STATUSES))

    args = parser.parse_args(argv)
    handlers = {
        "fetch": cmd_fetch,
        "check": cmd_check,
        "sync-docs": cmd_sync_docs,
        "sync": cmd_sync,
        "reset": cmd_reset,
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
