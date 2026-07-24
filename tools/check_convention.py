#! python3
# -*- coding: utf-8 -*-
"""Чекер конвенции MM LAB — структурные правила и CLI (план 03-01).

Машинный гейт конвенции написания pyRevit-скриптов MM LAB (решение D-06).
Проверяет папки кнопок ``*.pushbutton`` и одиночные ``*.py`` без запуска
проверяемого кода: только чтение байтов и ``ast.parse`` (никаких
import/exec/eval чужого кода). Переиспользуется командой приёмки
``/mm-adopt-script`` как гейт (в режиме ``--strict``).

Вызов::

    py -3 tools/check_convention.py [PATHS...] [--all] [--strict] [--json]
        [--baseline PATH] [--write-baseline PATH] [--root PATH]

* ``PATHS`` — папки ``*.pushbutton`` ИЛИ одиночные ``*.py`` (режим «сырого
  скрипта»: только правила уровня файла MM000–MM004; структурные MM005–MM007,
  MM013 пропускаются).
* ``--all`` — обход ``<root>/MM LAB.extension/MM Lab.tab/**/*.pushbutton``
  плюс проверка орфанов layout в tab- и panel-``bundle.yaml``. Папка
  ``templates/`` в ``--all`` не входит (проверяется явным путём).
* ``--root`` — корень репозитория; по умолчанию текущая директория.
* ``--baseline PATH`` — JSON с допущенными нарушениями legacy-кнопок;
  отфильтровывает совпадающие пары (путь юнита, код правила).
* ``--write-baseline PATH`` — записать baseline из ВСЕХ текущих нарушений
  и завершиться с кодом 0.
* ``--strict`` — baseline игнорируется, warning считаются error
  (гейт приёмки, решение D-08).
* ``--json`` — только машинный вывод в stdout (ровно один JSON-объект).

Exit-коды: 0 — чисто, 1 — есть нарушения (error; в ``--strict`` — и warning),
2 — ошибка использования/внутренняя.

Правила этого модуля (AST-правила MM008–MM012 и MM014 добавляет план 03-03):

===== ======== ==============================================================
Код   Severity Проверка
===== ======== ==============================================================
MM000 error    файл не читается или не парсится ``ast.parse``
MM001 error    строка 1 script.py — ``#! python3``
MM002 error    строка 2 — ``# -*- coding: utf-8 -*-``
MM003 error    файл начинается с UTF-8 BOM (EF BB BF) — BOM запрещён
MM004 warning  docstring модуля содержит «Совместимость:» и «Зависимости:»
MM005 error    в папке кнопки есть bundle.yaml с ключами title: и tooltip:
MM006 error    в папке кнопки есть README.md
MM007 error    (а) кнопка зарегистрирована в layout родительского
               panel-bundle.yaml (кнопка вне ``*.panel`` — пропуск);
               (б) каждая запись layout имеет папку на диске, иначе орфан
MM013 warning  в папке кнопки нет мусора: __pycache__/, *.pyc, .vs/, *.csv
===== ======== ==============================================================

Ограничение парсера bundle.yaml
-------------------------------
PyYAML в stdlib нет и сторонние пакеты запрещены, поэтому bundle.yaml
разбирается ПОСТРОЧНЫМ парсером ограниченной схемы:

* ключи верхнего уровня распознаются по ``ключ:`` без отступа
  (нужны только ``title:``/``tooltip:``/``layout:``);
* записи списка ``layout:`` — строки вида ``- <имя>`` (значение
  обрезается ``.strip()`` — хвостовые пробелы в реальных файлах);
* записи из 3+ символов ``-`` или ``>`` — разделители pyRevit
  (``---``/``>>>``), папки на диске им не нужны;
* вложенные структуры, якоря, многострочные значения и прочий YAML
  НЕ поддерживаются — для файлов репозитория этого достаточно.

Совместимость: Python >= 3.10 (обычный CPython, без Revit).
Зависимости: нет (только стандартная библиотека).
"""

from __future__ import annotations

import argparse
import ast
import dataclasses
import datetime
import json
import os
import sys
from pathlib import Path

# --- константы конвенции -------------------------------------------------

SHEBANG = "#! python3"
CODING_LINE = "# -*- coding: utf-8 -*-"
BOM = b"\xef\xbb\xbf"
DOCSTRING_MARKERS = ("Совместимость:", "Зависимости:")

EXTENSION_DIR_NAME = "MM LAB.extension"
TAB_DIR_NAME = "MM Lab.tab"
PUSHBUTTON_SUFFIX = ".pushbutton"
PANEL_SUFFIX = ".panel"

JUNK_DIR_NAMES = {"__pycache__", ".vs"}
JUNK_FILE_SUFFIXES = {".pyc", ".csv"}

SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"

# Код правила -> (severity, базовое русское сообщение).
RULES: dict[str, tuple[str, str]] = {
    "MM000": (SEVERITY_ERROR, "файл не читается или не парсится"),
    "MM001": (SEVERITY_ERROR, "первая строка должна быть '#! python3'"),
    "MM002": (SEVERITY_ERROR,
              "вторая строка должна быть '# -*- coding: utf-8 -*-'"),
    "MM003": (SEVERITY_ERROR,
              "файл начинается с UTF-8 BOM (EF BB BF) — сохраните в UTF-8 без BOM"),
    "MM004": (SEVERITY_WARNING,
              "docstring модуля должен содержать «Совместимость:» и «Зависимости:»"),
    "MM005": (SEVERITY_ERROR,
              "в папке кнопки нет bundle.yaml с ключами title: и tooltip:"),
    "MM006": (SEVERITY_ERROR, "в папке кнопки нет README.md"),
    "MM007": (SEVERITY_ERROR,
              "кнопка/запись не согласована с layout bundle.yaml"),
    "MM013": (SEVERITY_WARNING,
              "мусор в папке кнопки (__pycache__/, *.pyc, .vs/, *.csv)"),
}

BASELINE_NOTE = ("Grandfathered legacy-кнопки. "
                 "При адаптации кнопки удали её запись.")


@dataclasses.dataclass
class Violation:
    """Одно нарушение конвенции.

    path — POSIX-relpath от root (папка кнопки, .py или bundle.yaml).
    """

    path: str
    code: str
    severity: str
    line: int | None
    message: str


def _violation(code: str, path, line: int | None = None,
               message: str | None = None) -> Violation:
    """Собирает Violation: severity из RULES, message — из RULES или явный."""
    severity, default_message = RULES[code]
    return Violation(
        path=str(path),
        code=code,
        severity=severity,
        line=line,
        message=default_message if message is None else message,
    )


# --- пути и чтение файлов -------------------------------------------------

def _rel_posix(path, root) -> str:
    """POSIX-путь path относительно root; вне root — абсолютный POSIX-путь."""
    resolved = Path(path).resolve()
    root_resolved = Path(root).resolve()
    try:
        return resolved.relative_to(root_resolved).as_posix()
    except ValueError:
        return resolved.as_posix()


def _read_bytes(path) -> bytes:
    with open(path, "rb") as handle:
        return handle.read()


def _read_text(path) -> str:
    """Текст в UTF-8: BOM срезается (utf-8-sig), битые байты заменяются."""
    with open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
        return handle.read()


# --- построчный парсер bundle.yaml -----------------------------------------

def _parse_bundle(path):
    """Разбирает bundle.yaml ограниченной схемы (см. docstring модуля).

    Возвращает (keys, entries): keys — множество ключей верхнего уровня,
    entries — список пар (значение записи layout после strip, номер строки).
    """
    keys: set[str] = set()
    entries: list[tuple[str, int]] = []
    in_layout = False
    try:
        text = _read_text(path)
    except OSError:
        return keys, entries
    for lineno, raw_line in enumerate(text.splitlines(), 1):
        stripped = raw_line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        indented = raw_line[:1].isspace()
        if not indented:
            if stripped.startswith("-"):
                # Список на верхнем уровне без отступа — переварим.
                if in_layout:
                    entries.append((_layout_entry_value(stripped), lineno))
                continue
            if ":" in stripped:
                key = stripped.split(":", 1)[0].strip()
                keys.add(key)
                in_layout = key == "layout"
            else:
                in_layout = False
            continue
        if in_layout and stripped.startswith("-"):
            entries.append((_layout_entry_value(stripped), lineno))
    return keys, entries


def _layout_entry_value(stripped_line: str) -> str:
    """Значение записи layout из строки вида '- имя' (с обрезкой пробелов)."""
    if stripped_line.startswith("- "):
        return stripped_line[2:].strip()
    # "-", "---", "-----" — пустая запись либо разделитель без пробела.
    return stripped_line


def _is_separator(entry: str) -> bool:
    """Разделитель layout: 3+ символов '-' (или '>' — slide-out pyRevit)."""
    return len(entry) >= 3 and (set(entry) == {"-"} or set(entry) == {">"})


def _entry_has_folder(base_dir: Path, entry: str) -> bool:
    """Есть ли на диске папка для записи layout (имя без суффикса бандла)."""
    try:
        children = list(base_dir.iterdir())
    except OSError:
        return False
    for child in children:
        if not child.is_dir():
            continue
        if child.name == entry:
            return True
        if child.name.rsplit(".", 1)[0] == entry:
            return True
    return False


# --- правила уровня файла (MM000–MM004) ------------------------------------

def _check_script_file(script_path: Path, unit_path: str) -> list[Violation]:
    """MM000–MM004 для одного файла; unit_path — значение Violation.path."""
    violations: list[Violation] = []
    try:
        raw = _read_bytes(script_path)
    except OSError as exc:
        base = RULES["MM000"][1]
        violations.append(_violation("MM000", unit_path,
                                     message=f"{base}: {exc}"))
        return violations

    if raw.startswith(BOM):
        violations.append(_violation("MM003", unit_path, line=1))

    text = raw.decode("utf-8-sig", errors="replace")
    lines = text.splitlines()
    first = lines[0].rstrip() if lines else ""
    second = lines[1].rstrip() if len(lines) > 1 else ""
    if first != SHEBANG:
        violations.append(_violation("MM001", unit_path, line=1))
    if second != CODING_LINE:
        violations.append(_violation("MM002", unit_path, line=2))

    try:
        tree = ast.parse(text, filename=str(script_path))
    except (SyntaxError, ValueError) as exc:
        base = RULES["MM000"][1]
        detail = getattr(exc, "msg", None) or str(exc)
        violations.append(_violation(
            "MM000", unit_path,
            line=getattr(exc, "lineno", None),
            message=f"{base}: {detail}",
        ))
        return violations  # дальше файл не проверяем

    docstring = ast.get_docstring(tree)
    if not docstring or any(marker not in docstring
                            for marker in DOCSTRING_MARKERS):
        violations.append(_violation("MM004", unit_path))
    return violations


# --- структурные правила кнопки (MM005–MM007а, MM013) ----------------------

def _check_bundle_yaml(button_dir: Path, unit_path: str) -> list[Violation]:
    """MM005: bundle.yaml кнопки существует и содержит title:/tooltip:."""
    bundle = button_dir / "bundle.yaml"
    if not bundle.is_file():
        return [_violation(
            "MM005", unit_path,
            message="в папке кнопки нет bundle.yaml (нужны ключи title: и tooltip:)",
        )]
    keys, _entries = _parse_bundle(bundle)
    missing = [key for key in ("title", "tooltip") if key not in keys]
    if missing:
        listed = ", ".join(f"{key}:" for key in missing)
        return [_violation(
            "MM005", unit_path,
            message=f"в bundle.yaml кнопки нет ключей: {listed}",
        )]
    return []


def _check_layout_registration(button_dir: Path, root: Path,
                               unit_path: str) -> list[Violation]:
    """MM007(а): кнопка зарегистрирована в layout родительской панели.

    Кнопка вне папки ``*.panel`` (например, templates/) — правило пропускается.
    """
    parent = button_dir.parent
    if not parent.name.endswith(PANEL_SUFFIX):
        return []
    button_name = button_dir.name
    if button_name.endswith(PUSHBUTTON_SUFFIX):
        button_name = button_name[: -len(PUSHBUTTON_SUFFIX)]
    panel_bundle = parent / "bundle.yaml"
    bundle_rel = _rel_posix(panel_bundle, root)
    if not panel_bundle.is_file():
        return [_violation(
            "MM007", unit_path,
            message=(f"кнопка «{button_name}» не зарегистрирована: "
                     f"в панели нет bundle.yaml с layout"),
        )]
    _keys, entries = _parse_bundle(panel_bundle)
    layout_names = {entry for entry, _line in entries
                    if entry and not _is_separator(entry)}
    if button_name not in layout_names:
        return [_violation(
            "MM007", unit_path,
            message=(f"кнопка «{button_name}» отсутствует в layout "
                     f"{bundle_rel}"),
        )]
    return []


def _check_junk(button_dir: Path, unit_path: str) -> list[Violation]:
    """MM013: мусор в папке кнопки (__pycache__/, *.pyc, .vs/, *.csv)."""
    violations: list[Violation] = []
    for dirpath, dirnames, filenames in os.walk(button_dir):
        current = Path(dirpath)
        for junk_dir in sorted(name for name in dirnames
                               if name in JUNK_DIR_NAMES):
            rel = (current / junk_dir).relative_to(button_dir).as_posix()
            violations.append(_violation(
                "MM013", unit_path,
                message=f"мусор в папке кнопки: {rel}/",
            ))
        # В мусорные папки не спускаемся — не плодим дубли на *.pyc внутри.
        dirnames[:] = [name for name in dirnames
                       if name not in JUNK_DIR_NAMES]
        for file_name in sorted(filenames):
            if Path(file_name).suffix.lower() in JUNK_FILE_SUFFIXES:
                rel = (current / file_name).relative_to(button_dir).as_posix()
                violations.append(_violation(
                    "MM013", unit_path,
                    message=f"мусор в папке кнопки: {rel}",
                ))
    return violations


# --- публичные проверки ----------------------------------------------------

def check_pushbutton(button_dir, root) -> list[Violation]:
    """Полная проверка папки кнопки ``*.pushbutton`` (юнит конвенции)."""
    button_path = Path(button_dir).resolve()
    root_path = Path(root).resolve()
    unit_path = _rel_posix(button_path, root_path)
    violations: list[Violation] = []
    violations.extend(_check_script_file(button_path / "script.py", unit_path))
    violations.extend(_check_bundle_yaml(button_path, unit_path))
    if not (button_path / "README.md").is_file():
        violations.append(_violation("MM006", unit_path))
    violations.extend(_check_layout_registration(button_path, root_path,
                                                 unit_path))
    violations.extend(_check_junk(button_path, unit_path))
    return violations


def check_script(script_path) -> list[Violation]:
    """Режим «сырого скрипта»: только правила уровня файла MM000–MM004.

    Структурные правила (MM005–MM007, MM013) не применяются — у одиночного
    .py нет папки кнопки, bundle.yaml и layout.
    """
    path = Path(script_path)
    return _check_script_file(path, path.as_posix())


def check_layouts(root) -> list[Violation]:
    """MM007(б): орфаны layout в tab- и panel-``bundle.yaml``.

    Каждая запись layout (кроме разделителей) обязана иметь папку на диске;
    нарушение приписывается соответствующему bundle.yaml.
    """
    root_path = Path(root).resolve()
    tab_dir = root_path / EXTENSION_DIR_NAME / TAB_DIR_NAME
    violations: list[Violation] = []
    if not tab_dir.is_dir():
        return violations
    bundles: list[Path] = []
    tab_bundle = tab_dir / "bundle.yaml"
    if tab_bundle.is_file():
        bundles.append(tab_bundle)
    for panel_dir in sorted(child for child in tab_dir.iterdir()
                            if child.is_dir()
                            and child.name.endswith(PANEL_SUFFIX)):
        panel_bundle = panel_dir / "bundle.yaml"
        if panel_bundle.is_file():
            bundles.append(panel_bundle)
    for bundle in bundles:
        bundle_rel = _rel_posix(bundle, root_path)
        _keys, entries = _parse_bundle(bundle)
        for entry, lineno in entries:
            if not entry or _is_separator(entry):
                continue
            if not _entry_has_folder(bundle.parent, entry):
                violations.append(_violation(
                    "MM007", bundle_rel, line=lineno,
                    message=(f"запись layout «{entry}» не имеет папки "
                             f"на диске (орфан)"),
                ))
    return violations


def iter_pushbuttons(root) -> list[Path]:
    """Все папки ``*.pushbutton`` в ``<root>/MM LAB.extension/MM Lab.tab``.

    Пути за пределами root (симлинки и т.п.) отбрасываются (гейт V5).
    """
    root_path = Path(root).resolve()
    tab_dir = root_path / EXTENSION_DIR_NAME / TAB_DIR_NAME
    if not tab_dir.is_dir():
        return []
    buttons: list[Path] = []
    for candidate in sorted(tab_dir.rglob("*" + PUSHBUTTON_SUFFIX)):
        if not candidate.is_dir():
            continue
        resolved = candidate.resolve()
        try:
            resolved.relative_to(root_path)
        except ValueError:
            continue
        buttons.append(resolved)
    return buttons


# --- baseline ---------------------------------------------------------------

def load_baseline(path) -> dict:
    """Читает baseline JSON (схема: generated/note/units)."""
    with open(path, "r", encoding="utf-8-sig") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("baseline должен быть JSON-объектом")
    return data


def apply_baseline(violations, baseline) -> list[Violation]:
    """Отфильтровывает нарушения, допущенные baseline (пары путь+код)."""
    units = baseline.get("units", {}) if isinstance(baseline, dict) else {}
    allowed: set[tuple[str, str]] = set()
    for unit_path, codes in units.items():
        for code in codes:
            allowed.add((str(unit_path), str(code)))
    return [violation for violation in violations
            if (violation.path, violation.code) not in allowed]


def write_baseline(violations, path) -> None:
    """Пишет baseline из всех переданных нарушений (units: путь -> коды)."""
    units: dict[str, set[str]] = {}
    for violation in violations:
        units.setdefault(violation.path, set()).add(violation.code)
    data = {
        "generated": datetime.date.today().isoformat(),
        "note": BASELINE_NOTE,
        "units": {unit_path: sorted(codes)
                  for unit_path, codes in sorted(units.items())},
    }
    with open(path, "w", encoding="utf-8", newline="\n") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


# --- CLI --------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="check_convention.py",
        description=("Чекер конвенции MM LAB: структурные правила "
                     "MM000–MM007, MM013 для pyRevit-кнопок."),
    )
    parser.add_argument(
        "paths", nargs="*", metavar="PATH",
        help="папки *.pushbutton или одиночные файлы *.py",
    )
    parser.add_argument(
        "--all", action="store_true",
        help=("проверить все кнопки в <root>/MM LAB.extension/MM Lab.tab "
              "и орфаны layout"),
    )
    parser.add_argument(
        "--strict", action="store_true",
        help="игнорировать baseline; warning считать error (гейт приёмки)",
    )
    parser.add_argument(
        "--json", action="store_true",
        help="только машинный JSON-вывод в stdout",
    )
    parser.add_argument(
        "--baseline", metavar="PATH",
        help="JSON с допущенными нарушениями legacy-кнопок",
    )
    parser.add_argument(
        "--write-baseline", dest="write_baseline", metavar="PATH",
        help="записать baseline из всех текущих нарушений и выйти с кодом 0",
    )
    parser.add_argument(
        "--root", metavar="PATH",
        help="корень репозитория (по умолчанию — текущая директория)",
    )
    return parser


def _collect_targets(args, root: Path):
    """Возвращает (buttons, scripts) либо кидает SystemExit(2) при ошибке."""
    buttons: list[Path] = []
    scripts: list[Path] = []
    if args.all:
        buttons.extend(iter_pushbuttons(root))
    for raw_path in args.paths:
        resolved = Path(raw_path).resolve()
        if resolved.is_dir() and resolved.name.endswith(PUSHBUTTON_SUFFIX):
            buttons.append(resolved)
        elif resolved.is_file() and resolved.suffix == ".py":
            scripts.append(resolved)
        else:
            print(f"Ошибка: путь не является папкой *{PUSHBUTTON_SUFFIX} "
                  f"или файлом *.py: {raw_path}", file=sys.stderr)
            raise SystemExit(2)
    # Дедупликация с сохранением порядка (--all + явный путь).
    buttons = list(dict.fromkeys(buttons))
    scripts = list(dict.fromkeys(scripts))
    return buttons, scripts


def main(argv=None) -> int:
    """Точка входа CLI. Возвращает exit-код (0/1/2)."""
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError, OSError):
        pass  # io.StringIO в тестах / экзотические консоли

    parser = _build_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:  # argparse сам печатает usage
        code = exc.code
        return code if isinstance(code, int) else 2

    root = Path(args.root).resolve() if args.root else Path.cwd().resolve()

    if not args.paths and not args.all:
        print("Ошибка: укажите пути к *.pushbutton/*.py или флаг --all",
              file=sys.stderr)
        return 2

    try:
        buttons, scripts = _collect_targets(args, root)
    except SystemExit as exc:
        code = exc.code
        return code if isinstance(code, int) else 2

    violations: list[Violation] = []
    checked = 0
    for button_dir in buttons:
        violations.extend(check_pushbutton(button_dir, root))
        checked += 1
    for script_path in scripts:
        script_violations = check_script(script_path)
        rel = _rel_posix(script_path, root)
        for violation in script_violations:
            violation.path = rel
        violations.extend(script_violations)
        checked += 1
    if args.all:
        violations.extend(check_layouts(root))

    if args.write_baseline:
        try:
            write_baseline(violations, args.write_baseline)
        except OSError as exc:
            print(f"Ошибка записи baseline: {exc}", file=sys.stderr)
            return 2
        if not args.json:
            print(f"Baseline записан: {args.write_baseline} "
                  f"(нарушений: {len(violations)})")
        return 0

    if args.baseline and not args.strict:
        try:
            baseline = load_baseline(args.baseline)
        except (OSError, ValueError) as exc:
            # ValueError покрывает и json.JSONDecodeError (его подкласс).
            print(f"Ошибка чтения baseline: {exc}", file=sys.stderr)
            return 2
        violations = apply_baseline(violations, baseline)

    errors = sum(1 for violation in violations
                 if violation.severity == SEVERITY_ERROR)
    warnings = sum(1 for violation in violations
                   if violation.severity == SEVERITY_WARNING)

    if args.json:
        payload = {
            "checked": checked,
            "errors": errors,
            "warnings": warnings,
            "violations": [dataclasses.asdict(violation)
                           for violation in violations],
        }
        print(json.dumps(payload, ensure_ascii=False))
    else:
        for violation in violations:
            line_part = (f"строка {violation.line}: "
                         if violation.line is not None else "")
            print(f"{violation.path}: {violation.code} "
                  f"[{violation.severity}] {line_part}{violation.message}")
        print(f"Проверено: {checked}, ошибок: {errors}, "
              f"предупреждений: {warnings}")

    if errors > 0:
        return 1
    if args.strict and warnings > 0:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
