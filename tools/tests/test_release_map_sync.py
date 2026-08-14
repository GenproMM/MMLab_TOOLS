# -*- coding: utf-8 -*-
"""Тесты парсера «Карты релизов» (RELEASE_MAP/gsd_release_sync.py).

Исполняемая спецификация правил синхронизации листа «Скрипты_Карта релизов»:

    * версия релиза берётся из строки-заголовка и наследуется заданиями ниже;
    * задание со статусом «Без статуса» (или пустым) НЕ синхронизируется;
    * задание с пустым «Название плагина» НЕ синхронизируется;
    * «MVP = TRUE» → приоритет MVP, иначе «Обычный»;
    * check ловит ошибки структуры (задание до релиза, дубль ID, битая версия);
    * регенерация .planning/*.md сохраняет ручной текст вне маркеров.

Запуск:
    py -3 -m unittest discover -s tools/tests -p "test_release_map*.py" -q
"""

import contextlib
import importlib.util
import io
import tempfile
import unittest
from pathlib import Path

# tools/tests/ -> tools/ -> корень репозитория
REPO = Path(__file__).resolve().parents[2]
MODULE_PATH = REPO / "RELEASE_MAP" / "gsd_release_sync.py"


def _load_module():
    """Импортирует скрипт по пути: RELEASE_MAP/ не является пакетом."""
    spec = importlib.util.spec_from_file_location("gsd_release_sync", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


sync = _load_module()

FIELDNAMES = list(sync.REQUIRED_COLUMNS)


def _row(**kwargs):
    """Строка листа: все колонки пустые, кроме переданных."""
    row = {name: "" for name in FIELDNAMES}
    row.update(kwargs)
    return row


def _release_header(version):
    return _row(**{sync.COL_RELEASE: version})


def _task(rid, plugin="Плагин А", status="Не начато", mvp="TRUE", desc="Задание", group="ФУНКЦИОНАЛ"):
    return _row(**{
        sync.COL_ID: rid,
        sync.COL_PLUGIN: plugin,
        sync.COL_STATUS: status,
        sync.COL_MVP: mvp,
        sync.COL_DESC: desc,
        sync.COL_GROUP: group,
    })


class TestParseTasks(unittest.TestCase):
    """Разбор листа: наследование релиза и правила отсева."""

    def test_release_inherited_from_header_row(self):
        """Задания получают версию релиза из ближайшей строки-заголовка выше."""
        rows = [
            _release_header("v250407"),
            _task("2102001"),
            _release_header("v251205"),
            _task("2102010"),
            _task("2102011"),
        ]
        tasks, skipped, orphans = sync.parse_tasks(rows)
        self.assertEqual([], skipped)
        self.assertEqual([], orphans)
        self.assertEqual(
            [("2102001", "v250407"), ("2102010", "v251205"), ("2102011", "v251205")],
            [(t["ID"], t["Релиз"]) for t in tasks],
        )

    def test_status_bez_statusa_not_synced(self):
        """«Без статуса» и пустой статус не синхронизируются."""
        rows = [
            _release_header("v250407"),
            _task("2102005", status=sync.STATUS_NONE),
            _task("2102006", status=""),
            _task("2102007", status="Готово"),
        ]
        tasks, skipped, _ = sync.parse_tasks(rows)
        self.assertEqual(["2102007"], [t["ID"] for t in tasks])
        self.assertEqual(
            [("2102005", sync.SKIP_NO_STATUS), ("2102006", sync.SKIP_NO_STATUS)],
            [(t["ID"], t["Причина пропуска"]) for t in skipped],
        )

    def test_task_without_plugin_not_synced(self):
        """Задание без «Название плагина» не синхронизируется даже со статусом."""
        rows = [
            _release_header("v250625"),
            _task("2102006", plugin="", status="Релиз"),
        ]
        tasks, skipped, _ = sync.parse_tasks(rows)
        self.assertEqual([], tasks)
        self.assertEqual([("2102006", sync.SKIP_NO_PLUGIN)],
                         [(t["ID"], t["Причина пропуска"]) for t in skipped])

    def test_no_status_wins_over_no_plugin(self):
        """У строки без статуса И без плагина причина пропуска — статус (проверяется первым)."""
        rows = [_release_header("v250613"), _task("2102005", plugin="", status=sync.STATUS_NONE)]
        _tasks, skipped, _ = sync.parse_tasks(rows)
        self.assertEqual(sync.SKIP_NO_STATUS, skipped[0]["Причина пропуска"])

    def test_mvp_flag_maps_to_priority(self):
        """MVP=TRUE → приоритет MVP; FALSE и пусто → «Обычный»."""
        rows = [
            _release_header("v251205"),
            _task("2102015", mvp="TRUE"),
            _task("2102014", mvp="FALSE"),
            _task("2102013", mvp=""),
        ]
        tasks, _, _ = sync.parse_tasks(rows)
        self.assertEqual(
            [("2102015", sync.PRIORITY_MVP), ("2102014", sync.PRIORITY_REGULAR),
             ("2102013", sync.PRIORITY_REGULAR)],
            [(t["ID"], t["Приоритет"]) for t in tasks],
        )

    def test_task_before_first_release_is_orphan(self):
        """Задание до первой строки-заголовка релиза — сирота, а не задание."""
        rows = [_task("2102001"), _release_header("v250407"), _task("2102002")]
        tasks, _, orphans = sync.parse_tasks(rows)
        self.assertEqual(["2102002"], [t["ID"] for t in tasks])
        self.assertEqual(["2102001"], [t["ID"] for t in orphans])


class TestValidate(unittest.TestCase):
    """check: ошибки структуры листа, предупреждения и пропуски."""

    def test_valid_sheet_has_no_errors(self):
        rows = [_release_header("v250407"), _task("2102001", status="Готово")]
        errors, warnings, notes = sync.validate(rows, FIELDNAMES)
        self.assertEqual([], errors)
        self.assertEqual([], warnings)
        self.assertEqual([], notes)

    def test_missing_columns_is_error(self):
        rows = [_release_header("v250407")]
        errors, _, _ = sync.validate(rows, [sync.COL_ID, sync.COL_STATUS])
        self.assertEqual(1, len(errors))
        self.assertIn("отсутствуют обязательные колонки", errors[0])

    def test_bad_release_version_is_error(self):
        rows = [_release_header("2025-04-07"), _task("2102001")]
        errors, _, _ = sync.validate(rows, FIELDNAMES)
        self.assertTrue(any("vYYMMDD" in e for e in errors), errors)

    def test_duplicate_id_is_error(self):
        rows = [_release_header("v250407"), _task("2102001"), _task("2102001")]
        errors, _, _ = sync.validate(rows, FIELDNAMES)
        self.assertTrue(any("дубль ID" in e for e in errors), errors)

    def test_unknown_status_is_error(self):
        rows = [_release_header("v250407"), _task("2102001", status="В процессе")]
        errors, _, _ = sync.validate(rows, FIELDNAMES)
        self.assertTrue(any("недопустимый Статус" in e for e in errors), errors)

    def test_orphan_task_is_error(self):
        rows = [_task("2102001"), _release_header("v250407")]
        errors, _, _ = sync.validate(rows, FIELDNAMES)
        self.assertTrue(any("до первой строки-заголовка релиза" in e for e in errors), errors)

    def test_skipped_tasks_are_notes_not_errors(self):
        """Пропущенные задания — информационные строки: check остаётся валидным."""
        rows = [_release_header("v250407"), _task("2102001", status="Готово"),
                _task("2102002", plugin="")]
        errors, _, notes = sync.validate(rows, FIELDNAMES)
        self.assertEqual([], errors)
        self.assertEqual(1, len(notes))
        self.assertIn(sync.SKIP_NO_PLUGIN, notes[0])

    def test_empty_description_is_warning(self):
        rows = [_release_header("v250407"), _task("2102014", desc="")]
        errors, warnings, _ = sync.validate(rows, FIELDNAMES)
        self.assertEqual([], errors)
        self.assertTrue(any("пустое" in w for w in warnings), warnings)


class TestManualTextPreserved(unittest.TestCase):
    """Регенерация .planning/*.md не затирает ручной текст вне маркеров."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)
        # генератор печатает пути относительно ROOT и пишет бэкапы в LEGACY_DIR
        self._orig = (sync.ROOT, sync.LEGACY_DIR)
        sync.ROOT, sync.LEGACY_DIR = tmp, tmp / "_legacy"
        self.path = tmp / "ROADMAP.md"

    def tearDown(self):
        sync.ROOT, sync.LEGACY_DIR = self._orig
        self._tmp.cleanup()

    @staticmethod
    def _write_quietly(path, text):
        """_write_generated печатает отчёт в stdout — в тестах он лишний шум."""
        with contextlib.redirect_stdout(io.StringIO()):
            sync._write_generated(path, text)

    def test_migration_keeps_manual_section_and_drops_generated(self):
        """Файл без маркеров: секции «## Релиз ...» уходят, фаза GSD остаётся."""
        self.path.write_text(
            "# ROADMAP\n\n> Сгенерировано `x`.\n\n"
            "## Релиз v250407\n\n- [x] старое задание\n\n"
            "## Phase 3: Конвенция\n\n**Plans:** 7/7\n",
            encoding="utf-8",
        )
        self._write_quietly(self.path, "# ROADMAP\n\n## Релиз v251205\n\n- [ ] новое задание\n")
        text = self.path.read_text(encoding="utf-8")
        self.assertIn("## Phase 3: Конвенция", text)
        self.assertIn("**Plans:** 7/7", text)
        self.assertIn("## Релиз v251205", text)
        self.assertNotIn("старое задание", text)
        self.assertTrue(list(sync.LEGACY_DIR.glob("ROADMAP.md.*.bak")), "нет резервной копии")

    def test_regeneration_replaces_only_marked_block(self):
        """Файл с маркерами: заменяется только блок между маркерами."""
        self._write_quietly(self.path, "# ROADMAP\n\n## Релиз v250407\n")
        with self.path.open("a", encoding="utf-8") as fh:
            fh.write("\n## Phase 3: Конвенция\n\nручной текст\n")
        self._write_quietly(self.path, "# ROADMAP\n\n## Релиз v251205\n")
        text = self.path.read_text(encoding="utf-8")
        self.assertIn("## Релиз v251205", text)
        self.assertNotIn("## Релиз v250407", text)
        self.assertIn("ручной текст", text)
        self.assertEqual(1, text.count(sync.MARK_BEGIN))
        self.assertEqual(1, text.count(sync.MARK_END))
        self.assertEqual([], list(sync.LEGACY_DIR.glob("*.bak")),
                         "повторная генерация не должна плодить резервные копии")


class TestRealSheet(unittest.TestCase):
    """Реальный CSV репозитория проходит check и разбирается ожидаемо."""

    def test_repo_csv_is_valid(self):
        rows, fieldnames, _ = sync.read_rows()
        errors, _, _ = sync.validate(rows, fieldnames)
        self.assertEqual([], errors, "CSV «Карта релизов» в репозитории невалиден")

    def test_repo_csv_has_synced_and_skipped_tasks(self):
        rows, _, _ = sync.read_rows()
        tasks, skipped, orphans = sync.parse_tasks(rows)
        self.assertEqual([], orphans)
        self.assertTrue(tasks, "ни одно задание реального листа не синхронизируется")
        for task in tasks:
            self.assertTrue(task["Плагин"], "задание %s без плагина попало в синк" % task["ID"])
            self.assertIn(task["Статус"], sync.STATUSES)
        for task in skipped:
            self.assertIn(task["Причина пропуска"], (sync.SKIP_NO_STATUS, sync.SKIP_NO_PLUGIN))


if __name__ == "__main__":
    unittest.main()
