# -*- coding: utf-8 -*-
"""Тесты парсера «Карты релизов» (RELEASE_MAP/gsd_release_sync.py).

Исполняемая спецификация правил синхронизации листа «Скрипты_Карта релизов»:

    * версия релиза берётся из строки-заголовка и наследуется заданиями ниже;
    * задание со статусом «Без статуса» (или пустым) НЕ синхронизируется;
    * задание с пустым «Название плагина» НЕ синхронизируется;
    * «MVP = TRUE» → приоритет MVP, иначе «Обычный»;
    * check ловит ошибки структуры (задание до релиза, дубль ID, битая версия);
    * регенерация .planning/*.md сохраняет ручной текст вне маркеров;
    * fetch забирает свежий экспорт листа из «Загрузок» (%USERPROFILE%\\Downloads);
    * все поля заведённых заданий обновляются из нового CSV (прошлый
      release-map.json ничего не переопределяет);
    * новое задание заводится в GSD только со статусом «Не начато».

Запуск:
    py -3 -m unittest discover -s tools/tests -p "test_release_map*.py" -q
"""

import contextlib
import importlib.util
import io
import json
import os
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


class PlanningSandbox(unittest.TestCase):
    """База для тестов sync-docs/reset: изолированные CSV и .planning во временном каталоге."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)
        self._orig = (sync.ROOT, sync.CSV_PATH, sync.PLANNING_DIR, sync.LEGACY_DIR,
                      sync.RELEASE_MAP_JSON, sync.OUTPUT_DIR)
        sync.ROOT = tmp
        (tmp / "RELEASE_MAP").mkdir()
        sync.CSV_PATH = tmp / "RELEASE_MAP" / sync.CSV_NAME
        sync.OUTPUT_DIR = tmp / "RELEASE_MAP" / "output"
        sync.PLANNING_DIR = tmp / ".planning"
        sync.LEGACY_DIR = sync.PLANNING_DIR / "_legacy"
        sync.RELEASE_MAP_JSON = sync.PLANNING_DIR / "release-map.json"

    def tearDown(self):
        (sync.ROOT, sync.CSV_PATH, sync.PLANNING_DIR, sync.LEGACY_DIR,
         sync.RELEASE_MAP_JSON, sync.OUTPUT_DIR) = self._orig
        self._tmp.cleanup()

    def _write_csv(self, rows):
        sync.write_rows(rows, FIELDNAMES, ",", sync.CSV_PATH)

    def _sync(self):
        """Прогон sync-docs; возвращает (release-map.json, stdout)."""
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            sync.cmd_sync_docs(None)
        return json.loads(sync.RELEASE_MAP_JSON.read_text(encoding="utf-8")), out.getvalue()

    @staticmethod
    def _task_by_id(data, rid):
        for rel in data["releases"]:
            for task in rel["tasks"]:
                if task["ID"] == rid:
                    return task
        return None


class TestSyncDocsFromCsv(PlanningSandbox):
    """sync-docs: CSV — единственный источник значений; гейт новых заданий."""

    def test_all_fields_updated_from_new_csv(self):
        """Второй синк подхватывает новые статус, описание, комментарий, автора, вес."""
        self._write_csv([_release_header("v250407"),
                         _task("2102001", status="Не начато", desc="старое описание")])
        self._sync()

        self._write_csv([_release_header("v250407"),
                         _row(**{sync.COL_ID: "2102001", sync.COL_PLUGIN: "Плагин Б",
                                 sync.COL_STATUS: "Готово", sync.COL_MVP: "FALSE",
                                 sync.COL_DESC: "новое описание", sync.COL_GROUP: "ОШИБКИ",
                                 sync.COL_COMMENT: "новый комментарий", sync.COL_AUTHOR: "Автор",
                                 sync.COL_WEIGHT: "5"})])
        data, out = self._sync()

        task = self._task_by_id(data, "2102001")
        self.assertEqual("Готово", task["Статус"])
        self.assertEqual("новое описание", task["Задание"])
        self.assertEqual("новый комментарий", task["Комментарий"])
        self.assertEqual("Плагин Б", task["Плагин"])
        self.assertEqual("Автор", task["Автор"])
        self.assertEqual("5", task["Вес"])
        self.assertEqual("ОШИБКИ", task["Группа задач"])
        self.assertEqual(sync.PRIORITY_REGULAR, task["Приоритет"])
        self.assertIn("ОБНОВЛЕНО", out)

    def test_status_rollback_in_csv_is_applied(self):
        """Откат статуса в CSV принимается: прошлый release-map.json не переопределяет."""
        self._write_csv([_release_header("v250407"), _task("2102001", status="Готово")])
        self._sync()
        self._write_csv([_release_header("v250407"), _task("2102001", status="В работе")])
        data, _ = self._sync()
        self.assertEqual("В работе", self._task_by_id(data, "2102001")["Статус"])

    def test_sync_docs_does_not_write_csv(self):
        """sync-docs не трогает CSV: параметры меняются только в Карте релизов."""
        self._write_csv([_release_header("v250407"), _task("2102001", status="Готово")])
        self._sync()
        before = sync.CSV_PATH.read_bytes()
        self._write_csv([_release_header("v250407"), _task("2102001", status="Не начато")])
        expected = sync.CSV_PATH.read_bytes()
        self.assertNotEqual(before, expected)
        self._sync()
        self.assertEqual(expected, sync.CSV_PATH.read_bytes())

    def test_task_removed_from_csv_disappears(self):
        """Задание, удалённое из листа, уходит из плановых документов."""
        self._write_csv([_release_header("v250407"), _task("2102001"), _task("2102002")])
        self._sync()
        self._write_csv([_release_header("v250407"), _task("2102001")])
        data, out = self._sync()
        self.assertIsNone(self._task_by_id(data, "2102002"))
        self.assertIn("2102002 удалено из листа", out)

    def test_new_task_enters_only_as_not_started(self):
        """Новое задание со статусом «Готово» в GSD не заводится."""
        self._write_csv([_release_header("v250407"), _task("2102001")])
        self._sync()
        self._write_csv([_release_header("v250407"), _task("2102001"),
                         _task("2102099", status="Готово")])
        data, _ = self._sync()
        self.assertIsNone(self._task_by_id(data, "2102099"))
        self.assertEqual([("2102099", sync.SKIP_NEW_NOT_STARTED)],
                         [(t["ID"], t["Причина пропуска"]) for t in data["skipped"]])

    def test_new_task_with_not_started_is_added(self):
        """Новое задание со статусом «Не начато» заводится в GSD."""
        self._write_csv([_release_header("v250407"), _task("2102001")])
        self._sync()
        self._write_csv([_release_header("v250407"), _task("2102001"),
                         _task("2102099", status=sync.STATUS_NEW)])
        data, out = self._sync()
        self.assertIsNotNone(self._task_by_id(data, "2102099"))
        self.assertIn("новое требование REQ-2102099", out)

    def test_known_task_keeps_updating_in_any_status(self):
        """Гейт «Не начато» действует только на новые: заведённое задание идёт дальше по циклу."""
        self._write_csv([_release_header("v250407"), _task("2102001", status=sync.STATUS_NEW)])
        self._sync()
        for status in ("В работе", "Готово", "Релиз"):
            self._write_csv([_release_header("v250407"), _task("2102001", status=status)])
            data, _ = self._sync()
            self.assertEqual(status, self._task_by_id(data, "2102001")["Статус"])

    def test_first_sync_without_baseline_keeps_all_statuses(self):
        """Первый синк (нет release-map.json) заводит весь состав, а не только «Не начато»."""
        self._write_csv([_release_header("v250407"),
                         _task("2102001", status="Релиз"), _task("2102002", status="Готово")])
        data, _ = self._sync()
        self.assertEqual(2, data["summary"]["total"])

    def test_release_is_gsd_phase_and_task_is_req(self):
        """Релиз → фаза «Phase vYYMMDD», задание → требование «REQ-<ID>»."""
        self._write_csv([_release_header("v260814"), _task("2830001")])
        data, _ = self._sync()
        self.assertEqual("Phase v260814", data["releases"][0]["phase"])
        self.assertEqual("REQ-2830001", self._task_by_id(data, "2830001")["REQ"])
        roadmap = (sync.PLANNING_DIR / "ROADMAP.md").read_text(encoding="utf-8")
        self.assertIn("## Phase v260814 — 2026-08-14", roadmap)
        self.assertIn("**Требования:** REQ-2830001", roadmap)
        requirements = (sync.PLANNING_DIR / "REQUIREMENTS.md").read_text(encoding="utf-8")
        self.assertIn("**REQ-2830001**", requirements)

    def test_release_date_comes_from_version_only(self):
        """Дата фазы выводится из версии релиза; «План релиза» не импортируется."""
        header = _release_header("v260814")
        header[sync.COL_PLAN] = "01.01.2000"  # рудимент листа — должен игнорироваться
        self._write_csv([header, _task("2830001")])
        data, _ = self._sync()
        self.assertEqual("2026-08-14", data["releases"][0]["date"])
        self.assertEqual("2026-08-14", self._task_by_id(data, "2830001")["Дата релиза"])
        self.assertNotIn("План релиза", self._task_by_id(data, "2830001"))
        self.assertNotIn("01.01.2000", (sync.PLANNING_DIR / "ROADMAP.md").read_text(encoding="utf-8"))

    def test_impossible_date_in_version_is_error(self):
        """Версия vYYMMDD с несуществующей датой — ошибка check."""
        errors, _, _ = sync.validate([_release_header("v250632"), _task("2102001")], FIELDNAMES)
        self.assertTrue(any("несуществующая дата" in e for e in errors), errors)

    def test_release_status_closes_requirement_and_phase(self):
        """Статус «Релиз» закрывает требование и — когда закрыты все — фазу."""
        self._write_csv([_release_header("v260814"),
                         _task("2830001", status="Готово"), _task("2830002", status="Готово")])
        data, _ = self._sync()
        self.assertEqual(sync.PHASE_READY, data["releases"][0]["phase_status"])
        self.assertNotIn("[x]", (sync.PLANNING_DIR / "ROADMAP.md").read_text(encoding="utf-8"))

        self._write_csv([_release_header("v260814"),
                         _task("2830001", status="Релиз"), _task("2830002", status="Готово")])
        data, out = self._sync()
        # работа по обоим требованиям сделана, но релизом закрыто пока одно
        self.assertEqual(sync.PHASE_READY, data["releases"][0]["phase_status"])
        self.assertEqual(1, data["releases"][0]["closed"])
        self.assertIn("требование REQ-2830001 закрыто релизом v260814", out)

        self._write_csv([_release_header("v260814"),
                         _task("2830001", status="Релиз"), _task("2830002", status="Релиз")])
        data, out = self._sync()
        self.assertEqual(sync.PHASE_CLOSED, data["releases"][0]["phase_status"])
        self.assertIn("фаза Phase v260814 закрыта", out)
        roadmap = (sync.PLANNING_DIR / "ROADMAP.md").read_text(encoding="utf-8")
        self.assertIn("**Статус фазы:** %s — закрыто 2 из 2" % sync.PHASE_CLOSED, roadmap)

    def test_gotovo_does_not_close_requirement(self):
        """«Готово» не закрывает требование: отметка [x] только по «Релиз»."""
        self._write_csv([_release_header("v260814"), _task("2830001", status="Готово")])
        self._sync()
        roadmap = (sync.PLANNING_DIR / "ROADMAP.md").read_text(encoding="utf-8")
        self.assertIn("- [ ] **REQ-2830001**", roadmap)

    def test_manual_gsd_phase_survives_regeneration(self):
        """Ручная фаза GSD «## Phase 3: ...» не путается с генерируемой «## Phase vYYMMDD»."""
        self._write_csv([_release_header("v260814"), _task("2830001")])
        self._sync()
        roadmap = sync.PLANNING_DIR / "ROADMAP.md"
        with roadmap.open("a", encoding="utf-8") as fh:
            fh.write("\n## Phase 3: Конвенция\n\nручной текст\n")
        self._write_csv([_release_header("v260814"), _task("2830001", status="Релиз")])
        self._sync()
        text = roadmap.read_text(encoding="utf-8")
        self.assertIn("## Phase 3: Конвенция", text)
        self.assertIn("ручной текст", text)
        self.assertIn("## Phase v260814", text)


class TestReset(PlanningSandbox):
    """reset: зачистка .planning под новый набор заданий, фазы GSD не трогаются."""

    def _reset(self):
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            sync.cmd_reset(None)
        return json.loads(sync.RELEASE_MAP_JSON.read_text(encoding="utf-8")), out.getvalue()

    def test_reset_clears_generated_blocks(self):
        """Требования прошлого набора уходят из всех трёх документов."""
        self._write_csv([_release_header("v250407"), _task("2102001", status="Релиз")])
        self._sync()
        data, _ = self._reset()
        self.assertEqual([], data["releases"])
        self.assertEqual(0, data["summary"]["total"])
        for name in ("ROADMAP.md", "REQUIREMENTS.md", "PROJECT.md"):
            text = (sync.PLANNING_DIR / name).read_text(encoding="utf-8")
            self.assertNotIn("2102001", text, "%s: остался ID прошлого набора" % name)
            self.assertNotIn("REQ-2102001", text, "%s: осталось требование" % name)

    def test_reset_keeps_manual_gsd_phases(self):
        """Ручные фазы GSD вне маркеров переживают зачистку."""
        self._write_csv([_release_header("v250407"), _task("2102001")])
        self._sync()
        roadmap = sync.PLANNING_DIR / "ROADMAP.md"
        with roadmap.open("a", encoding="utf-8") as fh:
            fh.write("\n## Phase 3: Конвенция\n\n**Plans:** 7/7\n")
        self._reset()
        text = roadmap.read_text(encoding="utf-8")
        self.assertIn("## Phase 3: Конвенция", text)
        self.assertIn("**Plans:** 7/7", text)
        self.assertNotIn("REQ-2102001", text)

    def test_reset_does_not_touch_state_and_phases(self):
        """STATE.md и .planning/phases/ — не артефакты Карты релизов: reset их не трогает."""
        state = sync.PLANNING_DIR / "STATE.md"
        state.parent.mkdir(parents=True, exist_ok=True)
        state.write_text("status: verifying\n", encoding="utf-8")
        phase = sync.PLANNING_DIR / "phases" / "03-convention" / "03-PLAN.md"
        phase.parent.mkdir(parents=True, exist_ok=True)
        phase.write_text("CONV-STD\n", encoding="utf-8")
        self._reset()
        self.assertEqual("status: verifying\n", state.read_text(encoding="utf-8"))
        self.assertEqual("CONV-STD\n", phase.read_text(encoding="utf-8"))

    def test_gate_active_on_first_sync_after_reset(self):
        """После reset база пуста → новый набор заводится только со статусом «Не начато»."""
        self._write_csv([_release_header("v250407"), _task("2102001", status="Релиз")])
        self._sync()
        self._reset()
        self._write_csv([_release_header("v260814"),
                         _task("2830001", status=sync.STATUS_NEW),
                         _task("2830002", status="Готово")])
        data, _ = self._sync()
        self.assertEqual(["2830001"], [t["ID"] for rel in data["releases"] for t in rel["tasks"]])
        self.assertEqual([("2830002", sync.SKIP_NEW_NOT_STARTED)],
                         [(t["ID"], t["Причина пропуска"]) for t in data["skipped"]])

    def test_reset_is_idempotent(self):
        """Повторный reset ничего не ломает и не плодит резервных копий."""
        self._write_csv([_release_header("v250407"), _task("2102001")])
        self._sync()
        self._reset()
        first = (sync.PLANNING_DIR / "ROADMAP.md").read_text(encoding="utf-8")
        self._reset()
        self.assertEqual(first, (sync.PLANNING_DIR / "ROADMAP.md").read_text(encoding="utf-8"))
        self.assertEqual([], list(sync.LEGACY_DIR.glob("*.bak")) if sync.LEGACY_DIR.is_dir() else [])


class TestFetchFromDownloads(unittest.TestCase):
    """fetch: экспорт листа берётся из «Загрузок» (%USERPROFILE%\\Downloads)."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)
        self.downloads = tmp / "Downloads"
        self.downloads.mkdir()
        self.repo = tmp / "repo"
        (self.repo / "RELEASE_MAP").mkdir(parents=True)

        self._orig = (sync.ROOT, sync.CSV_PATH)
        sync.ROOT = self.repo
        sync.CSV_PATH = self.repo / "RELEASE_MAP" / sync.CSV_NAME

        self._env = os.environ.get(sync.DOWNLOADS_ENV)
        os.environ[sync.DOWNLOADS_ENV] = str(self.downloads)

    def tearDown(self):
        sync.ROOT, sync.CSV_PATH = self._orig
        if self._env is None:
            os.environ.pop(sync.DOWNLOADS_ENV, None)
        else:
            os.environ[sync.DOWNLOADS_ENV] = self._env
        self._tmp.cleanup()

    @staticmethod
    def _fetch_quietly():
        """cmd_fetch печатает отчёт в stdout — в тестах он лишний шум."""
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = sync.cmd_fetch(None)
        return code, out.getvalue()

    def _download(self, name, text, mtime):
        path = self.downloads / name
        path.write_text(text, encoding="utf-8")
        os.utime(path, (mtime, mtime))
        return path

    def test_export_from_downloads_replaces_repo_csv(self):
        """Экспорт из «Загрузок» копируется в RELEASE_MAP/ поверх старого CSV."""
        sync.CSV_PATH.write_text("старый CSV", encoding="utf-8")
        self._download(sync.CSV_NAME, "свежий CSV", mtime=1_000_000)
        code, _ = self._fetch_quietly()
        self.assertEqual(0, code)
        self.assertEqual("свежий CSV", sync.CSV_PATH.read_text(encoding="utf-8"))

    def test_newest_duplicate_download_wins(self):
        """Из «… (1).csv» и одноимённого берётся самый свежий по времени файл."""
        stem = sync.CSV_NAME[:-len(".csv")]
        self._download(sync.CSV_NAME, "старая выгрузка", mtime=1_000_000)
        self._download("%s (1).csv" % stem, "новая выгрузка", mtime=2_000_000)
        code, _ = self._fetch_quietly()
        self.assertEqual(0, code)
        self.assertEqual("новая выгрузка", sync.CSV_PATH.read_text(encoding="utf-8"))

    def test_foreign_csv_in_downloads_is_ignored(self):
        """Посторонний CSV в «Загрузках» не считается экспортом листа."""
        sync.CSV_PATH.write_text("текущий CSV", encoding="utf-8")
        self._download("Совсем другой файл.csv", "чужие данные", mtime=2_000_000)
        code, out = self._fetch_quietly()
        self.assertEqual(0, code)
        self.assertIn("ВНИМАНИЕ", out)
        self.assertEqual("текущий CSV", sync.CSV_PATH.read_text(encoding="utf-8"))

    def test_no_export_keeps_current_csv_with_warning(self):
        """Нет свежей выгрузки, но CSV в репозитории есть — предупреждение, синк продолжается."""
        sync.CSV_PATH.write_text("текущий CSV", encoding="utf-8")
        code, out = self._fetch_quietly()
        self.assertEqual(0, code)
        self.assertIn("ВНИМАНИЕ", out)
        self.assertEqual("текущий CSV", sync.CSV_PATH.read_text(encoding="utf-8"))

    def test_no_export_and_no_csv_is_fatal(self):
        """Нет ни выгрузки, ни CSV в репозитории — fail-fast."""
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as ctx:
                sync.cmd_fetch(None)
        self.assertEqual(1, ctx.exception.code)

    def test_identical_export_is_not_copied(self):
        """Совпадающий по содержимому экспорт не перезаписывает CSV репозитория."""
        sync.CSV_PATH.write_text("одинаковый CSV", encoding="utf-8")
        before = sync.CSV_PATH.stat().st_mtime_ns
        self._download(sync.CSV_NAME, "одинаковый CSV", mtime=2_000_000)
        code, out = self._fetch_quietly()
        self.assertEqual(0, code)
        self.assertIn("совпадает", out)
        self.assertEqual(before, sync.CSV_PATH.stat().st_mtime_ns)

    def test_downloads_dir_defaults_to_userprofile(self):
        """Без переопределения каталог «Загрузок» — %USERPROFILE%\\Downloads."""
        os.environ.pop(sync.DOWNLOADS_ENV, None)
        expected = Path(os.environ.get("USERPROFILE") or os.path.expanduser("~")) / "Downloads"
        self.assertEqual(expected, sync.downloads_dir())


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
