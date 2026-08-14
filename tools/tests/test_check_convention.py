# -*- coding: utf-8 -*-
"""Тесты чекера конвенции MM LAB (планы 03-01 и 03-03).

Исполняемая спецификация tools/check_convention.py: структурные правила
MM000–MM007, MM013 (план 03-01), AST-правила MM008–MM012, MM014
(план 03-03), режим сырого скрипта, exit-коды CLI, --json, --strict,
--baseline и --write-baseline.

Запуск:
    py -3 -m unittest discover -s tools/tests -p "test_check_convention*.py" -q

Фикстуры (кириллица и пробелы в путях — нарочно, чекер обязан их переваривать):
    fixtures/repo_ok/  — эталонная кнопка, все правила проходят;
    fixtures/repo_bad/ — кнопка-нарушитель (BOM, нет шапки/README/bundle.yaml,
                         не в layout, сторонний импорт, wildcard,
                         LookupParameter-литерал, голый except, pyrevit.forms,
                         legacy-бутстрап EXTENSION_ROOT) + орфаны layout
                         «Призрак» и «Нет папки».

Мусор для MM013 (__pycache__/, *.pyc, *.csv) в фикстуры НЕ коммитится
(.gitignore не даст) — тесты создают его во временных копиях кнопки.
"""

import ast
import contextlib
import io
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

# tools/ — в sys.path, чтобы импортировать check_convention напрямую.
TOOLS_DIR = Path(__file__).resolve().parent.parent
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import check_convention  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"
REPO_OK = FIXTURES / "repo_ok"
REPO_BAD = FIXTURES / "repo_bad"
GOOD_BUTTON = (
    REPO_OK / "MM_LAB.extension" / "MM Lab.tab"
    / "Тестовая панель.panel" / "Хорошая кнопка.pushbutton"
)
BAD_BUTTON = (
    REPO_BAD / "MM_LAB.extension" / "MM Lab.tab"
    / "Плохая панель.panel" / "Плохая кнопка.pushbutton"
)
TAB_BUNDLE_REL = "MM_LAB.extension/MM Lab.tab/bundle.yaml"
PANEL_BUNDLE_REL = "MM_LAB.extension/MM Lab.tab/Плохая панель.panel/bundle.yaml"


def run_main(args):
    """Запускает check_convention.main(args) и возвращает (exit_code, stdout)."""
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        try:
            code = check_convention.main(list(args))
        except SystemExit as exc:  # argparse и ранние выходы
            code = exc.code if isinstance(exc.code, int) else 2
    return code, buffer.getvalue()


def copy_button_to_tmp(test_case, with_panel=False):
    """Копирует «Хорошую кнопку» во временный корень.

    with_panel=True  -> tmp/MM_LAB.extension/MM Lab.tab/П.panel/Хорошая кнопка.pushbutton
    with_panel=False -> tmp/Хорошая кнопка.pushbutton (вне *.panel)

    Возвращает (root, button_dir).
    """
    tmp = tempfile.TemporaryDirectory()
    test_case.addCleanup(tmp.cleanup)
    root = Path(tmp.name)
    if with_panel:
        dest = (root / "MM_LAB.extension" / "MM Lab.tab"
                / "П.panel" / "Хорошая кнопка.pushbutton")
    else:
        dest = root / "Хорошая кнопка.pushbutton"
    shutil.copytree(GOOD_BUTTON, dest)
    return root, dest


def make_junk(button_dir):
    """Создаёт мусор для MM013 внутри папки кнопки (только во временных копиях)."""
    pycache = button_dir / "__pycache__"
    pycache.mkdir()
    (pycache / "cache.pyc").write_bytes(b"\x00")
    (button_dir / "отчёт.csv").write_text("id;name\n", encoding="utf-8")


def write_tmp_script(test_case, source):
    """Пишет одиночный tmp-скрипт (UTF-8 без BOM) и возвращает его Path."""
    tmp = tempfile.TemporaryDirectory()
    test_case.addCleanup(tmp.cleanup)
    path = Path(tmp.name) / "скрипт.py"
    path.write_text(source, encoding="utf-8")
    return path


class GoodButtonTests(unittest.TestCase):
    """Эталонная кнопка проходит все правила."""

    def test_good_button_clean(self):
        violations = check_convention.check_pushbutton(GOOD_BUTTON, REPO_OK)
        self.assertEqual(violations, [])
        code, _out = run_main([str(GOOD_BUTTON), "--root", str(REPO_OK)])
        self.assertEqual(code, 0)


class BadButtonTests(unittest.TestCase):
    """Каждое структурное правило находит свой код на «Плохой кнопке»."""

    @classmethod
    def setUpClass(cls):
        cls.violations = check_convention.check_pushbutton(BAD_BUTTON, REPO_BAD)

    def by_code(self, code):
        return [v for v in self.violations if v.code == code]

    def assert_error(self, code):
        found = self.by_code(code)
        self.assertTrue(found, "ожидалось нарушение %s" % code)
        for violation in found:
            self.assertEqual(violation.severity, "error")

    def test_mm001_missing_shebang(self):
        self.assert_error("MM001")

    def test_mm002_missing_coding(self):
        self.assert_error("MM002")

    def test_mm003_bom(self):
        self.assert_error("MM003")

    def test_mm003_fixture_guard(self):
        # Защита от редакторов, срезающих BOM: первые 3 байта фикстуры.
        raw = (BAD_BUTTON / "script.py").read_bytes()
        self.assertEqual(raw[:3], b"\xef\xbb\xbf")

    def test_mm004_docstring_warning(self):
        found = self.by_code("MM004")
        self.assertTrue(found, "ожидалось нарушение MM004")
        for violation in found:
            self.assertEqual(violation.severity, "warning")

    def test_mm005_missing_bundle_yaml(self):
        self.assert_error("MM005")

    def test_mm006_missing_readme(self):
        self.assert_error("MM006")

    def test_mm007_button_not_in_layout(self):
        # «Плохая кнопка» отсутствует в layout панельного bundle.yaml.
        self.assert_error("MM007")


class LayoutTests(unittest.TestCase):
    """MM007(б): орфаны layout в tab- и panel-bundle.yaml."""

    @classmethod
    def setUpClass(cls):
        cls.violations = check_convention.check_layouts(REPO_BAD)

    def test_mm007_orphan_button_entry(self):
        # Запись «Нет папки» в panel layout без папки; path — panel bundle.yaml.
        orphans = [
            v for v in self.violations
            if v.code == "MM007" and v.path == PANEL_BUNDLE_REL
        ]
        self.assertEqual(len(orphans), 1)
        self.assertIn("Нет папки", orphans[0].message)

    def test_mm007_orphan_panel_in_tab(self):
        # «Призрак» в tab layout записан с хвостовыми пробелами — парсер
        # обязан стрипать; разделитель «-----» орфаном НЕ считается.
        orphans = [
            v for v in self.violations
            if v.code == "MM007" and v.path == TAB_BUNDLE_REL
        ]
        self.assertEqual(len(orphans), 1)
        self.assertIn("Призрак", orphans[0].message)

    def test_mm007_skipped_outside_panel(self):
        # Кнопка вне родителя *.panel (например, templates/) — MM007 пропускается.
        root, button = copy_button_to_tmp(self, with_panel=False)
        violations = check_convention.check_pushbutton(button, root)
        self.assertNotIn("MM007", {v.code for v in violations})


class FileLevelTests(unittest.TestCase):
    """MM000, MM013 и режим сырого скрипта."""

    def test_mm000_unparseable(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        broken = Path(tmp.name) / "битый.py"
        broken.write_text("def broken(:\n    pass\n", encoding="utf-8")
        violations = check_convention.check_script(broken)  # без исключений
        found = [v for v in violations if v.code == "MM000"]
        self.assertTrue(found)
        for violation in found:
            self.assertEqual(violation.severity, "error")

    def test_mm013_junk_warning(self):
        root, button = copy_button_to_tmp(self, with_panel=False)
        make_junk(button)
        violations = check_convention.check_pushbutton(button, root)
        found = [v for v in violations if v.code == "MM013"]
        self.assertTrue(found, "ожидалось нарушение MM013")
        for violation in found:
            self.assertEqual(violation.severity, "warning")

    def test_raw_script_mode(self):
        # Одиночный .py: только правила уровня файла MM000–MM004.
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        raw_script = Path(tmp.name) / "сырой скрипт.py"
        raw_script.write_text("import os\n", encoding="utf-8")
        violations = check_convention.check_script(raw_script)
        codes = {v.code for v in violations}
        self.assertIn("MM001", codes)
        self.assertFalse(codes & {"MM005", "MM006", "MM007", "MM013"})


class AstRuleTests(unittest.TestCase):
    """AST-правила MM008–MM012 и MM014 (план 03-03)."""

    def check_source(self, source, root=None):
        """check_script на tmp-скрипте; возвращает список Violation."""
        path = write_tmp_script(self, source)
        return check_convention.check_script(path, root)

    def codes_of(self, source, root=None):
        """Множество кодов нарушений check_script на tmp-скрипте."""
        return {v.code for v in self.check_source(source, root)}

    # --- MM008: белый список импортов -----------------------------------

    def test_mm008_third_party_import(self):
        # Прямой контракт check_ast_rules/allowed_import_roots.
        tree = ast.parse("import requests\n")
        roots = check_convention.allowed_import_roots(None)
        found = [v for v in check_convention.check_ast_rules(tree, roots)
                 if v.code == "MM008"]
        self.assertTrue(found, "ожидалось нарушение MM008")
        self.assertEqual(found[0].severity, "error")
        self.assertIn("requests", found[0].message)

    def test_mm008_allows_host_and_stdlib(self):
        source = (
            "import clr\n"
            "import System\n"
            "from Autodesk.Revit.DB import Transaction\n"
            "from pyrevit import script\n"
            "import os\n"
        )
        self.assertNotIn("MM008", self.codes_of(source))

    def test_mm008_allows_first_party_lib(self):
        # tmp-репо: first-party модуль в MM_LAB.extension/lib разрешён.
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        lib_dir = root / "MM_LAB.extension" / "lib"
        lib_dir.mkdir(parents=True)
        (lib_dir / "my_helpers.py").write_text("VALUE = 1\n", encoding="utf-8")
        self.assertIn("my_helpers",
                      check_convention.allowed_import_roots(root))
        script_path = root / "скрипт.py"
        script_path.write_text("import my_helpers\n", encoding="utf-8")
        violations = check_convention.check_script(script_path, root)
        self.assertNotIn("MM008", {v.code for v in violations})

    # --- MM009: wildcard-импорт ------------------------------------------

    def test_mm009_wildcard(self):
        violations = self.check_source("from Autodesk.Revit.DB import *\n")
        codes = {v.code for v in violations}
        self.assertIn("MM009", codes)
        # Autodesk — в белом списке: wildcard не должен давать ещё и MM008.
        self.assertNotIn("MM008", codes)
        for violation in violations:
            if violation.code == "MM009":
                self.assertEqual(violation.severity, "error")

    # --- MM010: LookupParameter со строковым литералом --------------------

    def test_mm010_lookup_literal(self):
        source = (
            "def fill(element):\n"
            "    return element.LookupParameter(\"GP_23_Назначение\")\n"
        )
        found = [v for v in self.check_source(source) if v.code == "MM010"]
        self.assertTrue(found, "ожидалось нарушение MM010")
        self.assertEqual(found[0].severity, "warning")
        self.assertIn("revit_compat", found[0].message)

    # --- MM011: голый except ----------------------------------------------

    def test_mm011_bare_except(self):
        source = (
            "try:\n"
            "    value = 1\n"
            "except:\n"
            "    pass\n"
        )
        found = [v for v in self.check_source(source) if v.code == "MM011"]
        self.assertTrue(found, "ожидалось нарушение MM011")
        self.assertEqual(found[0].severity, "warning")

    # --- MM012: pyrevit.forms под CPython3 ---------------------------------

    def test_mm012_pyrevit_forms(self):
        for source in ("from pyrevit import forms\n",
                       "import pyrevit.forms\n"):
            with self.subTest(source=source.strip()):
                found = [v for v in self.check_source(source)
                         if v.code == "MM012"]
                self.assertTrue(found, "ожидалось нарушение MM012")
                self.assertEqual(found[0].severity, "warning")

    # --- MM014: неканонический lib-бутстрап --------------------------------

    def test_mm014_extension_root_name(self):
        source = (
            "import os\n"
            "EXTENSION_ROOT = os.path.abspath(os.path.dirname(__file__))\n"
        )
        found = [v for v in self.check_source(source) if v.code == "MM014"]
        self.assertTrue(found, "ожидалось нарушение MM014")
        self.assertEqual(found[0].severity, "warning")

    def test_mm014_four_parent_hops(self):
        source = (
            "import os\n"
            "repo_root = os.path.join(\n"
            "    os.path.dirname(__file__), \"..\", \"..\", \"..\", \"..\")\n"
        )
        self.assertIn("MM014", self.codes_of(source))

    def test_mm014_noncanonical_syspath(self):
        append_source = (
            "import os\n"
            "import sys\n"
            "LIB_DIR = os.path.dirname(__file__)\n"
            "sys.path.append(LIB_DIR)\n"
        )
        insert_source = (
            "import os\n"
            "import sys\n"
            "LIB_DIR = os.path.dirname(__file__)\n"
            "sys.path.insert(0, LIB_DIR)\n"
        )
        for source in (append_source, insert_source):
            with self.subTest(call=source.splitlines()[-1]):
                self.assertIn("MM014", self.codes_of(source))

    def test_mm014_canonical_bootstrap_clean(self):
        # Канонический блок D-15 дословно — единственная чистая форма.
        source = (
            "import os\n"
            "import sys\n"
            "\n"
            "_SCRIPT_DIR = os.path.dirname(__file__)\n"
            "# pushbutton -> panel -> tab -> MM_LAB.extension\n"
            "_EXTENSION_DIR = os.path.normpath("
            "os.path.join(_SCRIPT_DIR, \"..\", \"..\", \"..\"))\n"
            "_LIB_DIR = os.path.join(_EXTENSION_DIR, \"lib\")\n"
            "if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:\n"
            "    sys.path.insert(0, _LIB_DIR)\n"
        )
        self.assertNotIn("MM014", self.codes_of(source))

    def test_mm014_on_bad_fixture(self):
        violations = check_convention.check_pushbutton(BAD_BUTTON, REPO_BAD)
        self.assertIn("MM014", {v.code for v in violations})

    # --- интеграция и гарды -------------------------------------------------

    def test_raw_script_ast_rules(self):
        # AST-правила работают в режиме сырого .py (приёмка до скаффолда).
        path = write_tmp_script(self, "import requests\n")
        violations = check_convention.check_script(path)
        self.assertIn("MM008", {v.code for v in violations})

    def test_good_button_still_clean(self):
        violations = check_convention.check_pushbutton(GOOD_BUTTON, REPO_OK)
        self.assertEqual(violations, [])

    def test_bad_fixture_bom_preserved(self):
        raw = (BAD_BUTTON / "script.py").read_bytes()
        self.assertEqual(raw[:3], b"\xef\xbb\xbf")


class CliTests(unittest.TestCase):
    """Exit-коды, --json, --strict, --baseline, --write-baseline."""

    def test_exit_codes(self):
        code_bad, _out = run_main(["--all", "--root", str(REPO_BAD)])
        self.assertEqual(code_bad, 1)
        code_ok, _out = run_main(["--all", "--root", str(REPO_OK)])
        self.assertEqual(code_ok, 0)

    def test_json_output(self):
        code, out = run_main(["--all", "--root", str(REPO_BAD), "--json"])
        self.assertEqual(code, 1)
        data = json.loads(out)  # stdout — ровно один валидный JSON-объект
        for key in ("checked", "errors", "warnings", "violations"):
            self.assertIn(key, data)
        self.assertTrue(data["violations"])
        for violation in data["violations"]:
            for key in ("path", "code", "severity", "line", "message"):
                self.assertIn(key, violation)

    def test_strict_escalates_warnings(self):
        # Кнопка только с warning (tmp-копия хорошей + мусор):
        # без --strict exit 0, с --strict exit 1.
        root, button = copy_button_to_tmp(self, with_panel=False)
        make_junk(button)
        code, _out = run_main([str(button), "--root", str(root)])
        self.assertEqual(code, 0)
        code_strict, _out = run_main([str(button), "--root", str(root), "--strict"])
        self.assertEqual(code_strict, 1)

    def test_baseline_roundtrip(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        baseline_path = Path(tmp.name) / "baseline.json"
        code, _out = run_main(
            ["--all", "--root", str(REPO_BAD),
             "--write-baseline", str(baseline_path)]
        )
        self.assertEqual(code, 0)
        self.assertTrue(baseline_path.is_file())
        code_again, _out = run_main(
            ["--all", "--root", str(REPO_BAD), "--baseline", str(baseline_path)]
        )
        self.assertEqual(code_again, 0)

    def test_strict_ignores_baseline(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        baseline_path = Path(tmp.name) / "baseline.json"
        code, _out = run_main(
            ["--all", "--root", str(REPO_BAD),
             "--write-baseline", str(baseline_path)]
        )
        self.assertEqual(code, 0)
        code_strict, _out = run_main(
            ["--all", "--root", str(REPO_BAD),
             "--baseline", str(baseline_path), "--strict"]
        )
        self.assertEqual(code_strict, 1)

    def test_write_baseline_function(self):
        # Прямой контракт write_baseline/load_baseline/apply_baseline.
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        baseline_path = Path(tmp.name) / "baseline.json"
        violations = check_convention.check_pushbutton(BAD_BUTTON, REPO_BAD)
        self.assertTrue(violations)
        check_convention.write_baseline(violations, baseline_path)
        data = json.loads(baseline_path.read_text(encoding="utf-8"))
        self.assertIn("generated", data)
        self.assertIn("units", data)
        baseline = check_convention.load_baseline(baseline_path)
        remaining = check_convention.apply_baseline(violations, baseline)
        self.assertEqual(remaining, [])


class CheckerRegressionTests(unittest.TestCase):
    """Регрессии фиксов ревью фазы 03 (итерация 1): WR-01, WR-02, WR-03, WR-07."""

    def _tmp_dir(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        return Path(tmp.name)

    def test_malformed_baseline_clean_exit_2(self):
        # WR-01: битый baseline -> чистый exit 2 с русским сообщением
        # в stderr (без traceback), а не необработанное исключение.
        cases = ('{"units": []}', '{"units": {"a": "MM001"}}', "не JSON")
        for payload in cases:
            with self.subTest(payload=payload):
                path = self._tmp_dir() / "bad_baseline.json"
                path.write_text(payload, encoding="utf-8")
                stderr = io.StringIO()
                with contextlib.redirect_stderr(stderr):
                    code, _out = run_main(
                        ["--all", "--root", str(REPO_OK),
                         "--baseline", str(path)]
                    )
                self.assertEqual(code, 2)
                self.assertIn("Ошибка чтения baseline", stderr.getvalue())
                self.assertNotIn("Traceback", stderr.getvalue())

    def test_iter_pushbuttons_skips_junk_and_nested(self):
        # WR-02: *.pushbutton внутри .vs/, __pycache__/ или другой кнопки —
        # артефакты редакторов, кнопками НЕ считаются.
        root, button = copy_button_to_tmp(self, with_panel=True)
        tab_dir = root / "MM_LAB.extension" / "MM Lab.tab"
        (tab_dir / ".vs" / "Артефакт.pushbutton").mkdir(parents=True)
        (tab_dir / "__pycache__" / "Кеш.pushbutton").mkdir(parents=True)
        (button / "Вложенная.pushbutton").mkdir()
        found = check_convention.iter_pushbuttons(root)
        self.assertEqual(found, [button.resolve()])

    def test_sibling_module_ast_rules(self):
        # WR-03: AST-правила ловят сторонний импорт в соседнем helpers.py
        # (сообщение с префиксом имени файла), а правила шапки script.py
        # (MM001/MM002/MM004) на соседний модуль не распространяются.
        root, button = copy_button_to_tmp(self, with_panel=False)
        (button / "helpers.py").write_text("import requests\n",
                                           encoding="utf-8")
        violations = check_convention.check_pushbutton(button, root)
        mm008 = [v for v in violations if v.code == "MM008"]
        self.assertTrue(mm008, "ожидалось нарушение MM008 в helpers.py")
        self.assertIn("helpers.py:", mm008[0].message)
        codes = {v.code for v in violations}
        self.assertFalse(codes & {"MM001", "MM002", "MM004"})

    def test_write_baseline_json_prints_json_object(self):
        # WR-07: --json --write-baseline печатает ровно один JSON-объект
        # статуса в stdout (пустой stdout запрещён контрактом --json).
        baseline_path = self._tmp_dir() / "baseline.json"
        code, out = run_main(
            ["--all", "--root", str(REPO_BAD),
             "--write-baseline", str(baseline_path), "--json"]
        )
        self.assertEqual(code, 0)
        data = json.loads(out)
        self.assertEqual(data["baseline_written"], str(baseline_path))
        self.assertGreater(data["violations"], 0)


class PendingAdoptionTests(unittest.TestCase):
    """Секция pending_adoption baseline — временные допуски ещё не принятых кнопок."""

    def _tmp_path(self, name):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        return Path(tmp.name) / name

    def test_apply_baseline_filters_pending_adoption(self):
        # apply_baseline учитывает обе секции: units и pending_adoption.
        violations = check_convention.check_pushbutton(BAD_BUTTON, REPO_BAD)
        self.assertTrue(violations)
        unit_path = violations[0].path
        baseline = {
            "units": {},
            "pending_adoption": {
                unit_path: sorted({v.code for v in violations}),
            },
        }
        remaining = check_convention.apply_baseline(violations, baseline)
        self.assertEqual(remaining, [])

    def test_load_baseline_validates_pending_adoption(self):
        # Битая секция pending_adoption — ValueError (как и units).
        path = self._tmp_path("baseline.json")
        path.write_text('{"units": {}, "pending_adoption": []}',
                        encoding="utf-8")
        with self.assertRaises(ValueError):
            check_convention.load_baseline(path)

    def test_write_baseline_preserves_pending_adoption(self):
        # --write-baseline сохраняет pending_adoption существующего файла
        # и НЕ переносит его пути в units (допуск не превращается
        # в грандфазеринг).
        path = self._tmp_path("baseline.json")
        violations = check_convention.check_pushbutton(BAD_BUTTON, REPO_BAD)
        self.assertTrue(violations)
        unit_path = violations[0].path
        pending = {unit_path: sorted({v.code for v in violations})}
        path.write_text(
            json.dumps({"generated": "2026-01-01", "note": "x",
                        "units": {}, "pending_adoption": pending},
                       ensure_ascii=False),
            encoding="utf-8",
        )
        check_convention.write_baseline(violations, path)
        data = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(data["pending_adoption"], pending)
        self.assertNotIn(unit_path, data["units"])


if __name__ == "__main__":
    unittest.main()
