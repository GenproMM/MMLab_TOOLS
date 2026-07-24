# -*- coding: utf-8 -*-
"""Тесты структурных правил чекера конвенции MM LAB (план 03-01).

Исполняемая спецификация tools/check_convention.py: правила MM000–MM007,
MM013, режим сырого скрипта, exit-коды CLI, --json, --strict, --baseline
и --write-baseline.

Запуск:
    py -3 -m unittest discover -s tools/tests -p "test_check_convention*.py" -q

Фикстуры (кириллица и пробелы в путях — нарочно, чекер обязан их переваривать):
    fixtures/repo_ok/  — эталонная кнопка, все правила проходят;
    fixtures/repo_bad/ — кнопка-нарушитель (BOM, нет шапки/README/bundle.yaml,
                         не в layout) + орфаны layout «Призрак» и «Нет папки».

Мусор для MM013 (__pycache__/, *.pyc, *.csv) в фикстуры НЕ коммитится
(.gitignore не даст) — тесты создают его во временных копиях кнопки.
"""

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
    REPO_OK / "MM LAB.extension" / "MM Lab.tab"
    / "Тестовая панель.panel" / "Хорошая кнопка.pushbutton"
)
BAD_BUTTON = (
    REPO_BAD / "MM LAB.extension" / "MM Lab.tab"
    / "Плохая панель.panel" / "Плохая кнопка.pushbutton"
)
TAB_BUNDLE_REL = "MM LAB.extension/MM Lab.tab/bundle.yaml"
PANEL_BUNDLE_REL = "MM LAB.extension/MM Lab.tab/Плохая панель.panel/bundle.yaml"


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

    with_panel=True  -> tmp/MM LAB.extension/MM Lab.tab/П.panel/Хорошая кнопка.pushbutton
    with_panel=False -> tmp/Хорошая кнопка.pushbutton (вне *.panel)

    Возвращает (root, button_dir).
    """
    tmp = tempfile.TemporaryDirectory()
    test_case.addCleanup(tmp.cleanup)
    root = Path(tmp.name)
    if with_panel:
        dest = (root / "MM LAB.extension" / "MM Lab.tab"
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


if __name__ == "__main__":
    unittest.main()
