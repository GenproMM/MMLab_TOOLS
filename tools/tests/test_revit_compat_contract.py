# -*- coding: utf-8 -*-
"""Контрактный тест публичного API MM_LAB.extension/lib/revit_compat.py (план 03-02).

Модуль revit_compat импортировать НЕЛЬЗЯ: он делает ``import clr``,
доступный только внутри Revit/pyRevit. Поэтому тест читает исходник
и проверяет контракт через ast.parse — без импорта и без исполнения.

Что защищает от дрейфа:
    - шапку файла (#! python3, coding, отсутствие UTF-8 BOM);
    - докстринг модуля (строки совместимости и зависимостей);
    - константу SUPPORTED_VERSIONS == (2020, 2022, 2024) (D-02);
    - полный публичный API из 13 функций (контракт для планов 03-04..03-06);
    - запрет голых except (конвенция MM LAB);
    - компилируемость модуля (py_compile).

Запуск:
    py -3 -m unittest discover -s tools/tests -p "test_revit_compat*.py" -q
"""

import ast
import os
import py_compile
import tempfile
import unittest
from pathlib import Path

# tools/tests -> tools -> корень репозитория
COMPAT_PATH = (
    Path(__file__).resolve().parents[2]
    / "MM_LAB.extension"
    / "lib"
    / "revit_compat.py"
)

# Публичный API compat — фиксированный контракт (13 функций).
PUBLIC_API = [
    "get_revit_version",
    "require_supported_version",
    "get_parameter",
    "get_shared_parameter",
    "element_id_value",
    "make_element_id",
    "convert_from_internal",
    "convert_to_internal",
    "create_floor",
    "to_net_list",
    "enum_from_int",
    "iter_count",
    "ensure_vendor_lib",
]


def _read_bytes():
    return COMPAT_PATH.read_bytes()


def _read_text():
    return _read_bytes().decode("utf-8")


def _parse_tree():
    return ast.parse(_read_text(), filename=str(COMPAT_PATH))


class TestRevitCompatContract(unittest.TestCase):
    """Контракт публичного API revit_compat (ast, без Revit)."""

    def test_header(self):
        raw = _read_bytes()
        self.assertNotEqual(
            raw[:3], b"\xef\xbb\xbf", "файл не должен начинаться с UTF-8 BOM"
        )
        lines = raw.decode("utf-8").splitlines()
        self.assertGreaterEqual(len(lines), 2, "в файле меньше двух строк")
        self.assertEqual(lines[0], "#! python3", "строка 1 — шебанг CPython")
        self.assertEqual(
            lines[1], "# -*- coding: utf-8 -*-", "строка 2 — объявление кодировки"
        )

    def test_docstring(self):
        docstring = ast.get_docstring(_parse_tree())
        self.assertIsNotNone(docstring, "у модуля должен быть докстринг")
        self.assertIn("Совместимость: Revit 2020 / 2022 / 2024", docstring)
        self.assertIn("Зависимости: нет", docstring)

    def test_supported_versions(self):
        value = None
        for node in _parse_tree().body:
            if not isinstance(node, ast.Assign):
                continue
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "SUPPORTED_VERSIONS":
                    value = ast.literal_eval(node.value)
        self.assertEqual(
            value,
            (2020, 2022, 2024),
            "SUPPORTED_VERSIONS должен быть кортежем (2020, 2022, 2024) (D-02)",
        )

    def test_public_api(self):
        top_level_functions = {
            node.name
            for node in _parse_tree().body
            if isinstance(node, ast.FunctionDef)
        }
        missing = [name for name in PUBLIC_API if name not in top_level_functions]
        self.assertEqual(
            missing, [], "в revit_compat отсутствуют функции контракта: %s" % missing
        )

    def test_no_bare_except(self):
        bare_excepts = [
            node.lineno
            for node in ast.walk(_parse_tree())
            if isinstance(node, ast.ExceptHandler) and node.type is None
        ]
        self.assertEqual(
            bare_excepts,
            [],
            "голые except: запрещены конвенцией (строки: %s)" % bare_excepts,
        )

    def test_compiles(self):
        # cfile — во временный каталог: без него py_compile пишет
        # __pycache__/ в живое MM_LAB.extension/lib (а /mm-doctor,
        # запускающий этот тест, обязан быть read-only).
        with tempfile.TemporaryDirectory() as tmp:
            py_compile.compile(
                str(COMPAT_PATH),
                cfile=os.path.join(tmp, "revit_compat.pyc"),
                doraise=True,
            )


if __name__ == "__main__":
    unittest.main()
