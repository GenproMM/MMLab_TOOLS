# -*- coding: utf-8 -*-
"""
Движок валидации параметров ВОР.
Запускает пользовательские скрипты.
"""

import os
import sys
import re
import codecs
import time
import traceback
import imp
from pyrevit import revit, script, DB

logger = script.get_logger()

BUNDLE_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _decode_unicode_escapes(s):
    """Decode \\uXXXX escape sequences in a captured regex group value."""
    return re.sub(r'\\u([0-9a-fA-F]{4})', lambda m: unichr(int(m.group(1), 16)), s)


def validate_script_id_format(value):
    """Validate SCRIPT_ID format: vor_ + 8 hex chars."""
    if not value:
        return False
    return bool(re.match(r'^vor_[0-9a-f]{8}$', value))


def generate_script_id():
    """Generate a new unique SCRIPT_ID: vor_ + 8 hex chars."""
    import random
    return "vor_{:08x}".format(random.randint(0, 0xFFFFFFFF))


def extract_script_metadata(script_path):
    """Extract metadata from a script file without executing it."""
    base_name = os.path.splitext(os.path.basename(script_path))[0]
    metadata = {
        "name": base_name,
        "description": "",
        "has_settings": False,
        "script_id": None
    }

    try:
        with codecs.open(script_path, "r", "utf-8") as f:
            content = f.read()

        # Triple-quoted description first (multiline)
        desc_match = re.search(
            r'^SCRIPT_DESCRIPTION\s*=\s*(?:u)?"""(.*?)"""', content,
            re.MULTILINE | re.DOTALL
        )
        if not desc_match:
            desc_match = re.search(
                r"^SCRIPT_DESCRIPTION\s*=\s*(?:u)?['\"](.+?)['\"]", content,
                re.MULTILINE
            )
        if desc_match:
            metadata["description"] = _decode_unicode_escapes(desc_match.group(1).strip())

        # Name override
        name_match = re.search(
            r"^SCRIPT_NAME\s*=\s*(?:u)?['\"](.+?)['\"]", content, re.MULTILINE
        )
        if name_match:
            metadata["name"] = _decode_unicode_escapes(name_match.group(1))

        # Has settings flag
        if re.search(r'^HAS_SETTINGS\s*=\s*True', content, re.MULTILINE):
            metadata["has_settings"] = True

        # Script ID
        id_match = re.search(
            r"^SCRIPT_ID\s*=\s*['\"](vor_[0-9a-f]{8})['\"]", content,
            re.MULTILINE
        )
        if id_match:
            metadata["script_id"] = id_match.group(1)
    except Exception:
        pass

    return metadata


class ValidationResult:
    """Результат одной проверки."""

    def __init__(self, check_name, passed, message, elements=None, skip_summary=False,
                 script_id=None, duration_ms=0, has_results_window=False):
        self.check_name = check_name
        self.passed = passed
        self.message = message
        self.elements = elements or []  # ElementId проблемных элементов
        self.skip_summary = skip_summary  # True = скрипт сам показывает результаты
        self.script_id = script_id          # vor_xxxxxxxx (если известен)
        self.duration_ms = duration_ms      # длительность выполнения в мс
        self.has_results_window = has_results_window  # есть ли у скрипта show_results

    def __str__(self):
        status = "OK" if self.passed else "FAIL"
        return "[{}] {}: {}".format(status, self.check_name, self.message)


def run_custom_script(script_info, section, project):
    """
    Запустить пользовательский скрипт валидации.

    Скрипт должен содержать функцию run(doc, section, project),
    возвращающую ValidationResult.
    
    Args:
        script_info: dict — {'name': str, 'path': str}
    """
    script_name = script_info["name"]
    script_path = script_info["path"]

    if not os.path.exists(script_path):
        return ValidationResult(
            check_name=script_name,
            passed=False,
            message="Скрипт не найден: {}".format(script_path),
            script_id=script_info.get("id")
        )

    try:
        # IronPython-совместимая загрузка модуля
        module_name = "custom_{}".format(os.path.splitext(script_name)[0])
        
        # Удаляем из кэша если был загружен ранее
        if module_name in sys.modules:
            del sys.modules[module_name]
        
        # Получаем директорию скрипта для добавления в sys.path
        script_dir = os.path.dirname(script_path)
        if script_dir not in sys.path:
            sys.path.insert(0, script_dir)
        
        # Добавляем bundle_path для импортов из core
        if BUNDLE_PATH not in sys.path:
            sys.path.insert(0, BUNDLE_PATH)
        
        # Загружаем модуль по абсолютному пути
        base = os.path.splitext(os.path.basename(script_path))[0]
        mod_file, mod_path, mod_desc = imp.find_module(base, [script_dir])
        module = imp.load_module(module_name, mod_file, mod_path, mod_desc)
        if mod_file:
            mod_file.close()

        # Вызов функции run
        if hasattr(module, "run"):
            settings = script_info.get("settings", {})
            has_results_window = hasattr(module, "show_results")
            t0 = time.time()
            try:
                result = module.run(revit.doc, section, project, settings)
            except TypeError:
                result = module.run(revit.doc, section, project)
            duration_ms = int((time.time() - t0) * 1000)

            if isinstance(result, ValidationResult):
                # Дополняем метаданными, если движок их знает лучше
                result.duration_ms = duration_ms or result.duration_ms
                result.has_results_window = has_results_window
                if not result.script_id:
                    script_id = script_info.get("id")
                    if script_id:
                        result.script_id = script_id
                return result
            else:
                return ValidationResult(
                    check_name=script_name,
                    passed=False,
                    message="Скрипт вернул неверный тип. Ожидается ValidationResult.",
                    duration_ms=duration_ms,
                    has_results_window=has_results_window,
                    script_id=script_info.get("id")
                )
        else:
            return ValidationResult(
                check_name=script_name,
                passed=False,
                message="В скрипте отсутствует функция run(doc, section, project)",
                script_id=script_info.get("id")
            )

    except Exception as e:
        tb = traceback.format_exc()
        logger.error("Ошибка выполнения скрипта {}: {}\n{}".format(script_name, e, tb))
        return ValidationResult(
            check_name=script_name,
            passed=False,
            message="Ошибка выполнения: {}".format(str(e)),
            script_id=script_info.get("id")
        )


def run_validation(custom_scripts, section, project):
    """
    Запустить все выбранные скрипты и собрать результаты.

    Args:
        custom_scripts: list[dict] — [{name, path}]
        section: str — выбранный раздел
        project: str — выбранный проект

    Returns:
        list[ValidationResult] — результаты всех проверок
    """
    results = []

    for script_info in custom_scripts:
        logger.info("Запуск скрипта: {}".format(script_info.get("name")))
        result = run_custom_script(script_info, section, project)
        results.append(result)

    return results
