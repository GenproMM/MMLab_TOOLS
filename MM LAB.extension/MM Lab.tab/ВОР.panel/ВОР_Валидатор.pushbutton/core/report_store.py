# -*- coding: utf-8 -*-
"""
Хранилище последнего отчёта прогона ВОР Валидатора.

Хранит ТОЛЬКО последний прогон (перезаписывается при каждом новом).
Путь: %APPDATA%\\pyRevit\\ВОР_Валидатор\\report_last.json

Это позволяет при повторном открытии окна валидатора показать
результаты последнего прогона (память процесса при этом пуста,
т.к. окно каждый раз открывается в новом процессе pyRevit).
"""

import os
import json
import codecs
import datetime

from core.user_paths import get_user_config_dir
from pyrevit import script

logger = script.get_logger()


def get_report_file():
    """Вернуть путь к файлу последнего отчёта."""
    return os.path.join(get_user_config_dir(), "report_last.json")


def _result_to_dict(result):
    """Сериализовать ValidationResult в плоский dict для JSON.

    ElementId сериализуются как int (IntegerValue) — обратное
    восстановление через DB.ElementId(int).
    """
    element_ids = []
    try:
        for eid in (result.elements or []):
            try:
                element_ids.append(int(eid.IntegerValue))
            except Exception:
                pass
    except Exception:
        pass

    if result.passed:
        status = "passed"
    else:
        # Отличаем реальные FAIL от ошибок выполнения по сообщению.
        msg = (result.message or "").lower()
        if msg.startswith("\u043e\u0448\u0438\u0431\u043a\u0430"):  # "ошибка"
            status = "error"
        else:
            status = "failed"

    return {
        "name": result.check_name,
        "script_id": result.script_id,
        "status": status,            # passed | failed | error
        "passed": bool(result.passed),
        "message": result.message or "",
        "element_count": len(element_ids),
        "element_ids": element_ids,
        "duration_ms": int(getattr(result, "duration_ms", 0) or 0),
        "has_results_window": bool(getattr(result, "has_results_window", False)),
    }


def save_report(section, project, results, was_stopped=False, scripts_total=None):
    """
    Сохранить отчёт последнего прогона (перезаписать report_last.json).

    Args:
        section: str — раздел
        project: str — проект
        results: list[ValidationResult] — выполненные проверки
        was_stopped: bool — был ли прогон остановлен пользователем
        scripts_total: int|None — сколько всего скриптов должно было
            выполниться (для отображения «N из M» при останове)
    """
    payload = {
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "section": section,
        "project": project,
        "was_stopped": bool(was_stopped),
        "scripts_total": scripts_total,
        "scripts": [_result_to_dict(r) for r in results],
    }

    report_path = get_report_file()
    try:
        with codecs.open(report_path, "w", "utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        logger.info(u"Отчёт сохранён: {} (проверок: {})".format(
            report_path, len(payload["scripts"])))
    except Exception as e:
        logger.error(u"Не удалось сохранить отчёт: {}".format(e))


def load_report():
    """
    Вернуть dict последнего отчёта или None, если файла нет/битый.

    Структура — как в save_report.
    """
    report_path = get_report_file()
    if not os.path.exists(report_path):
        return None
    try:
        with codecs.open(report_path, "r", "utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.warning(u"Не удалось прочитать отчёт {}: {}".format(report_path, e))
        return None
