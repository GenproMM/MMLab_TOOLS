# -*- coding: utf-8 -*-
"""
Импорт единого конфига ВОР Валидатора.
"""

import os
import json
import codecs
import shutil
from datetime import datetime
from pyrevit import script

logger = script.get_logger()


def _backup_existing_config(config_path):
    """Создать резервную копию конфига с меткой времени."""
    if not os.path.exists(config_path):
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(config_path)
    backup_path = "{}_backup_{}{}".format(base, timestamp, ext)

    try:
        shutil.copy2(config_path, backup_path)
        logger.info("Резервная копия: {}".format(backup_path))
        return backup_path
    except Exception as e:
        logger.error("Ошибка создания резервной копии: {}".format(str(e)))
        return None


def import_config(file_path):
    """
    Импортировать конфиг из JSON файла.
    Старый пользовательский конфиг сохраняется с меткой времени.

    Returns:
        (success: bool, message: str)
    """
    if not os.path.exists(file_path):
        return (False, u"Файл не найден: {}".format(file_path))

    try:
        with codecs.open(file_path, "r", "utf-8") as f:
            data = json.load(f)
    except Exception as e:
        return (False, u"Ошибка чтения файла: {}".format(str(e)))

    has_store = ("sections" in data) or ("projects" in data)
    has_settings = "rp_configs" in data

    if not has_store and not has_settings:
        return (False, u"Файл не содержит данных конфигурации")

    from core.user_paths import get_config_file

    config_path = get_config_file()

    backup_path = _backup_existing_config(config_path)

    full = {
        "sections": data.get("sections", []),
        "projects": data.get("projects", []),
        "rp_configs": data.get("rp_configs", {})
    }

    with codecs.open(config_path, "w", "utf-8") as f:
        json.dump(full, f, indent=2, ensure_ascii=False)

    # Миграция inline-настроек в отдельные файлы
    from core.config_manager import migrate_inline_settings
    migrate_inline_settings()

    msg = u"Конфигурация импортирована"
    if backup_path:
        msg += u"\nСтарый конфиг сохранён: {}".format(os.path.basename(backup_path))

    logger.info(msg)
    return (True, msg)
