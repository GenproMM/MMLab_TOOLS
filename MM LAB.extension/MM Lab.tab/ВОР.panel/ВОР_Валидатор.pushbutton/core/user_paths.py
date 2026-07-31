# -*- coding: utf-8 -*-
"""
Пер-пользовательские пути для хранения конфигураций ВОР Валидатора.
Единый файл config.json содержит разделы, проекты и настройки скриптов.
"""

import os
import json
import codecs
from pyrevit import script

logger = script.get_logger()

PLUGIN_DIR_NAME = u"ВОР_Валидатор"


def get_user_config_dir():
    """Вернуть %APPDATA%\\pyRevit\\ВОР_Валидатор, создав при необходимости."""
    from pyrevit import PYREVIT_APP_DIR
    config_dir = os.path.join(PYREVIT_APP_DIR, PLUGIN_DIR_NAME)
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)
    return config_dir


def get_config_file():
    """Путь к единому файлу конфигурации."""
    return os.path.join(get_user_config_dir(), "config.json")


def get_script_config_file(script_name):
    """Путь к файлу настроек скрипта (config_<name>.json)."""
    safe_name = script_name.replace("/", "_").replace("\\", "_")
    return os.path.join(get_user_config_dir(), "config_{}.json".format(safe_name))


def _find_old_file(user_name, bundle_name):
    """Найти старый файл: сначала per-user, потом в bundle."""
    user_path = os.path.join(get_user_config_dir(), user_name)
    if os.path.exists(user_path):
        return user_path
    bundle_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "configs", bundle_name
    )
    if os.path.exists(bundle_path):
        return bundle_path
    return None


def migrate_from_bundle_if_needed():
    """
    Если config.json не существует — собрать из старых файлов
    (settings_by_rp.json + sections_projects_store.json).
    Возвращает True если миграция произошла.
    """
    config_file = get_config_file()
    if os.path.exists(config_file):
        return False

    old_settings = _find_old_file("settings_by_rp.json", "settings_by_rp.json")
    old_store = _find_old_file("sections_projects_store.json", "sections_projects_store.json")

    if not old_settings and not old_store:
        return False

    unified = {"sections": [], "projects": [], "rp_configs": {}}

    if old_store:
        try:
            with codecs.open(old_store, "r", "utf-8") as f:
                store = json.load(f)
                unified["sections"] = store.get("sections", [])
                unified["projects"] = store.get("projects", [])
        except:
            pass

    if old_settings:
        try:
            with codecs.open(old_settings, "r", "utf-8") as f:
                unified["rp_configs"] = json.load(f)
        except:
            pass

    with codecs.open(config_file, "w", "utf-8") as f:
        json.dump(unified, f, indent=2, ensure_ascii=False)

    logger.info(u"Миграция: создан единый config.json")
    return True
