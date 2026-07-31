# -*- coding: utf-8 -*-
"""
Менеджер конфигураций ВОР Валидатора.
Работает с секцией rp_configs единого config.json.
Настройки скриптов хранятся в отдельных файлах config_<name>.json.
"""

import os
import json
import codecs
from pyrevit import script

logger = script.get_logger()

from core.user_paths import get_config_file, get_script_config_file


def _load_full():
    """Загрузить весь единый конфиг."""
    cf = get_config_file()
    if os.path.exists(cf):
        try:
            with codecs.open(cf, "r", "utf-8") as f:
                return json.load(f)
        except:
            pass
    return {"sections": [], "projects": [], "rp_configs": {}}


def _save_full(full):
    """Сохранить весь единый конфиг."""
    with codecs.open(get_config_file(), "w", "utf-8") as f:
        json.dump(full, f, indent=2, ensure_ascii=False)


def _load_all_settings():
    """Загрузить секцию rp_configs."""
    full = _load_full()
    return full.get("rp_configs", {})


def _save_all_settings(rp_data):
    """Сохранить секцию rp_configs, не трогая остальное."""
    full = _load_full()
    full["rp_configs"] = rp_data
    _save_full(full)


def _make_key(section, project):
    """Создать ключ для Р+П."""
    return "{}|{}".format(section, project)


def load_rp_config(section, project):
    """
    Загрузить конфигурацию для конкретной комбинации Р+П.
    Настройки скриптов НЕ включаются (они в отдельных файлах).

    Returns:
        dict или None если нет сохранённых настроек
    """
    key = _make_key(section, project)
    all_settings = _load_all_settings()
    config = all_settings.get(key, None)
    if config is None:
        return None
    # Убираем inline-настройки если остались от старого формата
    scripts = config.get("custom_scripts", [])
    for s in scripts:
        s.pop("settings", None)
    return config


def save_rp_config(section, project, data):
    """
    Сохранить конфигурацию для Р+П (без настроек скриптов).

    data = {
        "custom_scripts": [{"name": "my_check", "path": "...", "enabled": true}]
    }
    """
    key = _make_key(section, project)
    # Убираем inline-настройки перед сохранением, сохраняем id
    clean_scripts = []
    for s in data.get("custom_scripts", []):
        clean = {"name": s["name"], "path": s["path"], "enabled": s.get("enabled", True)}
        if s.get("id"):
            clean["id"] = s["id"]
        clean_scripts.append(clean)
    clean_data = {"custom_scripts": clean_scripts}

    all_settings = _load_all_settings()
    all_settings[key] = clean_data
    _save_all_settings(all_settings)
    logger.info("Сохранена конфигурация для {}: {}".format(key, clean_data))


def delete_script_from_rp_config(section, project, script_name):
    """
    Удалить скрипт только из конкретной конфигурации Р+П.

    Returns:
        bool — успех
    """
    key = _make_key(section, project)
    all_settings = _load_all_settings()

    if key not in all_settings:
        return False

    config = all_settings[key]
    scripts = config.get("custom_scripts", [])
    original_count = len(scripts)

    config["custom_scripts"] = [s for s in scripts if s.get("name") != script_name]

    if len(config["custom_scripts"]) < original_count:
        all_settings[key] = config
        _save_all_settings(all_settings)
        logger.info("Скрипт '{}' удалён из конфигурации {}".format(script_name, key))
        return True

    return False


def list_custom_scripts():
    """Список всех уникальных скриптов из всех конфигураций."""
    all_settings = _load_all_settings()
    seen = {}

    for key, config in all_settings.items():
        for s in config.get("custom_scripts", []):
            name = s.get("name")
            if name and name not in seen:
                seen[name] = s

    return list(seen.values())


def get_all_settings():
    """Вернуть секцию rp_configs (для экспорта)."""
    return _load_all_settings()


# ================================================================
# НАСТРОЙКИ СКРИПТОВ — отдельные файлы
# ================================================================

def load_script_settings(script_name, section, project):
    """
    Загрузить настройки скрипта для конкретной комбинации Р+П
    из отдельного файла config_<script_name>.json.

    Returns:
        dict — настройки (пустой dict если файла нет)
    """
    path = get_script_config_file(script_name)
    if not os.path.exists(path):
        return {}
    try:
        with codecs.open(path, "r", "utf-8") as f:
            data = json.load(f)
        key = _make_key(section, project)
        return data.get(key, {})
    except:
        return {}


def save_script_settings(script_name, section, project, settings):
    """
    Сохранить настройки скрипта для конкретной комбинации Р+П
    в отдельный файл config_<script_name>.json.
    """
    path = get_script_config_file(script_name)
    data = {}
    if os.path.exists(path):
        try:
            with codecs.open(path, "r", "utf-8") as f:
                data = json.load(f)
        except:
            data = {}

    key = _make_key(section, project)
    data[key] = settings

    with codecs.open(path, "w", "utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    logger.info("Настройки скрипта '{}' сохранены для {}".format(script_name, key))


def delete_script_settings(script_name, section, project):
    """Удалить настройки скрипта для конкретной комбинации Р+П."""
    path = get_script_config_file(script_name)
    if not os.path.exists(path):
        return
    try:
        with codecs.open(path, "r", "utf-8") as f:
            data = json.load(f)
        key = _make_key(section, project)
        data.pop(key, None)
        if data:
            with codecs.open(path, "w", "utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        else:
            os.remove(path)
    except:
        pass


def import_settings(data, merge=True):
    """
    Импортировать настройки.
    merge=True — импортированные ключи перезаписывают совпадающие, остальные сохраняются.
    merge=False — полная замена.
    """
    if merge:
        current = _load_all_settings()
        current.update(data)
        _save_all_settings(current)
    else:
        _save_all_settings(data)
    logger.info("Настройки импортированы (merge={})".format(merge))


# ================================================================
# МИГРАЦИЯ inline-настроек в отдельные файлы
# ================================================================

def migrate_inline_settings():
    """
    Перенести inline-настройки из config.json в отдельные файлы.
    Вызывается один раз при старте.
    """
    all_settings = _load_all_settings()
    migrated = False

    for rp_key, config in all_settings.items():
        for s in config.get("custom_scripts", []):
            inline_settings = s.get("settings")
            if not inline_settings:
                continue

            script_name = s.get("name")
            if not script_name:
                continue

            # Записать в отдельный файл
            path = get_script_config_file(script_name)
            existing = {}
            if os.path.exists(path):
                try:
                    with codecs.open(path, "r", "utf-8") as f:
                        existing = json.load(f)
                except:
                    existing = {}

            if rp_key not in existing:
                existing[rp_key] = inline_settings
                with codecs.open(path, "w", "utf-8") as f:
                    json.dump(existing, f, indent=2, ensure_ascii=False)
                logger.info("Миграция: настройки '{}' для {} вынесены в отдельный файл".format(
                    script_name, rp_key))

            # Удалить из config.json
            s.pop("settings", None)
            migrated = True

    if migrated:
        _save_all_settings(all_settings)
        logger.info("Миграция inline-настроек завершена")


# ================================================================
# МИГРАЦИЯ И РЕЗОЛВ SCRIPT_ID
# ================================================================

def migrate_script_ids_startup():
    """
    При старте: добавить id из SCRIPT_ID в записи конфига, где его нет.
    Идемпотентная — не трогает записи с id.
    """
    all_settings = _load_all_settings()
    changed = False

    for rp_key, config in all_settings.items():
        for s in config.get("custom_scripts", []):
            if s.get("id"):
                continue
            spath = s.get("path", "")
            if not spath or not os.path.exists(spath):
                continue
            try:
                from core.validation_engine import extract_script_metadata
                meta = extract_script_metadata(spath)
                sid = meta.get("script_id")
                if sid:
                    s["id"] = sid
                    changed = True
                    logger.info(u"Migration: added id='{}' to '{}'".format(sid, s.get("name")))
            except Exception:
                pass

    if changed:
        _save_all_settings(all_settings)
        logger.info("Migration: SCRIPT_IDs added to config entries")


def resolve_script_paths(scripts):
    """
    Резолв путей скриптов: если путь битый, попробовать найти по id в реестре.
    Возвращает (scripts, warnings_list).
    """
    from core.registry import lookup_by_id

    warnings = []
    for s in scripts:
        spath = s.get("path", "")
        sid = s.get("id")

        if spath and os.path.exists(spath):
            continue

        # Путь битый — пробуем найти по id
        if sid:
            resolved = lookup_by_id(sid)
            if resolved and resolved.get("abs_path"):
                new_path = resolved["abs_path"]
                if os.path.exists(new_path):
                    s["path"] = new_path
                    logger.info(u"Resolved script '{}' via id '{}' -> {}".format(
                        s.get("name"), sid, new_path))
                    continue
                else:
                    warnings.append(
                        u"Script '{}' (id={}): registry path also invalid: {}".format(
                            s.get("name"), sid, new_path))
            else:
                warnings.append(
                    u"Script '{}' (id={}): not found in registry".format(
                        s.get("name"), sid))
        else:
            warnings.append(
                u"Script '{}' (no id): file not found: {}".format(
                    s.get("name"), spath))

    return (scripts, warnings)
