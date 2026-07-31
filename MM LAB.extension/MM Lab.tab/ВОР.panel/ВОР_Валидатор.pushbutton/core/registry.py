# -*- coding: utf-8 -*-
"""
Реестр SCRIPT_ID — центральный файл соответствия ID и скриптов.
Файл scripts/script_registry.json в каталоге bundle (сетевой диск).
"""

import os
import json
import codecs
from pyrevit import script

logger = script.get_logger()

REGISTRY_FILENAME = "script_registry.json"


def _get_registry_path():
    """Путь к файлу реестра в bundle/scripts/."""
    bundle_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(bundle_path, "scripts", REGISTRY_FILENAME)


def load_registry():
    """Загрузить реестр. Возвращает {"_meta": {...}, "ids": {}}."""
    path = _get_registry_path()
    if not os.path.exists(path):
        return {"_meta": {"version": 1}, "ids": {}}
    try:
        with codecs.open(path, "r", "utf-8") as f:
            return json.load(f)
    except Exception:
        return {"_meta": {"version": 1}, "ids": {}}


def save_registry(registry):
    """Сохранить реестр."""
    path = _get_registry_path()
    with codecs.open(path, "w", "utf-8") as f:
        json.dump(registry, f, indent=2, ensure_ascii=False)


def _resolve_registry_path(rel_path):
    """Resolve a path from the registry (may be relative to bundle) to absolute."""
    if os.path.isabs(rel_path):
        return os.path.normcase(os.path.abspath(rel_path))
    bundle_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.normcase(os.path.abspath(os.path.join(bundle_path, rel_path)))


def register_script(script_id, name, abs_path, display_name=""):
    """
    Зарегистрировать SCRIPT_ID в реестре.
    Возвращает (success: bool, message: str).

    Если ID уже занят тем же файлом — обновляет данные.
    Если ID занят другим файлом — отказ.
    """
    registry = load_registry()
    norm_new = os.path.normcase(os.path.abspath(abs_path))

    if script_id in registry.get("ids", {}):
        existing = registry["ids"][script_id]
        existing_path = existing.get("path", "")
        norm_existing = _resolve_registry_path(existing_path)

        if norm_existing == norm_new:
            # Тот же файл — обновляем данные
            existing["name"] = name
            existing["path"] = abs_path
            if display_name:
                existing["display_name"] = display_name
            save_registry(registry)
            return (True, "Registry updated for existing script")
        else:
            return (False, u"ID '{}' already registered for:\n{}".format(
                script_id, existing_path
            ))

    # Новый ID
    if "ids" not in registry:
        registry["ids"] = {}
    registry["ids"][script_id] = {
        "name": name,
        "path": abs_path,
        "display_name": display_name or name
    }
    save_registry(registry)
    logger.info(u"Script registered: {} -> {}".format(script_id, name))
    return (True, "Script registered")


def lookup_by_id(script_id):
    """
    Найти скрипт по ID в реестре.
    Возвращает dict {"name", "abs_path", "display_name"} или None.
    """
    registry = load_registry()
    entry = registry.get("ids", {}).get(script_id)
    if not entry:
        return None
    return {
        "name": entry.get("name", ""),
        "abs_path": entry.get("path", ""),
        "display_name": entry.get("display_name", entry.get("name", ""))
    }


def is_id_registered(script_id):
    """Проверить, занят ли ID в реестре."""
    registry = load_registry()
    return script_id in registry.get("ids", {})


def remove_id(script_id):
    """Удалить ID из реестра."""
    registry = load_registry()
    if script_id in registry.get("ids", {}):
        del registry["ids"][script_id]
        save_registry(registry)
        logger.info(u"Script ID removed from registry: {}".format(script_id))
