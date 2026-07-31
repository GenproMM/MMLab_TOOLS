# -*- coding: utf-8 -*-
"""
Сканирование папки scripts/ и синхронизация реестра SCRIPT_ID.
Запускается при старте плагина до загрузки конфигов.
"""

import os
from pyrevit import script

logger = script.get_logger()


def _discover_script_files(scripts_dir):
    """Найти все .py файлы скриптов в scripts/."""
    results = []

    try:
        items = os.listdir(scripts_dir)
    except Exception:
        return results

    for item in items:
        item_path = os.path.join(scripts_dir, item)

        if os.path.isfile(item_path):
            if item.endswith(".py") and not item.startswith(".") and item != "__init__.py":
                results.append(os.path.abspath(item_path))

        elif os.path.isdir(item_path):
            if item.startswith(".") or item == "lib":
                continue

            primary = os.path.join(item_path, item + ".py")
            if os.path.isfile(primary):
                results.append(os.path.abspath(primary))
                continue

            fallback = os.path.join(item_path, "script.py")
            if os.path.isfile(fallback):
                results.append(os.path.abspath(fallback))
                continue

            # Папка переименована, а .py внутри сохранил старое имя — сканируем все .py
            try:
                for sub in os.listdir(item_path):
                    sub_path = os.path.join(item_path, sub)
                    if os.path.isfile(sub_path) and sub.endswith(".py") \
                            and not sub.startswith(".") and sub != "__init__.py":
                        results.append(os.path.abspath(sub_path))
            except Exception:
                pass

    return results


def sync_registry_from_filesystem():
    """
    Сканирует scripts/, извлекает SCRIPT_ID из каждого файла
    и обновляет пути в реестре. Вызывается один раз при старте.
    """
    bundle_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    scripts_dir = os.path.join(bundle_path, "scripts")

    if not os.path.isdir(scripts_dir):
        return

    script_files = _discover_script_files(scripts_dir)
    if not script_files:
        return

    from core.validation_engine import extract_script_metadata
    from core.registry import load_registry, save_registry, _resolve_registry_path

    discovered = {}
    for py_path in script_files:
        try:
            meta = extract_script_metadata(py_path)
            sid = meta.get("script_id")
            if not sid:
                continue
            display_name = meta.get("name", os.path.splitext(os.path.basename(py_path))[0])
            discovered[sid] = {
                "path": py_path,
                "name": display_name,
                "display_name": display_name,
            }
        except Exception:
            continue

    if not discovered:
        return

    registry = load_registry()
    changed = False

    for sid, info in discovered.items():
        abs_path = info["path"]
        norm_new = os.path.normcase(os.path.abspath(abs_path))

        if sid in registry.get("ids", {}):
            existing = registry["ids"][sid]
            existing_path = existing.get("path", "")
            norm_existing = _resolve_registry_path(existing_path)

            if norm_existing != norm_new:
                existing["path"] = abs_path
                existing["name"] = info["name"]
                existing["display_name"] = info["display_name"]
                changed = True
                logger.info(u"Registry path updated for {}: {} -> {}".format(
                    sid, existing_path, abs_path))
        else:
            if "ids" not in registry:
                registry["ids"] = {}
            registry["ids"][sid] = {
                "path": abs_path,
                "name": info["name"],
                "display_name": info["display_name"],
            }
            changed = True
            logger.info(u"Registry: new script discovered: {} -> {}".format(
                sid, info["name"]))

    if changed:
        save_registry(registry)
