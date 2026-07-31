# -*- coding: utf-8 -*-
"""
Хранилище настроек: разделы и проекты с флагом видимости.
Работает с секциями sections/projects единого config.json.
"""

import os
import json
import codecs
from pyrevit import script

logger = script.get_logger()

from core.user_paths import get_config_file

DEFAULT_SECTIONS = [
    "АР", "КР", "КМ", "КМД", "КЖ", "КД",
    "ОВ", "ВК", "ЭО", "ЭМ", "СС", "ТМ",
    "ПС", "АК", "НВК", "НЭ", "ГП", "АУП",
]


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


def _load_store():
    """Загрузить секции sections/projects."""
    full = _load_full()
    sections = full.get("sections")
    projects = full.get("projects")

    if sections is None and projects is None:
        # Создаём дефолтные разделы
        store = {
            "sections": [{"name": s, "visible": True} for s in DEFAULT_SECTIONS],
            "projects": []
        }
        full["sections"] = store["sections"]
        full["projects"] = store["projects"]
        _save_full(full)
        return store

    return {
        "sections": sections or [],
        "projects": projects or []
    }


def _save_store(store_data):
    """Сохранить sections/projects, не трогая rp_configs."""
    full = _load_full()
    full["sections"] = store_data.get("sections", [])
    full["projects"] = store_data.get("projects", [])
    _save_full(full)


def load_last_selection():
    """Загрузить последние выбранные раздел и проект. Возвращает (section, project)."""
    full = _load_full()
    sel = full.get("last_selection")
    if not sel:
        return (None, None)
    return (sel.get("section"), sel.get("project"))


def save_last_selection(section, project):
    """Сохранить последние выбранные раздел и проект."""
    full = _load_full()
    full["last_selection"] = {"section": section or "", "project": project or ""}
    _save_full(full)


def get_sections(visible_only=True):
    """Получить список разделов."""
    store = _load_store()
    sections = store.get("sections", [])
    if visible_only:
        return [s["name"] for s in sections if s.get("visible", True)]
    return sections


def get_projects(visible_only=True):
    """Получить список проектов."""
    store = _load_store()
    projects = store.get("projects", [])
    if visible_only:
        return [p["name"] for p in projects if p.get("visible", True)]
    return projects


def add_section(name):
    """Добавить раздел. Возвращает False если уже существует."""
    name = name.strip()
    if not name:
        return False

    store = _load_store()
    for s in store["sections"]:
        if s["name"] == name:
            return False

    store["sections"].append({"name": name, "visible": True})
    _save_store(store)
    return True


def add_project(name):
    """Добавить проект. Возвращает False если уже существует."""
    name = name.strip()
    if not name:
        return False

    store = _load_store()
    for p in store["projects"]:
        if p["name"] == name:
            return False

    store["projects"].append({"name": name, "visible": True})
    _save_store(store)
    return True


def toggle_section_visibility(name):
    """Переключить видимость раздела."""
    store = _load_store()
    for s in store["sections"]:
        if s["name"] == name:
            s["visible"] = not s.get("visible", True)
            _save_store(store)
            return s["visible"]
    return None


def toggle_project_visibility(name):
    """Переключить видимость проекта."""
    store = _load_store()
    for p in store["projects"]:
        if p["name"] == name:
            p["visible"] = not p.get("visible", True)
            _save_store(store)
            return p["visible"]
    return None


def remove_section(name):
    """Удалить раздел из хранилища (без удаления конфигурации)."""
    store = _load_store()
    original = len(store["sections"])
    store["sections"] = [s for s in store["sections"] if s["name"] != name]
    if len(store["sections"]) < original:
        _save_store(store)
        return True
    return False


def remove_project(name):
    """Удалить проект из хранилища (без удаления конфигурации)."""
    store = _load_store()
    original = len(store["projects"])
    store["projects"] = [p for p in store["projects"] if p["name"] != name]
    if len(store["projects"]) < original:
        _save_store(store)
        return True
    return False


def get_full_store():
    """Вернуть секции sections/projects (для экспорта)."""
    return _load_store()


def import_store(data, merge=True):
    """
    Импортировать хранилище.
    merge=True — импортированные элементы перезаписывают по имени, остальные сохраняются.
    merge=False — полная замена.
    """
    if merge:
        current = _load_store()
        current_sections = {s["name"]: s for s in current.get("sections", [])}
        for s in data.get("sections", []):
            current_sections[s["name"]] = s
        current_projects = {p["name"]: p for p in current.get("projects", [])}
        for p in data.get("projects", []):
            current_projects[p["name"]] = p
        merged = {
            "sections": list(current_sections.values()),
            "projects": list(current_projects.values())
        }
        _save_store(merged)
    else:
        _save_store(data)
    logger.info("Хранилище импортировано (merge={})".format(merge))
