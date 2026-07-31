# -*- coding: utf-8 -*-
"""
Управление разделами и проектами.
Загрузка списков разделов и проектов из хранилища настроек.
"""

from pyrevit import revit, DB
from core.settings_store import get_sections, get_projects as get_stored_projects


def get_available_sections():
    """Получить список видимых разделов."""
    return get_sections(visible_only=True)


def get_available_projects():
    """Получить список видимых проектов из хранилища."""
    return get_stored_projects(visible_only=True)


def add_project_to_list(project_name, storage_file=None):
    """Добавить новый проект в список."""
    from core.settings_store import add_project
    return add_project(project_name)
