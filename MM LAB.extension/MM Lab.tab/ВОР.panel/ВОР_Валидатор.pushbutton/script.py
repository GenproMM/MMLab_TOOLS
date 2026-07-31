# -*- coding: utf-8 -*-
"""
ВОР Валидатор — точка входа.
Главный скрипт pyRevit, запускающий окно плагина.
"""

# Добавляем путь в sys.path
import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Запускаем UI
from core.user_paths import migrate_from_bundle_if_needed
migrate_from_bundle_if_needed()

from core.registry_scan import sync_registry_from_filesystem
sync_registry_from_filesystem()

from core.config_manager import migrate_inline_settings, migrate_script_ids_startup
migrate_inline_settings()
migrate_script_ids_startup()

from ui.main_window import MainWindow

try:
    MainWindow()
except Exception as e:
    from pyrevit import script
    logger = script.get_logger()
    logger.error("Ошибка запуска ВОР Валидатора: {}".format(e))
    import traceback
    logger.error(traceback.format_exc())
