# -*- coding: utf-8 -*-
"""
ВОР Экспорт — выгрузка спецификаций в Excel.
Точка входа pyRevit.
"""

import os
import sys
import traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Собственная папка — insert(0) чтобы ui/ резолвился локально
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

# Валидатор — append чтобы НЕ перекрывать локальный ui/
VALIDATOR_DIR = os.path.normpath(
    os.path.join(SCRIPT_DIR, "..", u"\u0412\u041e\u0420_\u0412\u0430\u043b\u0438\u0434\u0430\u0442\u043e\u0440.pushbutton")
)
if VALIDATOR_DIR not in sys.path:
    sys.path.append(VALIDATOR_DIR)

# Миграция конфига (при первом запуске создаёт config.json)
from core.user_paths import migrate_from_bundle_if_needed
migrate_from_bundle_if_needed()

# Запуск UI
from ui.export_window import ExportMainWindow

try:
    main = ExportMainWindow()
except Exception:
    from pyrevit import script
    logger = script.get_logger()
    logger.error(traceback.format_exc())
    try:
        main.window.Close()
    except Exception:
        pass
