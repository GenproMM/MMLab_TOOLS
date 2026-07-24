#! python3
# -*- coding: utf-8 -*-
"""IFC Окна

Классифицирует окна по МССК и заполняет общие параметры
GP_01_КодКлассифМССК и GP_01_ИмяКлассифМССК. Балконные блоки
распознаются отдельно (по имени семейства/типа и параметру «Модель»)
и получают собственный код; все остальные элементы — код окна.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "IFC\nОкна"
__author__ = "GENPRO LAB"

# Канонический lib-бутстрап (D-15).
import os
import sys

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM LAB.extension
_EXTENSION_DIR = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", ".."))
_LIB_DIR = os.path.join(_EXTENSION_DIR, "lib")
if os.path.isdir(_LIB_DIR) and _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import BuiltInParameter
from Autodesk.Revit.DB import FamilyInstance
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI import TaskDialogCommonButtons
from Autodesk.Revit.UI import TaskDialogResult

from System.Windows.Forms import (
    Button,
    CheckedListBox,
    DialogResult,
    DockStyle,
    Form,
    FormBorderStyle,
    FormStartPosition,
    Label,
    Panel,
)
from System.Drawing import Font as DrawFont
from System.Drawing import Point
from System.Drawing import Size

from pyrevit import script

import revit_compat
from revit_ui_helpers import alert


COMMAND_NAME = u"IFC Окна"

# GUID общих параметров (ФОП GENPRO, файл «ГП_ФОП2025.txt»).
GP_01_CODE_GUID = u"4df18cfa-3e5b-4e84-aef6-5ac3385a7d4f"    # GP_01_КодКлассифМССК
GP_01_NAME_GUID = u"91f5b762-e361-462c-8611-3d952be1777b"    # GP_01_ИмяКлассифМССК


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========


def get_family_name_safe(element, doc):
    """Безопасно получает имя семейства элемента."""
    try:
        if element is None:
            return None
        elem_type = doc.GetElement(element.GetTypeId())
        if elem_type:
            return elem_type.FamilyName
    except Exception:
        pass
    return None


def get_type_name_safe(element, doc):
    """Безопасно получает имя типа элемента."""
    try:
        if element is None:
            return None
        elem_type = doc.GetElement(element.GetTypeId())
        if elem_type:
            return elem_type.Name
    except Exception:
        pass
    return None


def _model_has_balcony(element):
    """True, если параметр «Модель» элемента содержит слово «балкон»."""
    try:
        param_model = revit_compat.get_parameter(
            element, BuiltInParameter.ALL_MODEL_MODEL, u"Модель"
        )
        if param_model and param_model.HasValue:
            val = param_model.AsString()
            if val and u"балкон" in val.lower():
                return True
    except Exception:
        pass
    return False


def is_balcony_block(element, doc):
    """Проверяет, является ли элемент балконным блоком:
    - имя семейства содержит «бблок»;
    - параметр «Модель» (экземпляра или типа) содержит «балкон»;
    - имя типа содержит «балкон».
    """
    try:
        if element is None:
            return False

        # Проверка по имени семейства
        family_name = get_family_name_safe(element, doc)
        if family_name and u"бблок" in family_name.lower():
            return True

        # Проверка по параметру «Модель» (экземпляра)
        if _model_has_balcony(element):
            return True

        # Проверка по параметру «Модель» (типа)
        try:
            elem_type = doc.GetElement(element.GetTypeId())
            if elem_type and _model_has_balcony(elem_type):
                return True
        except Exception:
            pass

        # Проверка по имени типа
        type_name = get_type_name_safe(element, doc)
        if type_name and u"балкон" in type_name.lower():
            return True
    except Exception:
        pass
    return False


def _set_codes(window, code, name):
    """Пишет код и имя классификации в общие параметры окна.

    Возвращает (True, "") при успехе либо (False, причина).
    """
    p_code = revit_compat.get_shared_parameter(window, GP_01_CODE_GUID)
    p_name = revit_compat.get_shared_parameter(window, GP_01_NAME_GUID)

    if p_code is None:
        return False, u"Параметр 'GP_01_КодКлассифМССК' не найден"
    if p_name is None:
        return False, u"Параметр 'GP_01_ИмяКлассифМССК' не найден"
    if p_code.IsReadOnly:
        return False, u"'GP_01_КодКлассифМССК' только для чтения"
    if p_name.IsReadOnly:
        return False, u"'GP_01_ИмяКлассифМССК' только для чтения"

    p_code.Set(code)
    p_name.Set(name)
    return True, u""


def process_window(window, doc):
    """Обрабатывает одно окно (с проверкой на балконный блок)."""
    try:
        # ===== ПРОВЕРКА НА БАЛКОННЫЙ БЛОК =====
        if is_balcony_block(window, doc):
            code = u"ЭЛ 30 18 09"
            name = u"Балконный блок"
        else:
            code = u"ЭЛ 30 18 40"
            name = u"Окно"

        success, msg = _set_codes(window, code, name)
        if not success:
            return False, msg

        family_name = get_family_name_safe(window, doc) or u"Неизвестно"
        type_name = get_type_name_safe(window, doc) or u"Неизвестно"
        return True, u"{0} ({1}) -> {2} [{3}]".format(
            family_name, type_name, name, code
        )
    except Exception as ex:
        return False, u"Ошибка: {0}".format(ex)


# ========== ДИАЛОГИ ==========


def confirm(message):
    """Диалог Да/Нет; возвращает True, если пользователь подтвердил."""
    dialog = TaskDialog(COMMAND_NAME)
    dialog.MainContent = message
    dialog.CommonButtons = (
        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    )
    dialog.DefaultButton = TaskDialogResult.Yes
    return dialog.Show() == TaskDialogResult.Yes


def select_families(family_names, total_windows):
    """WinForms-диалог множественного выбора семейств окон.

    Возвращает список выбранных имён семейств либо None (отмена).
    """
    form = Form()
    form.Text = u"Выбор семейств окон (всего {0} окон)".format(total_windows)
    form.Size = Size(420, 480)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False

    label = Label()
    label.Text = u"Отметьте семейства окон для обработки:"
    label.Dock = DockStyle.Top
    label.Height = 26
    label.Font = DrawFont("Arial", 9)

    checked_list = CheckedListBox()
    checked_list.Dock = DockStyle.Fill
    checked_list.CheckOnClick = True
    for name in family_names:
        checked_list.Items.Add(name)

    panel = Panel()
    panel.Dock = DockStyle.Bottom
    panel.Height = 46

    btn_ok = Button()
    btn_ok.Text = u"Обработать"
    btn_ok.DialogResult = DialogResult.OK
    btn_ok.Size = Size(150, 30)
    btn_ok.Location = Point(60, 8)

    btn_cancel = Button()
    btn_cancel.Text = u"Отмена"
    btn_cancel.DialogResult = DialogResult.Cancel
    btn_cancel.Size = Size(120, 30)
    btn_cancel.Location = Point(225, 8)

    panel.Controls.Add(btn_ok)
    panel.Controls.Add(btn_cancel)

    # Порядок добавления: Top и Bottom заявляют края, Fill добавляем последним.
    form.Controls.Add(label)
    form.Controls.Add(panel)
    form.Controls.Add(checked_list)
    form.AcceptButton = btn_ok
    form.CancelButton = btn_cancel

    if form.ShowDialog() != DialogResult.OK:
        return None
    return [str(item) for item in checked_list.CheckedItems]


# ========== ОСНОВНАЯ ФУНКЦИЯ ==========


def main(doc):
    """Точка входа: сбор окон, выбор семейств, транзакционная запись кодов."""
    revit_compat.require_supported_version(COMMAND_NAME)

    output = script.get_output()

    # 1. Все экземпляры окон
    all_windows = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Windows)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    windows = [w for w in all_windows if isinstance(w, FamilyInstance)]
    if not windows:
        alert(u"В проекте нет экземпляров окон!", COMMAND_NAME)
        return

    # 2. Группировка по семействам
    families_dict = {}
    for window in windows:
        family_name = get_family_name_safe(window, doc)
        if family_name:
            families_dict.setdefault(family_name, []).append(window)

    if not families_dict:
        alert(u"Не удалось определить семейства!", COMMAND_NAME)
        return

    # 3. Выбор семейств
    family_names = sorted(families_dict.keys())
    total_windows = sum(len(items) for items in families_dict.values())

    selected_families = select_families(family_names, total_windows)
    if selected_families is None:
        return  # отмена
    if not selected_families:
        alert(u"Вы не выбрали ни одного семейства!", COMMAND_NAME)
        return

    windows_to_process = []
    for family_name in selected_families:
        windows_to_process.extend(families_dict[family_name])

    if not windows_to_process:
        alert(u"В выбранных семействах нет окон!", COMMAND_NAME)
        return

    if not confirm(
        u"Будет обработано {0} окон из {1} семейств.\nПродолжить?".format(
            len(windows_to_process), len(selected_families)
        )
    ):
        return

    # 4. Обработка
    output.print_md(u"## Обработка окон")
    output.print_md(u"Всего окон: **{0}**".format(len(windows_to_process)))

    ok = 0
    fail = 0
    balcony_processed = 0
    window_processed = 0

    transaction = Transaction(doc, u"Заполнение кодов окон")
    transaction.Start()
    try:
        for window in windows_to_process:
            success, msg = process_window(window, doc)
            window_id = revit_compat.element_id_value(window.Id)
            if success:
                ok += 1
                if u"Балконный блок" in msg:
                    balcony_processed += 1
                else:
                    window_processed += 1
                output.print_md(u"✅ **{0}** — {1}".format(window_id, msg))
            else:
                fail += 1
                output.print_md(u"❌ **{0}** — {1}".format(window_id, msg))
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    output.print_md(
        u"---\nИтого: ✅ {0} | ❌ {1} (балконных блоков: {2}, окон: {3})".format(
            ok, fail, balcony_processed, window_processed
        )
    )

    # 5. Финальный диалог: закрыть ли окно результатов
    if confirm(
        u"ГОТОВО!\n\n"
        u"Обработано: {0}\n"
        u"Пропущено: {1}\n"
        u"Балконных блоков: {2}\n"
        u"Окон: {3}\n\n"
        u"Закрыть окно с результатами?".format(
            ok, fail, balcony_processed, window_processed
        )
    ):
        output.close()


def _entry():
    """Готовит doc/uidoc и вызывает main (doc/uidoc — параметрами, правило 18)."""
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None or uidoc.Document is None:
        TaskDialog.Show(COMMAND_NAME, u"Открой проект Revit и повтори команду.")
        return
    main(uidoc.Document)


try:
    _entry()
except SystemExit:
    pass  # require_supported_version уже показал свой диалог
except Exception as ex:
    TaskDialog.Show(COMMAND_NAME, u"Ошибка:\n{0}".format(ex))
