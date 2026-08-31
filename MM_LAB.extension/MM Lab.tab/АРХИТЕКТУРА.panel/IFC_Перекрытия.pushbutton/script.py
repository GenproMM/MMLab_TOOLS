#! python3
# -*- coding: utf-8 -*-
"""IFC Перекрытия

Классифицирует перекрытия по МССК и заполняет общие параметры
GP_01_КодКлассифМССК и GP_01_ИмяКлассифМССК. Типоразмеры перекрытий
для обработки выбирает пользователь; фундаментные плиты из списка
исключаются.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "IFC\nПерекрытия"
__author__ = "GENPRO LAB"

# Канонический lib-бутстрап (D-15).
import os
import sys

_SCRIPT_DIR = os.path.dirname(__file__)
# pushbutton -> panel -> tab -> MM_LAB.extension
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
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import FloorType
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


COMMAND_NAME = u"IFC Перекрытия"

# GUID общих параметров (ФОП GENPRO, файл «ГП_ФОП2025.txt»).
GP_01_CODE_GUID = u"4df18cfa-3e5b-4e84-aef6-5ac3385a7d4f"    # GP_01_КодКлассифМССК
GP_01_NAME_GUID = u"91f5b762-e361-462c-8611-3d952be1777b"    # GP_01_ИмяКлассифМССК

# Классификация МССК для перекрытий.
FLOOR_CODE = u"ЭЛ 30 10 40"
FLOOR_NAME = u"Перекрытие"


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========


def get_type_display_name(floor_type):
    """Имя типоразмера для списка выбора: SYMBOL_NAME_PARAM, фолбэк — .Name."""
    try:
        parameter = revit_compat.get_parameter(
            floor_type, BuiltInParameter.SYMBOL_NAME_PARAM
        )
        if parameter is not None:
            value = parameter.AsString()
            if value:
                return value
    except Exception:
        pass
    try:
        return floor_type.Name if floor_type.Name else u""
    except Exception:
        return u""


def is_foundation_slab(floor_type):
    """True, если типоразмер — фундаментная плита (её не классифицируем)."""
    try:
        return bool(floor_type.IsFoundationSlab)
    except Exception:
        return False


def set_classification(floor):
    """Пишет код и имя классификации МССК в общие параметры перекрытия.

    Возвращает (True, "") при успехе либо (False, причина).
    """
    p_code = revit_compat.get_shared_parameter(floor, GP_01_CODE_GUID)
    p_name = revit_compat.get_shared_parameter(floor, GP_01_NAME_GUID)

    if p_code is None:
        return False, u"Параметр 'GP_01_КодКлассифМССК' не найден"
    if p_name is None:
        return False, u"Параметр 'GP_01_ИмяКлассифМССК' не найден"
    if p_code.IsReadOnly:
        return False, u"'GP_01_КодКлассифМССК' только для чтения"
    if p_name.IsReadOnly:
        return False, u"'GP_01_ИмяКлассифМССК' только для чтения"

    p_code.Set(FLOOR_CODE)
    p_name.Set(FLOOR_NAME)
    return True, u""


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


def select_floor_types(type_names):
    """WinForms-диалог множественного выбора типоразмеров перекрытий.

    Возвращает список индексов выбранных типоразмеров либо None (отмена).
    """
    form = Form()
    form.Text = u"Выберите типоразмеры перекрытий"
    form.Size = Size(460, 480)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False

    label = Label()
    label.Text = u"Отметьте типоразмеры перекрытий для обработки:"
    label.Dock = DockStyle.Top
    label.Height = 26
    label.Font = DrawFont("Arial", 9)

    checked_list = CheckedListBox()
    checked_list.Dock = DockStyle.Fill
    checked_list.CheckOnClick = True
    for name in type_names:
        checked_list.Items.Add(name)

    panel = Panel()
    panel.Dock = DockStyle.Bottom
    panel.Height = 46

    btn_ok = Button()
    btn_ok.Text = u"Обработать"
    btn_ok.DialogResult = DialogResult.OK
    btn_ok.Size = Size(150, 30)
    btn_ok.Location = Point(70, 8)

    btn_cancel = Button()
    btn_cancel.Text = u"Отмена"
    btn_cancel.DialogResult = DialogResult.Cancel
    btn_cancel.Size = Size(120, 30)
    btn_cancel.Location = Point(245, 8)

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
    return [int(index) for index in checked_list.CheckedIndices]


# ========== ОСНОВНАЯ ФУНКЦИЯ ==========


def main(doc):
    """Точка входа: выбор типоразмеров, транзакционная запись кодов МССК."""
    revit_compat.require_supported_version(COMMAND_NAME)

    output = script.get_output()

    # 1. Типоразмеры перекрытий (без фундаментных плит).
    all_floor_types = (
        FilteredElementCollector(doc).OfClass(FloorType).ToElements()
    )
    floor_types = [ft for ft in all_floor_types if not is_foundation_slab(ft)]
    if not floor_types:
        alert(u"В проекте нет типоразмеров перекрытий!", COMMAND_NAME)
        return

    floor_types = sorted(floor_types, key=get_type_display_name)
    type_names = [get_type_display_name(ft) for ft in floor_types]

    # 2. Выбор типоразмеров.
    selected_indices = select_floor_types(type_names)
    if selected_indices is None:
        return  # отмена
    if not selected_indices:
        alert(u"Вы не выбрали ни одного типоразмера!", COMMAND_NAME)
        return

    selected_type_ids = set(
        revit_compat.element_id_value(floor_types[i].Id) for i in selected_indices
    )

    # 3. Экземпляры перекрытий выбранных типоразмеров.
    all_floors = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Floors)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    floors_to_process = [
        f
        for f in all_floors
        if revit_compat.element_id_value(f.GetTypeId()) in selected_type_ids
    ]
    if not floors_to_process:
        alert(u"В выбранных типоразмерах нет перекрытий!", COMMAND_NAME)
        return

    if not confirm(
        u"Будет обработано {0} перекрытий из {1} типоразмеров.\nПродолжить?".format(
            len(floors_to_process), len(selected_indices)
        )
    ):
        return

    # 4. Обработка.
    output.print_md(u"## Классификация перекрытий")
    output.print_md(u"Всего перекрытий: **{0}**".format(len(floors_to_process)))

    ok = 0
    fail = 0

    transaction = Transaction(doc, u"Заполнение кодов перекрытий")
    transaction.Start()
    try:
        for floor in floors_to_process:
            success, msg = set_classification(floor)
            floor_id = revit_compat.element_id_value(floor.Id)
            if success:
                ok += 1
            else:
                fail += 1
                output.print_md(u"❌ **{0}** — {1}".format(floor_id, msg))
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    output.print_md(
        u"---\nИтого: ✅ {0} | ❌ {1} → {2} ({3})".format(
            ok, fail, FLOOR_NAME, FLOOR_CODE
        )
    )

    alert(
        u"Обработка завершена!\n\n"
        u"Заполнено перекрытий: {0}\n"
        u"Пропущено: {1}".format(ok, fail),
        COMMAND_NAME,
    )


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
