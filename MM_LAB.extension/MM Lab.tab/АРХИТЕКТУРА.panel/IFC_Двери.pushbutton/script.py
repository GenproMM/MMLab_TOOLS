#! python3
# -*- coding: utf-8 -*-
"""IFC Двери

Классифицирует двери по МССК и заполняет общие параметры
GP_01_КодКлассифМССК и GP_01_ИмяКлассифМССК. Тип двери
(межкомнатная / входная / ворота / люк) определяется по назначению
помещений с обеих сторон (FromRoom/ToRoom, параметр GP_23_Назначение)
и по имени семейства / типа элемента.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "IFC\nДвери"
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


COMMAND_NAME = u"IFC Двери"

# GUID общих параметров (ФОП GENPRO, файл «ГП_ФОП2025.txt»).
GP_01_CODE_GUID = u"4df18cfa-3e5b-4e84-aef6-5ac3385a7d4f"    # GP_01_КодКлассифМССК
GP_01_NAME_GUID = u"91f5b762-e361-462c-8611-3d952be1777b"    # GP_01_ИмяКлассифМССК
GP_23_PURPOSE_GUID = u"0b3dbc34-30a7-4278-b1c5-8ba8819f9db4"  # GP_23_Назначение


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========


def get_room_value(room):
    """Читает GP_23_Назначение у помещения; фолбэк — имя помещения."""
    if room is None:
        return u""
    try:
        p = revit_compat.get_shared_parameter(room, GP_23_PURPOSE_GUID)
        if p and p.HasValue:
            return p.AsString()
    except Exception:
        pass
    try:
        return room.Name if room.Name else u""
    except Exception:
        return u""


def get_room_name(room):
    """Получает имя помещения."""
    if room is None:
        return u""
    try:
        return room.Name if room.Name else u""
    except Exception:
        return u""


def get_room_number(room):
    """Получает номер помещения."""
    if room is None:
        return u""
    try:
        return room.Number if room.Number else u""
    except Exception:
        return u""


def get_room_level(room, doc):
    """Получает имя уровня помещения."""
    if room is None:
        return u""
    try:
        level = doc.GetElement(room.LevelId)
        if level:
            return level.Name
    except Exception:
        pass
    return u""


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


def is_luk_element(door, doc):
    """Проверяет, является ли элемент люком, по имени семейства, типа
    или наличию слова «технич»."""
    try:
        if door is None:
            return False

        # Проверка по имени семейства
        family_name = get_family_name_safe(door, doc)
        if family_name and u"люк" in family_name.lower():
            return True

        # Проверка по имени типа
        type_name = get_type_name_safe(door, doc)
        if type_name:
            if u"люк" in type_name.lower():
                return True
            if u"технич" in type_name.lower():
                return True

        # Проверка по имени элемента
        try:
            elem_name = door.Name
            if elem_name:
                if u"люк" in elem_name.lower():
                    return True
                if u"технич" in elem_name.lower():
                    return True
        except Exception:
            pass

        # Проверка по типу (через элемент типа)
        try:
            elem_type = doc.GetElement(door.GetTypeId())
            if elem_type:
                if elem_type.Name:
                    if u"люк" in elem_type.Name.lower():
                        return True
                    if u"технич" in elem_type.Name.lower():
                        return True
                if elem_type.FamilyName:
                    if u"люк" in elem_type.FamilyName.lower():
                        return True
                    if u"технич" in elem_type.FamilyName.lower():
                        return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def is_gate_element(door, doc):
    """Проверяет, является ли элемент воротами, по имени семейства, типа
    или параметру «Модель»."""
    try:
        if door is None:
            return False

        # Проверка по имени семейства
        family_name = get_family_name_safe(door, doc)
        if family_name and u"ворот" in family_name.lower():
            return True

        # Проверка по имени типа
        type_name = get_type_name_safe(door, doc)
        if type_name and u"ворот" in type_name.lower():
            return True

        # Проверка по имени элемента
        try:
            elem_name = door.Name
            if elem_name and u"ворот" in elem_name.lower():
                return True
        except Exception:
            pass

        # Проверка по параметру «Модель» (экземпляра)
        try:
            param_model = revit_compat.get_parameter(
                door, BuiltInParameter.ALL_MODEL_MODEL, u"Модель"
            )
            if param_model and param_model.HasValue:
                val = param_model.AsString()
                if val and u"ворот" in val.lower():
                    return True
        except Exception:
            pass

        # Проверка по типу: параметр «Модель» и имена типа/семейства
        try:
            elem_type = doc.GetElement(door.GetTypeId())
            if elem_type:
                param_model_type = revit_compat.get_parameter(
                    elem_type, BuiltInParameter.ALL_MODEL_MODEL, u"Модель"
                )
                if param_model_type and param_model_type.HasValue:
                    val = param_model_type.AsString()
                    if val and u"ворот" in val.lower():
                        return True
                if elem_type.Name and u"ворот" in elem_type.Name.lower():
                    return True
                if elem_type.FamilyName and u"ворот" in elem_type.FamilyName.lower():
                    return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def classify(v1, v2):
    """По нормализованным именам помещений возвращает (код, название)."""
    def key(v):
        v = v.lower().strip()
        if u"квартир" in v:
            return u"Квартиры"
        if u"моп" in v:
            return u"МОП"
        if v == u"" or u"наруж" in v or u"улиц" in v:
            return u"Наружу"
        return v

    k1, k2 = key(v1), key(v2)

    # Оба помещения — квартиры -> межкомнатная
    if k1 == u"Квартиры" and k2 == u"Квартиры":
        return u"ЭЛ 30 18 30 20", u"Дверь межкомнатная"

    # Квартира <-> МОП -> входная
    if (k1 == u"Квартиры" and k2 == u"МОП") or (k1 == u"МОП" and k2 == u"Квартиры"):
        return u"ЭЛ 30 18 30 30", u"Дверь входная"

    # Квартира <-> пусто -> входная
    if (k1 == u"Квартиры" and k2 == u"") or (k1 == u"" and k2 == u"Квартиры"):
        return u"ЭЛ 30 18 30 30", u"Дверь входная"

    # Квартира <-> наружу -> входная
    if (k1 == u"Квартиры" and k2 == u"Наружу") or (k1 == u"Наружу" and k2 == u"Квартиры"):
        return u"ЭЛ 30 18 30 30", u"Дверь входная"

    # Остальные случаи
    if k1 == u"" and k2 == u"":
        return u"ЭЛ 30 18 30", u"Дверь"
    if (k1 == u"МОП" and k2 == u"") or (k1 == u"" and k2 == u"МОП"):
        return u"ЭЛ 30 18 30", u"Дверь"
    if k1 == u"МОП" and k2 == u"МОП":
        return u"ЭЛ 30 18 30", u"Дверь"
    return u"ЭЛ 30 18 30", u"Дверь"


def get_rooms_from_door(door, doc):
    """Получает помещения с обеих сторон двери через FromRoom/ToRoom.

    Возвращает (room1, room2) либо (None, None).
    """
    try:
        # Перебираем все фазы проекта
        for phase in doc.Phases:
            try:
                from_room = door.get_FromRoom(phase)
                to_room = door.get_ToRoom(phase)
                if from_room is not None or to_room is not None:
                    return from_room, to_room
            except Exception:
                continue

        # Если через фазы не нашли — пробуем напрямую
        try:
            return door.FromRoom, door.ToRoom
        except Exception:
            pass
    except Exception:
        pass

    return None, None


def process_door(door, doc):
    """Обрабатывает одну дверь (с проверкой на люк и ворота)."""
    try:
        # ===== ВОРОТА =====
        if is_gate_element(door, doc):
            return _write_codes(door, u"ЭЛ 30 18 20", u"Ворота", u"ВОРОТА")

        # ===== ЛЮК =====
        if is_luk_element(door, doc):
            return _write_codes(door, u"ЭЛ 30 18 50", u"Люки", u"ЛЮК")

        # ===== ОБЫЧНАЯ ОБРАБОТКА =====
        room1, room2 = get_rooms_from_door(door, doc)

        if room1 is None and room2 is None:
            code = u"ЭЛ 30 18 30"
            name = u"Дверь"
            v1_display = v2_display = u"--"
            name1_display = name2_display = u"--"
            num1_display = num2_display = u"--"
            level1_display = level2_display = u"--"
        else:
            v1 = name1 = num1 = level1 = u""
            if room1:
                v1 = get_room_value(room1)
                name1 = get_room_name(room1)
                num1 = get_room_number(room1)
                level1 = get_room_level(room1, doc)

            v2 = name2 = num2 = level2 = u""
            if room2:
                v2 = get_room_value(room2)
                name2 = get_room_name(room2)
                num2 = get_room_number(room2)
                level2 = get_room_level(room2, doc)

            if not v1:
                v1 = u""
            if not v2:
                v2 = u""

            v1_display = v1 if v1 else u"--"
            v2_display = v2 if v2 else u"--"
            name1_display = name1 if name1 else u"--"
            name2_display = name2 if name2 else u"--"
            num1_display = num1 if num1 else u"--"
            num2_display = num2 if num2 else u"--"
            level1_display = level1 if level1 else u"--"
            level2_display = level2 if level2 else u"--"

            code, name = classify(v1, v2)

        success, msg = _set_codes(door, code, name)
        if not success:
            return False, msg

        level_info = (
            level1_display
            if level1_display == level2_display
            else u"{0}|{1}".format(level1_display, level2_display)
        )
        rooms_info = u"{0} (№{1}, {2}) | {3} (№{4}, {5})".format(
            name1_display, num1_display, v1_display,
            name2_display, num2_display, v2_display,
        )
        return True, u"{0} | {1} -> {2} [{3}]".format(rooms_info, level_info, name, code)

    except Exception as ex:
        return False, u"Ошибка: {0}".format(ex)


def _set_codes(door, code, name):
    """Пишет код и имя классификации в общие параметры двери.

    Возвращает (True, "") при успехе либо (False, причина).
    """
    p_code = revit_compat.get_shared_parameter(door, GP_01_CODE_GUID)
    p_name = revit_compat.get_shared_parameter(door, GP_01_NAME_GUID)

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


def _write_codes(door, code, name, tag):
    """Пишет коды и формирует сообщение для люка/ворот."""
    success, msg = _set_codes(door, code, name)
    if not success:
        return False, msg
    return True, u"{0} -> {1} ({2})".format(tag, name, code)


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


def select_families(family_names, total_doors):
    """WinForms-диалог множественного выбора семейств дверей.

    Возвращает список выбранных имён семейств либо None (отмена).
    """
    form = Form()
    form.Text = u"Выбор семейств дверей (всего {0} дверей)".format(total_doors)
    form.Size = Size(420, 480)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False

    label = Label()
    label.Text = u"Отметьте семейства дверей для обработки:"
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
    """Точка входа: сбор дверей, выбор семейств, транзакционная запись кодов."""
    revit_compat.require_supported_version(COMMAND_NAME)

    output = script.get_output()

    # 1. Все экземпляры дверей
    all_doors = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Doors)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    doors = [d for d in all_doors if isinstance(d, FamilyInstance)]
    if not doors:
        alert(u"В проекте нет экземпляров дверей!", COMMAND_NAME)
        return

    # 2. Группировка по семействам
    families_dict = {}
    for door in doors:
        family_name = get_family_name_safe(door, doc)
        if family_name:
            families_dict.setdefault(family_name, []).append(door)

    if not families_dict:
        alert(u"Не удалось определить семейства!", COMMAND_NAME)
        return

    # 3. Выбор семейств
    family_names = sorted(families_dict.keys())
    total_doors = sum(len(items) for items in families_dict.values())

    selected_families = select_families(family_names, total_doors)
    if selected_families is None:
        return  # отмена
    if not selected_families:
        alert(u"Вы не выбрали ни одного семейства!", COMMAND_NAME)
        return

    doors_to_process = []
    for family_name in selected_families:
        doors_to_process.extend(families_dict[family_name])

    if not doors_to_process:
        alert(u"В выбранных семействах нет дверей!", COMMAND_NAME)
        return

    if not confirm(
        u"Будет обработано {0} дверей из {1} семейств.\nПродолжить?".format(
            len(doors_to_process), len(selected_families)
        )
    ):
        return

    # 4. Обработка
    output.print_md(u"## Обработка дверей")
    output.print_md(u"Всего дверей: **{0}**".format(len(doors_to_process)))

    ok = 0
    fail = 0
    luk_processed = 0
    gate_processed = 0

    transaction = Transaction(doc, u"Заполнение кодов дверей")
    transaction.Start()
    try:
        for door in doors_to_process:
            success, msg = process_door(door, doc)
            door_id = revit_compat.element_id_value(door.Id)
            if success:
                ok += 1
                if u"ЛЮК" in msg:
                    luk_processed += 1
                elif u"ВОРОТА" in msg:
                    gate_processed += 1
                output.print_md(u"✅ **{0}** — {1}".format(door_id, msg))
            else:
                fail += 1
                output.print_md(u"❌ **{0}** — {1}".format(door_id, msg))
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    output.print_md(
        u"---\nИтого: ✅ {0} | ❌ {1} (люков: {2}, ворот: {3})".format(
            ok, fail, luk_processed, gate_processed
        )
    )

    # 5. Финальный диалог: закрыть ли окно результатов
    if confirm(
        u"ГОТОВО!\n\n"
        u"Обработано: {0}\n"
        u"Пропущено: {1}\n"
        u"Обработано люков: {2}\n"
        u"Обработано ворот: {3}\n\n"
        u"Закрыть окно с результатами?".format(
            ok, fail, luk_processed, gate_processed
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
