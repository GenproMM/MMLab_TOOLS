#! python3
# -*- coding: utf-8 -*-
"""IFC Стены

Классифицирует стены по МССК и заполняет общие параметры
GP_01_КодКлассифМССК и GP_01_ИмяКлассифМССК. Типоразмеры стен
для обработки выбирает пользователь; витражи из списка исключаются.
Категория стены (фасад, стена, перегородка, отделка) определяется
по имени типоразмера, толщине, функции типоразмера, материалам слоёв,
примыканию к сантехприборам и наличию помещений с двух сторон.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "IFC\nСтены"
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
from Autodesk.Revit.DB import Outline
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB import Wall
from Autodesk.Revit.DB import WallFunction
from Autodesk.Revit.DB import WallKind
from Autodesk.Revit.DB import WallType
from Autodesk.Revit.DB import XYZ
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


COMMAND_NAME = u"IFC Стены"

# GUID общих параметров (ФОП GENPRO, файл «ГП_ФОП2025.txt»).
GP_01_CODE_GUID = u"4df18cfa-3e5b-4e84-aef6-5ac3385a7d4f"    # GP_01_КодКлассифМССК
GP_01_NAME_GUID = u"91f5b762-e361-462c-8611-3d952be1777b"    # GP_01_ИмяКлассифМССК

# Классификация МССК для стен: код и имя.
FACADE_CODE = u"ЭЛ 30 10 60"
FACADE_NAME = u"Фасад"
STRUCTURAL_CODE = u"ЭЛ 30 10 15"
STRUCTURAL_NAME = u"Стена"
PARTITION_CODE = u"ЭЛ 30 10 20"
PARTITION_NAME = u"Перегородка"
FINISH_CODE = u"ЭЛ 30 26"
FINISH_NAME = u"Отделка стен"

# Пороги толщины, мм. 205 вместо 200 — запас на погрешность перевода единиц.
PARTITION_MAX_MM = 205       # ВС_ и материалы перегородок: до 205 мм — перегородка
THIN_WALL_MAX_MM = 150       # тонкая стена — кандидат в перегородки
WET_ZASHIVKA_MAX_MM = 125    # зашивка в мокрой зоне
FINISH_MAX_MM = 100          # отделка стен
CHAIN_ZASHIVKA_MAX_MM = 130  # цепное примыкание к зашивке

# Допуск раздувания BoundingBox при поиске примыканий, внутренние единицы
# Revit (футы): 0.5 фута ≈ 152 мм.
TOUCH_TOLERANCE = 0.5

# Максимум итераций цепного «заражения» категориями по примыканиям.
MAX_CHAIN_ITERATIONS = 10

ZASHIVKA_KEYWORDS = [u"зашив", u"фальш", u"короб", u"инстал", u"экран"]

WET_ROOM_KEYWORDS = [u"с/у", u"санузел", u"ванн", u"туалет", u"душ", u"санитар"]

PARTITION_MATERIAL_KEYWORDS = [
    u"гкл", u"гклв", u"гкло", u"гвл", u"гипсокарт", u"гипсоволокн",
    u"glk", u"gypsum", u"plasterboard", u"гипс",
    u"пгп", u"пазогребн",
    u"газобет", u"пенобет", u"керамзитобет", u"газосилик", u"силикатн",
    u"блок", u"газоблок", u"пеноблок",
    u"кирпич", u"дерев", u"дсп", u"мдф", u"стекл", u"фибролит",
]


# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========


def get_type_display_name(wall_type):
    """Имя типоразмера стены: SYMBOL_NAME_PARAM, фолбэк — .Name."""
    try:
        parameter = revit_compat.get_parameter(
            wall_type, BuiltInParameter.SYMBOL_NAME_PARAM
        )
        if parameter is not None:
            value = parameter.AsString()
            if value:
                return value
    except Exception:
        pass
    try:
        return wall_type.Name if wall_type.Name else u""
    except Exception:
        return u""


def is_curtain_type(wall_type):
    """True, если типоразмер — витраж (его не классифицируем)."""
    try:
        return wall_type.Kind == WallKind.Curtain
    except Exception:
        return False


def is_vs_prefix(type_name_lower):
    """Префикс ВС_ с учётом русской и английской раскладки."""
    return (
        type_name_lower.startswith(u"вс_")
        or type_name_lower.startswith(u"bc_")
        or type_name_lower.startswith(u"вc_")
        or type_name_lower.startswith(u"вs_")
    )


def get_wall_type(doc, wall):
    """Типоразмер стены либо None."""
    try:
        return doc.GetElement(wall.GetTypeId())
    except Exception:
        return None


def get_width_mm(wall_type):
    """Толщина типоразмера в миллиметрах; 0.0, если недоступна."""
    try:
        return revit_compat.convert_from_internal(wall_type.Width, u"mm")
    except Exception:
        return 0.0


def make_outline(bbox, tolerance=0.0):
    """Outline по BoundingBox, раздутый на tolerance по всем осям."""
    return Outline(
        XYZ(
            bbox.Min.X - tolerance,
            bbox.Min.Y - tolerance,
            bbox.Min.Z - tolerance,
        ),
        XYZ(
            bbox.Max.X + tolerance,
            bbox.Max.Y + tolerance,
            bbox.Max.Z + tolerance,
        ),
    )


def is_near_plumbing(wall, plumbing_outlines):
    """True, если стена примыкает к сантехприбору (зашивка / короб)."""
    wall_bb = wall.get_BoundingBox(None)
    if not wall_bb:
        return False

    # Допуск только в плане: границы стены по Z оставляем как есть.
    outline = Outline(
        XYZ(
            wall_bb.Min.X - TOUCH_TOLERANCE,
            wall_bb.Min.Y - TOUCH_TOLERANCE,
            wall_bb.Min.Z,
        ),
        XYZ(
            wall_bb.Max.X + TOUCH_TOLERANCE,
            wall_bb.Max.Y + TOUCH_TOLERANCE,
            wall_bb.Max.Z,
        ),
    )

    for plumbing_outline in plumbing_outlines:
        if outline.Intersects(plumbing_outline, 0):
            return True
    return False


def is_in_wet_room(doc, wall):
    """True, если середина стены попадает в помещение мокрой зоны."""
    wall_bb = wall.get_BoundingBox(None)
    if not wall_bb:
        return False

    mid_point = (wall_bb.Min + wall_bb.Max) * 0.5
    room = doc.GetRoomAtPoint(mid_point)
    if room is None:
        return False

    room_name = u""
    try:
        name_parameter = revit_compat.get_parameter(
            room, BuiltInParameter.ROOM_NAME
        )
        if name_parameter is not None:
            room_name = name_parameter.AsString() or u""
        if not room_name:
            room_name = room.Name or u""
    except Exception:
        room_name = u""

    room_name = room_name.lower()
    return any(keyword in room_name for keyword in WET_ROOM_KEYWORDS)


def walls_touch(wall1, wall2):
    """True, если BoundingBox стен пересекаются с допуском TOUCH_TOLERANCE."""
    bb1 = wall1.get_BoundingBox(None)
    bb2 = wall2.get_BoundingBox(None)
    if not bb1 or not bb2:
        return False

    return make_outline(bb1, TOUCH_TOLERANCE).Intersects(make_outline(bb2), 0)


def is_attached_to_wall(wall, other_walls):
    """True, если стена лежит «на» другой стене (общая площадь в плане)."""
    try:
        wall_bb = wall.get_BoundingBox(None)
        if not wall_bb:
            return False

        wall_id = revit_compat.element_id_value(wall.Id)
        wall_area = abs(
            (wall_bb.Max.X - wall_bb.Min.X) * (wall_bb.Max.Y - wall_bb.Min.Y)
        )

        for other_wall in other_walls:
            if revit_compat.element_id_value(other_wall.Id) == wall_id:
                continue

            other_bb = other_wall.get_BoundingBox(None)
            if not other_bb:
                continue

            other_area = abs(
                (other_bb.Max.X - other_bb.Min.X)
                * (other_bb.Max.Y - other_bb.Min.Y)
            )

            min_x = max(wall_bb.Min.X, other_bb.Min.X)
            max_x = min(wall_bb.Max.X, other_bb.Max.X)
            min_y = max(wall_bb.Min.Y, other_bb.Min.Y)
            max_y = min(wall_bb.Max.Y, other_bb.Max.Y)

            if min_x < max_x and min_y < max_y:
                intersection_area = (max_x - min_x) * (max_y - min_y)

                min_area = min(wall_area, other_area)
                if min_area > 0 and intersection_area / min_area > 0.5:
                    return True

                if (
                    abs(wall_bb.Min.Z - other_bb.Min.Z) < 0.1
                    and abs(wall_bb.Max.Z - other_bb.Max.Z) < 0.1
                ):
                    return True

        return False
    except Exception:
        return False


def get_room_from_wall(doc, wall, side=0):
    """Помещение с одной стороны стены (side=0 — FromRoom, иначе ToRoom)."""
    try:
        for phase in doc.Phases:
            try:
                if side == 0:
                    room = wall.get_FromRoom(phase)
                else:
                    room = wall.get_ToRoom(phase)
                if room is not None:
                    return room
            except Exception:
                continue
    except Exception:
        pass
    return None


def get_rooms_from_wall(doc, wall):
    """Пара помещений с двух сторон стены (любое может быть None)."""
    return (
        get_room_from_wall(doc, wall, 0),
        get_room_from_wall(doc, wall, 1),
    )


def is_exterior_wall(doc, wall):
    """True, если стена наружная (фасад)."""
    try:
        wall_type = get_wall_type(doc, wall)
        if wall_type is not None:
            type_name = get_type_display_name(wall_type).lower()
            if is_vs_prefix(type_name):
                return False
            if type_name.startswith(u"нс_") or u"наруж" in type_name:
                return True

            # Штатный признак Revit: функция типоразмера стены.
            try:
                if wall_type.Function == WallFunction.Exterior:
                    return True
            except Exception:
                pass

        room1, room2 = get_rooms_from_wall(doc, wall)
        if (room1 is None) != (room2 is None):
            return True

        return False
    except Exception:
        return False


def is_partition_material(doc, wall):
    """True, если в слоях типоразмера есть материал перегородки."""
    try:
        wall_type = get_wall_type(doc, wall)
        if wall_type is None:
            return False

        compound_structure = wall_type.GetCompoundStructure()
        if not compound_structure:
            return False

        for layer in compound_structure.GetLayers():
            try:
                material_id = layer.MaterialId
                if material_id is None:
                    continue
                if revit_compat.element_id_value(material_id) == -1:
                    continue

                material = doc.GetElement(material_id)
                if material is None:
                    continue

                material_name = (material.Name or u"").lower()
                for keyword in PARTITION_MATERIAL_KEYWORDS:
                    if keyword in material_name:
                        return True
            except Exception:
                continue

        return False
    except Exception:
        return False


def is_partition_wall(doc, wall):
    """True, если стена — перегородка (до 200 мм включительно)."""
    try:
        wall_type = get_wall_type(doc, wall)
        if wall_type is None:
            return False

        width_mm = get_width_mm(wall_type)
        type_name = get_type_display_name(wall_type).lower()

        # ПРОВЕРКА 0: стены с префиксом ВС_ до 205 мм.
        if is_vs_prefix(type_name) and width_mm <= PARTITION_MAX_MM:
            if u"несущ" not in type_name and u"bearing" not in type_name:
                return True

        # ПРОВЕРКА 1: материалы перегородок (только для стен <= 205 мм).
        if width_mm <= PARTITION_MAX_MM and is_partition_material(doc, wall):
            if is_exterior_wall(doc, wall):
                return False
            return True

        # ПРОВЕРКА 2: тонкие стены (<= 150 мм).
        if width_mm <= THIN_WALL_MAX_MM:
            if is_exterior_wall(doc, wall):
                return False

            if u"несущ" in type_name or u"bearing" in type_name:
                return False

            if type_name.startswith(u"нс_"):
                return False

            # Стена не наружная и не несущая — считаем перегородкой независимо
            # от того, определились ли помещения с двух сторон.
            return True

        return False
    except Exception:
        return False


def is_finish_wall(doc, wall, target_walls):
    """True, если стена — отделка стен (тонкая накладка на другую стену)."""
    try:
        wall_type = get_wall_type(doc, wall)
        if wall_type is None:
            return False

        if get_width_mm(wall_type) > FINISH_MAX_MM:
            return False

        if is_exterior_wall(doc, wall):
            return False

        if is_partition_wall(doc, wall):
            return False

        if is_attached_to_wall(wall, target_walls):
            room1, room2 = get_rooms_from_wall(doc, wall)
            if room1 is None or room2 is None:
                return True

        return False
    except Exception:
        return False


def set_classification(wall, code, name):
    """Пишет код и имя классификации МССК в общие параметры стены.

    Возвращает (True, "") при успехе либо (False, причина).
    """
    p_code = revit_compat.get_shared_parameter(wall, GP_01_CODE_GUID)
    p_name = revit_compat.get_shared_parameter(wall, GP_01_NAME_GUID)

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


def select_wall_types(type_names):
    """WinForms-диалог множественного выбора типоразмеров стен.

    Возвращает список индексов выбранных типоразмеров либо None (отмена).
    """
    form = Form()
    form.Text = u"Выберите типоразмеры стен"
    form.Size = Size(460, 480)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False

    label = Label()
    label.Text = u"Отметьте типоразмеры стен для обработки:"
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


# ========== КЛАССИФИКАЦИЯ ==========


def collect_plumbing_outlines(doc):
    """Outline сантехприборов проекта — для поиска зашивок и коробов."""
    fixtures = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_PlumbingFixtures)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    outlines = []
    for fixture in fixtures:
        try:
            bbox = fixture.get_BoundingBox(None)
        except Exception:
            continue
        if bbox:
            outlines.append(make_outline(bbox))
    return outlines


def classify_walls(doc, walls, all_walls, plumbing_outlines):
    """Раскладывает стены по категориям МССК.

    Возвращает dict «категория -> {id стены: стена}». Стены, не попавшие
    ни в одну категорию, не классифицируются и не обрабатываются.
    """
    groups = {
        u"ns": {},
        u"structural": {},
        u"zashivka": {},
        u"partition": {},
        u"finish": {},
    }

    def is_classified(wall_id):
        return any(wall_id in group for group in groups.values())

    # --- ЭТАП 1: первичный поиск ---
    for wall in walls:
        wall_id = revit_compat.element_id_value(wall.Id)
        wall_type = get_wall_type(doc, wall)
        if wall_type is None:
            continue

        type_name = get_type_display_name(wall_type)
        type_name_lower = type_name.lower()
        width_mm = get_width_mm(wall_type)

        # 1. Наружные стены.
        if type_name.startswith(u"НС_"):
            groups[u"ns"][wall_id] = wall

        # 2. Стены ВС_ толщиной больше 200 мм (205 мм с запасом) -> «Стена».
        elif is_vs_prefix(type_name_lower) and width_mm > PARTITION_MAX_MM:
            groups[u"structural"][wall_id] = wall

        # 3. Зашивки, короба, экраны.
        elif any(kw in type_name_lower for kw in ZASHIVKA_KEYWORDS):
            groups[u"zashivka"][wall_id] = wall
        elif is_near_plumbing(wall, plumbing_outlines):
            groups[u"zashivka"][wall_id] = wall
        elif width_mm <= WET_ZASHIVKA_MAX_MM and is_in_wet_room(doc, wall):
            groups[u"zashivka"][wall_id] = wall

        # 4. Перегородки (включая ВС_ <= 200 мм).
        elif is_partition_wall(doc, wall):
            groups[u"partition"][wall_id] = wall

        # 5. Отделка стен.
        elif is_finish_wall(doc, wall, all_walls):
            groups[u"finish"][wall_id] = wall

    # --- ЭТАП 2: геометрическое цепное «заражение» по примыканиям ---
    chain_rules = (
        (u"zashivka", CHAIN_ZASHIVKA_MAX_MM),
        (u"partition", THIN_WALL_MAX_MM),
        (u"finish", FINISH_MAX_MM),
    )

    added_new = True
    iteration = 0
    while added_new and iteration < MAX_CHAIN_ITERATIONS:
        added_new = False
        iteration += 1

        snapshots = dict(
            (key, list(groups[key].values())) for key, _ in chain_rules
        )

        for wall in walls:
            wall_id = revit_compat.element_id_value(wall.Id)
            if is_classified(wall_id):
                continue

            wall_type = get_wall_type(doc, wall)
            if wall_type is None:
                continue
            width_mm = get_width_mm(wall_type)

            for key, max_width_mm in chain_rules:
                if width_mm > max_width_mm:
                    continue
                if any(
                    walls_touch(wall, neighbour) for neighbour in snapshots[key]
                ):
                    groups[key][wall_id] = wall
                    added_new = True
                    break

    return groups


# ========== ОСНОВНАЯ ФУНКЦИЯ ==========


def main(doc):
    """Точка входа: выбор типоразмеров, классификация, запись кодов МССК."""
    revit_compat.require_supported_version(COMMAND_NAME)

    output = script.get_output()

    # 1. Типоразмеры стен (без витражей).
    all_wall_types = (
        FilteredElementCollector(doc).OfClass(WallType).ToElements()
    )
    wall_types = [wt for wt in all_wall_types if not is_curtain_type(wt)]
    if not wall_types:
        alert(u"В проекте нет типоразмеров стен!", COMMAND_NAME)
        return

    wall_types = sorted(wall_types, key=get_type_display_name)
    type_names = [get_type_display_name(wt) for wt in wall_types]

    # 2. Выбор типоразмеров.
    selected_indices = select_wall_types(type_names)
    if selected_indices is None:
        return  # отмена
    if not selected_indices:
        alert(u"Вы не выбрали ни одного типоразмера!", COMMAND_NAME)
        return

    selected_type_ids = set(
        revit_compat.element_id_value(wall_types[i].Id) for i in selected_indices
    )

    # 3. Экземпляры стен (без витражей).
    all_wall_instances = (
        FilteredElementCollector(doc)
        .OfClass(Wall)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    all_walls = []
    for wall in all_wall_instances:
        wall_type = get_wall_type(doc, wall)
        if wall_type is None or is_curtain_type(wall_type):
            continue
        all_walls.append(wall)

    walls_to_process = [
        wall
        for wall in all_walls
        if revit_compat.element_id_value(wall.GetTypeId()) in selected_type_ids
    ]
    if not walls_to_process:
        alert(u"В выбранных типоразмерах нет стен!", COMMAND_NAME)
        return

    if not confirm(
        u"Будет проанализировано {0} стен из {1} типоразмеров.\nПродолжить?".format(
            len(walls_to_process), len(selected_indices)
        )
    ):
        return

    # 4. Классификация.
    plumbing_outlines = collect_plumbing_outlines(doc)
    groups = classify_walls(doc, walls_to_process, all_walls, plumbing_outlines)

    classification = (
        (u"ns", FACADE_CODE, FACADE_NAME),
        (u"structural", STRUCTURAL_CODE, STRUCTURAL_NAME),
        (u"zashivka", PARTITION_CODE, PARTITION_NAME),
        (u"partition", PARTITION_CODE, PARTITION_NAME),
        (u"finish", FINISH_CODE, FINISH_NAME),
    )

    classified_count = sum(len(groups[key]) for key, _, _ in classification)
    unclassified_count = len(walls_to_process) - classified_count
    if not classified_count:
        alert(
            u"Ни одна из {0} стен не попала в классификацию МССК.".format(
                len(walls_to_process)
            ),
            COMMAND_NAME,
        )
        return

    # 5. Запись параметров.
    output.print_md(u"## Классификация стен")
    output.print_md(
        u"Проанализировано стен: **{0}**, классифицировано: **{1}**".format(
            len(walls_to_process), classified_count
        )
    )

    ok = 0
    fail = 0

    transaction = Transaction(doc, u"Заполнение кодов стен")
    transaction.Start()
    try:
        for key, code, name in classification:
            for wall_id, wall in groups[key].items():
                success, msg = set_classification(wall, code, name)
                if success:
                    ok += 1
                else:
                    fail += 1
                    output.print_md(u"❌ **{0}** — {1}".format(wall_id, msg))
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise

    output.print_md(
        u"---\n"
        u"- Наружные (НС_): {0} → {1} ({2})\n"
        u"- Стены (ВС_ > 200 мм): {3} → {4} ({5})\n"
        u"- Зашивки / короба: {6} → {7} ({8})\n"
        u"- Перегородки: {9} → {7} ({8})\n"
        u"- Отделка стен: {10} → {11} ({12})".format(
            len(groups[u"ns"]), FACADE_NAME, FACADE_CODE,
            len(groups[u"structural"]), STRUCTURAL_NAME, STRUCTURAL_CODE,
            len(groups[u"zashivka"]), PARTITION_NAME, PARTITION_CODE,
            len(groups[u"partition"]),
            len(groups[u"finish"]), FINISH_NAME, FINISH_CODE,
        )
    )
    output.print_md(
        u"---\nИтого: ✅ {0} | ❌ {1} | без классификации: {2}".format(
            ok, fail, unclassified_count
        )
    )

    alert(
        u"Обработка завершена!\n\n"
        u"Заполнено стен:\n"
        u"• Наружные (НС_): {0}\n"
        u"• Стены (ВС_ > 200 мм): {1}\n"
        u"• Перегородки: {2}\n"
        u"• Зашивки / короба: {3}\n"
        u"• Отделка стен: {4}\n\n"
        u"Всего заполнено: {5}\n"
        u"Пропущено (нет параметров GP_01): {6}\n"
        u"Без классификации: {7}".format(
            len(groups[u"ns"]),
            len(groups[u"structural"]),
            len(groups[u"partition"]),
            len(groups[u"zashivka"]),
            len(groups[u"finish"]),
            ok,
            fail,
            unclassified_count,
        ),
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
