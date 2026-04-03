#! python3
# -*- coding: utf-8 -*-
"""Проверка мокрых зон

Проверяет все помещения в проекте на пересечение с проекцией
мокрых помещений с уровня выше. Мокрые помещения определяются
по пользовательскому списку имён/частей имён. Формирует отчёт
с Id, номером и именем помещений, попадающих под мокрые зоны.

Алгоритм определения «уровня выше»:
1. Составляется список уровней, на которых есть хотя бы одно
   размещённое помещение в заданной стадии проекта.
2. Список сортируется по высотной отметке.
3. Для каждого помещения берётся следующий уровень из списка.

Совместимость: Revit 2020 / 2022 / 2024
Зависимости: нет
"""

__title__ = "Мокрые\nзоны"
__author__ = "GENPRO LAB"
__doc__ = (
    "Проверка помещений на пересечение с проекцией "
    "мокрых помещений уровнем выше. Помещения, "
    "попадающие под мокрые зоны, выводятся в отчёт."
)

# === IMPORTS ===
import clr
import sys
import os

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    SpatialElementBoundaryOptions,
    SpatialElementBoundaryLocation,
    BuiltInCategory,
    BuiltInParameter,
    XYZ,
    Line,
    Arc,
    CurveLoop,
    GeometryCreationUtilities,
    BooleanOperationsUtils,
    BooleanOperationsType,
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons

# .NET коллекции для передачи в Revit API
from System.Collections.Generic import List as NetList

# WinForms для диалога настроек
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import (
    Form, Label, TextBox, ComboBox, Button,
    DialogResult, FormStartPosition, FormBorderStyle,
    ComboBoxStyle, DataGridView, DataGridViewAutoSizeColumnsMode,
    DataGridViewSelectionMode, DataGridViewColumnHeadersHeightSizeMode,
    DockStyle,
)
from System.Drawing import Point, Size, Font as DrawFont

# === CONFIGURATION ===
# Значения по умолчанию для фильтра мокрых помещений
DEFAULT_WET_PATTERNS = "санузел, ванная, душ, туалет, постирочная, прачечная"


# === HELPERS ===

def alert(message, title="Мокрые зоны"):
    """Показать сообщение через Revit TaskDialog."""
    td = TaskDialog(title)
    td.MainContent = message
    td.CommonButtons = TaskDialogCommonButtons.Ok
    td.Show()


def get_all_phases(doc):
    """Получить все стадии проекта в виде списка (имя, Phase)."""
    return [(p.Name, p) for p in doc.Phases]


def get_rooms_in_phase(doc, phase):
    """Получить все размещённые помещения в заданной стадии."""
    elements = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    phase_id_int = phase.Id.IntegerValue
    rooms = []
    for room in elements:
        if room is None:
            continue
        try:
            if room.Area <= 0:
                continue
        except:
            continue
        phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
        if phase_param is None:
            continue
        if phase_param.AsElementId().IntegerValue == phase_id_int:
            rooms.append(room)
    return rooms


def get_room_boundary_loop(room, z=0.0):
    """Получить внешний контур помещения, спроецированный на отметку Z.

    Возвращает CurveLoop или None, если контур невозможно построить.
    """
    opts = SpatialElementBoundaryOptions()
    opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
    try:
        boundary = room.GetBoundarySegments(opts)
    except:
        return None

    if boundary is None or boundary.Count == 0:
        return None

    # Берём первый (внешний) контур
    segments = boundary[0]
    if segments is None or segments.Count == 0:
        return None

    curves = []
    for seg in segments:
        curve = seg.GetCurve()
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        new_p0 = XYZ(p0.X, p0.Y, z)
        new_p1 = XYZ(p1.X, p1.Y, z)
        # Пропускаем вырожденные сегменты
        if new_p0.DistanceTo(new_p1) < 1e-9:
            continue
        try:
            # Проверяем тип кривой — Arc или Line
            if type(curve).__name__ == "Arc":
                mid = curve.Evaluate(0.5, True)
                new_mid = XYZ(mid.X, mid.Y, z)
                curves.append(Arc.Create(new_p0, new_p1, new_mid))
            else:
                curves.append(Line.CreateBound(new_p0, new_p1))
        except:
            # .NET исключение при вырожденных кривых — пропускаем
            continue

    if len(curves) < 3:
        return None

    try:
        loop = CurveLoop()
        for c in curves:
            loop.Append(c)
        return loop
    except:
        return None


def make_solid_from_loop(loop):
    """Создать тонкую экструзию из CurveLoop для проверки пересечений."""
    profile = NetList[CurveLoop]()
    profile.Add(loop)
    return GeometryCreationUtilities.CreateExtrusionGeometry(
        profile, XYZ(0, 0, 1), 1.0
    )


def loops_intersect(loop1, loop2):
    """Проверить пересечение двух 2D-контуров.

    Создаёт тонкие экструзии и выполняет Boolean Intersect.
    Пересечение есть, если объём результата > 0.
    """
    try:
        solid1 = make_solid_from_loop(loop1)
        solid2 = make_solid_from_loop(loop2)
        result = BooleanOperationsUtils.ExecuteBooleanOperation(
            solid1, solid2, BooleanOperationsType.Intersect
        )
        return result is not None and result.Volume > 1e-9
    except:
        return False


def is_wet_room(room, patterns):
    """Проверить, является ли помещение мокрым (имя содержит одну из подстрок)."""
    name_param = room.get_Parameter(BuiltInParameter.ROOM_NAME)
    if name_param is None:
        return False
    room_name = (name_param.AsString() or "").strip().lower()
    if not room_name:
        return False
    for pattern in patterns:
        if pattern in room_name:
            return True
    return False


# === UI ===

def show_report(report, phase_name, filter_text, total_checked, uidoc):
    """Показать результат проверки в WinForms-окне с таблицей.

    Клик по ячейке Id элемента — выделяет помещение в модели
    и открывает его на подходящем виде.
    """
    from Autodesk.Revit.DB import ElementId
    from System.Windows.Forms import (
        DataGridViewClipboardCopyMode,
        DataGridViewCellStyle,
        FormWindowState,
        Cursors,
    )
    from System.Drawing import Color, FontStyle

    frm = Form()
    frm.Text = "Отчёт: помещения под мокрыми зонами"
    frm.Size = Size(700, 500)
    frm.StartPosition = FormStartPosition.CenterScreen

    # Заголовок
    lbl = Label()
    lbl.Text = (
        "Стадия: {}  |  Фильтр: {}  |  "
        "Проверено: {}  |  Найдено: {}".format(
            phase_name, filter_text, total_checked, len(report)
        )
    )
    lbl.Dock = DockStyle.Top
    lbl.Font = DrawFont("Arial", 9)
    lbl.Height = 30
    frm.Controls.Add(lbl)

    # Таблица
    dgv = DataGridView()
    dgv.Dock = DockStyle.Fill
    dgv.ReadOnly = True
    dgv.AllowUserToAddRows = False
    dgv.AllowUserToDeleteRows = False
    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    dgv.ColumnHeadersHeightSizeMode = (
        DataGridViewColumnHeadersHeightSizeMode.AutoSize
    )
    dgv.ClipboardCopyMode = (
        DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
    )

    dgv.ColumnCount = 4
    dgv.Columns[0].Name = "№ п/п"
    dgv.Columns[1].Name = "Id элемента"
    dgv.Columns[2].Name = "Номер"
    dgv.Columns[3].Name = "Имя помещения"

    dgv.Columns[0].FillWeight = 10
    dgv.Columns[1].FillWeight = 15
    dgv.Columns[2].FillWeight = 15
    dgv.Columns[3].FillWeight = 60

    # Стиль для столбца Id — синий, подчёркнутый, курсор «рука»
    link_style = DataGridViewCellStyle()
    link_style.ForeColor = Color.Blue
    link_style.Font = DrawFont(
        "Arial", 9, FontStyle.Underline
    )
    dgv.Columns[1].DefaultCellStyle = link_style

    from System import Array, String
    for i, (eid, num, name) in enumerate(report, 1):
        row = Array[String]([str(i), str(eid), str(num), str(name)])
        dgv.Rows.Add(row)

    # Обработчик клика по ячейке Id — выделение и показ элемента
    def on_cell_click(sender, e):
        if e.RowIndex < 0:
            return
        # Столбец 1 = Id элемента
        if e.ColumnIndex != 1:
            return
        cell_value = sender.Rows[e.RowIndex].Cells[1].Value
        if cell_value is None:
            return
        try:
            eid_int = int(cell_value)
            elem_id = ElementId(eid_int)
            ids = NetList[ElementId]()
            ids.Add(elem_id)
            uidoc.Selection.SetElementIds(ids)
            uidoc.ShowElements(elem_id)
            # Сворачиваем окно, чтобы пользователь видел модель
            frm.WindowState = FormWindowState.Minimized
        except:
            pass

    dgv.CellClick += on_cell_click

    # Курсор «рука» при наведении на столбец Id
    def on_cell_mouse_enter(sender, e):
        if e.ColumnIndex == 1 and e.RowIndex >= 0:
            sender.Cursor = Cursors.Hand
        else:
            sender.Cursor = Cursors.Default

    dgv.CellMouseEnter += on_cell_mouse_enter

    frm.Controls.Add(dgv)
    dgv.BringToFront()

    frm.ShowDialog()
    frm.Dispose()


class WetZonesDialog(Form):
    """Диалог настроек проверки мокрых зон."""

    def __init__(self, phases):
        Form.__init__(self)
        self.phases = phases
        self._setup_ui()

    def _setup_ui(self):
        self.Text = "Проверка мокрых зон"
        self.Size = Size(480, 250)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        # Названия мокрых помещений
        lbl_patterns = Label()
        lbl_patterns.Text = "Названия мокрых помещений (через запятую):"
        lbl_patterns.Location = Point(15, 15)
        lbl_patterns.Size = Size(430, 20)
        self.Controls.Add(lbl_patterns)

        self.txt_patterns = TextBox()
        self.txt_patterns.Location = Point(15, 38)
        self.txt_patterns.Size = Size(430, 24)
        self.txt_patterns.Text = DEFAULT_WET_PATTERNS
        self.Controls.Add(self.txt_patterns)

        # Стадия проекта
        lbl_phase = Label()
        lbl_phase.Text = "Стадия проекта:"
        lbl_phase.Location = Point(15, 78)
        lbl_phase.Size = Size(430, 20)
        self.Controls.Add(lbl_phase)

        self.cmb_phase = ComboBox()
        self.cmb_phase.Location = Point(15, 101)
        self.cmb_phase.Size = Size(430, 24)
        self.cmb_phase.DropDownStyle = ComboBoxStyle.DropDownList
        for name, _ in self.phases:
            self.cmb_phase.Items.Add(name)
        if self.cmb_phase.Items.Count > 0:
            self.cmb_phase.SelectedIndex = self.cmb_phase.Items.Count - 1
        self.Controls.Add(self.cmb_phase)

        # Кнопки
        btn_ok = Button()
        btn_ok.Text = "Проверить"
        btn_ok.Location = Point(260, 165)
        btn_ok.Size = Size(90, 30)
        btn_ok.DialogResult = DialogResult.OK
        self.AcceptButton = btn_ok
        self.Controls.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(360, 165)
        btn_cancel.Size = Size(80, 30)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.CancelButton = btn_cancel
        self.Controls.Add(btn_cancel)


# === MAIN ===

def main():
    doc = __revit__.ActiveUIDocument.Document

    # Получаем стадии проекта
    phases = get_all_phases(doc)
    if not phases:
        alert("В проекте не найдено стадий.")
        return

    # Показываем диалог настроек
    dlg = WetZonesDialog(phases)
    if dlg.ShowDialog() != DialogResult.OK:
        dlg.Dispose()
        return

    raw_patterns = dlg.txt_patterns.Text.strip()
    phase_idx = dlg.cmb_phase.SelectedIndex
    dlg.Dispose()

    if not raw_patterns:
        alert("Не указаны названия мокрых помещений.")
        return
    if phase_idx < 0:
        alert("Не выбрана стадия проекта.")
        return

    patterns = [p.strip().lower() for p in raw_patterns.split(",") if p.strip()]
    _, phase = phases[phase_idx]

    # Собираем все помещения в выбранной стадии
    all_rooms = get_rooms_in_phase(doc, phase)
    if not all_rooms:
        alert("В выбранной стадии не найдено размещённых помещений.")
        return

    # Составляем список уровней с помещениями, сортируем по отметке
    level_ids = set()
    for r in all_rooms:
        level_ids.add(r.LevelId.IntegerValue)

    levels_dict = {}
    for lid_int in level_ids:
        from Autodesk.Revit.DB import ElementId
        lvl = doc.GetElement(ElementId(lid_int))
        if lvl is not None:
            levels_dict[lid_int] = lvl

    sorted_levels = sorted(levels_dict.values(), key=lambda l: l.Elevation)
    # Карта: IntegerValue уровня -> индекс в отсортированном списке
    level_order = [l.Id.IntegerValue for l in sorted_levels]
    level_index_map = {lid_int: i for i, lid_int in enumerate(level_order)}

    # Группируем помещения по уровню (по IntegerValue)
    rooms_by_level = {}
    for r in all_rooms:
        lid_int = r.LevelId.IntegerValue
        if lid_int not in rooms_by_level:
            rooms_by_level[lid_int] = []
        rooms_by_level[lid_int].append(r)

    # Определяем мокрые помещения на каждом уровне
    wet_rooms_by_level = {}
    for lid_int, rooms in rooms_by_level.items():
        wet = [r for r in rooms if is_wet_room(r, patterns)]
        if wet:
            wet_rooms_by_level[lid_int] = wet

    # Кэшируем контуры мокрых помещений (проекция на Z=0)
    wet_loops_cache = {}
    for lid_int, wet_rooms in wet_rooms_by_level.items():
        loops = []
        for wr in wet_rooms:
            loop = get_room_boundary_loop(wr, z=0.0)
            if loop is not None:
                loops.append(loop)
        if loops:
            wet_loops_cache[lid_int] = loops

    # Проверяем каждое помещение на пересечение с мокрыми зонами сверху
    # (мокрые помещения исключаются из отчёта)
    report = []
    for room in all_rooms:
        # Само мокрое — пропускаем
        if is_wet_room(room, patterns):
            continue

        room_level_int = room.LevelId.IntegerValue
        idx = level_index_map.get(room_level_int)
        if idx is None:
            continue

        # Нет уровня выше — пропускаем
        if idx >= len(level_order) - 1:
            continue

        upper_level_int = level_order[idx + 1]

        # Есть ли мокрые помещения на верхнем уровне?
        if upper_level_int not in wet_loops_cache:
            continue
        wet_loops = wet_loops_cache[upper_level_int]

        # Получаем контур текущего помещения
        room_loop = get_room_boundary_loop(room, z=0.0)
        if room_loop is None:
            continue

        # Проверяем пересечение с каждым мокрым контуром сверху
        for wl in wet_loops:
            if loops_intersect(room_loop, wl):
                name_p = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                num_p = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                r_name = name_p.AsString() if name_p else ""
                r_number = num_p.AsString() if num_p else ""
                report.append((room.Id.IntegerValue, r_number, r_name))
                break  # Достаточно одного пересечения

    # Вывод отчёта
    if not report:
        alert(
            "Пересечений с мокрыми зонами не обнаружено.\n\n"
            "Стадия: {}\n"
            "Фильтр: {}\n"
            "Проверено помещений: {}".format(
                phase.Name, raw_patterns, len(all_rooms)
            )
        )
        return

    uidoc = __revit__.ActiveUIDocument
    show_report(report, phase.Name, raw_patterns, len(all_rooms), uidoc)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        pass
    except KeyboardInterrupt:
        pass
    except:
        import traceback
        alert(
            "Ошибка при выполнении:\n\n{}".format(traceback.format_exc()),
            title="Ошибка",
        )
