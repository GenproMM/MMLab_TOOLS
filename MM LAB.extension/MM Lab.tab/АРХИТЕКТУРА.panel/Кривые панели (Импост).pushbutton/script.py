# -*- coding: utf-8 -*-

__title__ = "Позиционирование импостов"
__author__ = 'Panov Andrey'
__doc__ = "Установка параметров Импост Снизу/Сверху по положению панели в витражной стене"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("RevitServices")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Wall,
    FamilyInstance,
    Transaction,
    BuiltInCategory
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# =====================================================================
# НАСТРОЙКИ (редактируйте здесь)
# =====================================================================
# Допуск в футах для определения "касается ли панель верха/низа стены".
# 0.5 фута ≈ 15 см. Уменьшите значение, если панели "перекрывают" стену,
# увеличьте, если между панелью и краем стены есть небольшой зазор.
TOLERANCE = 0.5

# Имена параметров в загружаемом семействе панели (экземплярные, Да/Нет)
PARAM_BOTTOM = "Импост Снизу"
PARAM_TOP = "Импост Сверху"
# =====================================================================


# Максимально строгий фильтр выбора: разрешает выбирать ТОЛЬКО объекты класса Wall
class WallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        # isinstance проверяет, является ли элемент именно Стеной (включая витражные)
        # Это самый надежный способ в Revit API, исключающий ошибки с категориями
        return isinstance(element, Wall)

    def AllowReference(self, reference, position):
        # Категорически запрещаем выбор граней, ребер, линий или любых других ссылок
        return False


def get_panel_contacts(panel, wall):
    """
    Возвращает (touches_bottom, touches_top) —
    касается ли панель низа и/или верха стены.
    """
    bbox = panel.get_BoundingBox(None)
    wall_bbox = wall.get_BoundingBox(None)
    if not bbox or not wall_bbox:
        return (False, False)

    p_min, p_max = bbox.Min.Z, bbox.Max.Z
    w_min, w_max = wall_bbox.Min.Z, wall_bbox.Max.Z

    touches_bottom = (p_min <= w_min + TOLERANCE)
    touches_top = (p_max >= w_max - TOLERANCE)
    return (touches_bottom, touches_top)


# --- 1. Запрос выбора стен пользователем ---
try:
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        WallSelectionFilter(),
        "Выберите витражные стены. Нажмите 'Готово' (или Esc для отмены)."
    )
except OperationCanceledException:
    TaskDialog.Show("Отмена", "Выбор отменён пользователем.")
    exit()
except Exception as ex:
    TaskDialog.Show("Ошибка", "Не удалось получить выбор: " + str(ex))
    exit()

# --- 2. Фильтрация и подготовка (двойная защита) ---
walls = []
for r in refs:
    elem = doc.GetElement(r)
    # Еще раз проверяем, что это точно стена, прежде чем добавить в список
    if isinstance(elem, Wall):
        walls.append(elem)

if not walls:
    TaskDialog.Show("Внимание", "Стены не выбраны или выбраны некорректно.")
    exit()


# --- 3. Основная логика ---
processed = 0
failed_ids = []  # сюда попадают панели, где что-то пошло не так

t = Transaction(doc, "Импосты панелей витража")
t.Start()

try:
    # Собираем все панели витража в проекте один раз для оптимизации скорости
    all_panels = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_CurtainWallPanels) \
        .WhereElementIsNotElementType()

    for wall in walls:
        for panel in all_panels:
            # Проверяем, что панель привязана именно к этой стене
            if panel.Host is None or panel.Host.Id.IntegerValue != wall.Id.IntegerValue:
                continue

            touches_bottom, touches_top = get_panel_contacts(panel, wall)

            bottom_param = panel.LookupParameter(PARAM_BOTTOM)
            top_param = panel.LookupParameter(PARAM_TOP)

            is_failed = False

            if bottom_param and not bottom_param.IsReadOnly:
                bottom_param.Set(True)  # по ТЗ — всегда включён
            else:
                is_failed = True

            if top_param and not top_param.IsReadOnly:
                top_param.Set(touches_top)  # включается только у верхних
            else:
                is_failed = True

            if is_failed:
                failed_ids.append(str(panel.Id))
            else:
                processed += 1

    t.Commit()
except Exception as ex:
    if t.HasStarted():
        t.RollBack()
    TaskDialog.Show("Ошибка", "Произошла ошибка: " + str(ex))
    exit()


# --- 4. Короткий отчёт ---
msg = "Готово! Обработано панелей: {}".format(processed)
if failed_ids:
    unique_failed = sorted(set(failed_ids))
    msg += "\n\nПроблемных панелей: {} — нет параметра или он ReadOnly.\nID: {}".format(
        len(unique_failed), ", ".join(unique_failed)
    )

TaskDialog.Show("Результат", msg)