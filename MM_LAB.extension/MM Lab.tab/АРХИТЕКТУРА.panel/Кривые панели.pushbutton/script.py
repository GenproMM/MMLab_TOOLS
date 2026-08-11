# -*- coding: utf-8 -*-

__title__ = "Кривые панели"
__author__ = 'Panov Andrey'
__doc__ = "Редактирование параметра радиус у панелей витража"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Wall,
    Arc,
    LocationCurve,
    FamilyInstance,
    Transaction,
    BuiltInCategory
)
from Autodesk.Revit.UI import TaskDialog

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def get_wall_radius(wall):
    location = wall.Location
    if hasattr(location, 'Curve') and isinstance(location, LocationCurve):
        curve = location.Curve
        if isinstance(curve, Arc):
            return curve.Radius
    return None

# Сбор стен: либо по выбору, либо все в проекте
selection = uidoc.Selection.GetElementIds()
if selection:
    walls = [doc.GetElement(id) for id in selection if doc.GetElement(id).Category.Id.IntegerValue == int(BuiltInCategory.OST_Walls)]
else:
    walls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()

curved_walls_found = 0
panels_processed = 0

t = Transaction(doc, "Запись радиуса стен в загружаемые панели витража")
t.Start()

try:
    for wall in walls:
        radius = get_wall_radius(wall)
        if radius is not None:
            curved_walls_found += 1
            # Находим все FamilyInstance (включая панели витража), привязанные к этой стене
            family_instances = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
            for instance in family_instances:
                host = instance.Host
                if host and host.Id.IntegerValue == wall.Id.IntegerValue:
                    # Проверяем, есть ли у экземпляра параметр GP_16_Радиус
                    param = instance.LookupParameter("GP_16_Радиус")
                    if param and param.IsReadOnly == False:
                        param.Set(radius)
                        panels_processed += 1

    t.Commit()
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    TaskDialog.Show("Ошибка", "Произошла ошибка: " + str(e))
    exit()

summary = (
    "Обнаружено криволинейных стен: {}\n"
    "Обработано панелей: {}"
).format(curved_walls_found, panels_processed)

TaskDialog.Show("Результат", summary)