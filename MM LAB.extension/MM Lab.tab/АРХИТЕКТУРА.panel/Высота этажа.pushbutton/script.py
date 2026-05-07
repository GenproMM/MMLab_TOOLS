# -*- coding: utf-8 -*-

__title__ = "Высота этажа"
__author__ = 'Panov Andrey'
__doc__ = "Расчет высоты между уровней для изменения высоты помещений"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Level,
    Floor,
    BuiltInParameter,
    Transaction,
    BuiltInCategory
)
from Autodesk.Revit.UI import TaskDialog

# Функция для конвертации футов в метры
def feet_to_meters(feet_value):
    return round(feet_value * 0.3048, 3)

doc = __revit__.ActiveUIDocument.Document

# Функция для получения толщины перекрытия
def get_floor_thickness(floor):
    floor_type = doc.GetElement(floor.GetTypeId())
    
    # Сначала пробуем получить через параметр Thickness
    param = floor_type.LookupParameter("Thickness")
    if param and param.HasValue:
        return param.AsDouble()

    # Если не получилось, пробуем через Compound Structure
    compound_structure = floor_type.GetCompoundStructure()
    if compound_structure:
        total_thickness = 0.0
        layers = compound_structure.GetLayers()
        for layer in layers:
            total_thickness += layer.Width
        return total_thickness

    return 0.0

# Получаем все уровни
levels_collector = FilteredElementCollector(doc).OfClass(Level)
levels = list(levels_collector)

if not levels:
    TaskDialog.Show("Ошибка", "Уровни не найдены.")
    exit()

# Сортируем уровни по высоте
levels.sort(key=lambda x: x.Elevation)

# Подготовим отчёт
report_lines = []

# Начинаем транзакцию
t = Transaction(doc, "Запись высот между уровнями и копирование в помещения")
t.Start()

try:
    for i in range(len(levels)):
        current_level = levels[i]

        if i < len(levels) - 1:
            next_level = levels[i + 1]
            height = next_level.Elevation - current_level.Elevation

            # Находим перекрытие, расположенное на следующем уровне
            floors_collector = FilteredElementCollector(doc).OfClass(Floor).WhereElementIsNotElementType()
            closest_floor = None
            min_dist = float('inf')

            for floor in floors_collector:
                base_offset_param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
                level_id = floor.LevelId
                floor_level = doc.GetElement(level_id)

                if floor_level.Id.IntegerValue == next_level.Id.IntegerValue and base_offset_param:
                    elevation = next_level.Elevation + base_offset_param.AsDouble()
                    dist = abs(elevation - next_level.Elevation)
                    if dist < min_dist:
                        min_dist = dist
                        closest_floor = floor

            thickness = 0.0
            if closest_floor:
                thickness = get_floor_thickness(closest_floor)

            final_height = height - thickness

            # Конвертируем в метры для отчёта
            height_m = feet_to_meters(height)
            thickness_m = feet_to_meters(thickness)
            final_height_m = feet_to_meters(final_height)

            report_lines.append(
                "{} → {}: {} м - {} м = {} м".format(
                    current_level.Name,
                    next_level.Name,
                    height_m,
                    thickness_m,
                    final_height_m
                )
            )

        else:
            final_height = 0  # Для последнего уровня
            report_lines.append(
                "{}: 0 м".format(current_level.Name)
            )

        # Записываем значение в пользовательский параметр уровня
        param_name = "GP_16_Высота"
        param = current_level.LookupParameter(param_name)

        if param and param.IsReadOnly == False:
            param.Set(final_height)
        else:
            TaskDialog.Show("Ошибка", "Параметр '" + param_name + "' не найден или недоступен для редактирования у уровня " + current_level.Name)

        # Теперь копируем значение в параметр "Высота этажа" у помещений на уровне
        rooms_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
        for room in rooms_collector:
            room_level = doc.GetElement(room.LevelId)
            if room_level and room_level.Id.IntegerValue == current_level.Id.IntegerValue:
                room_param = room.LookupParameter("Высота этажа")
                if room_param and room_param.IsReadOnly == False:
                    room_param.Set(final_height)

    t.Commit()
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    TaskDialog.Show("Ошибка", "Произошла ошибка: " + str(e))
    exit()

# Отображаем отчёт
report_text = "\n".join(report_lines)
TaskDialog.Show("Готово", "Расчёты завершены.\n\n" + report_text)