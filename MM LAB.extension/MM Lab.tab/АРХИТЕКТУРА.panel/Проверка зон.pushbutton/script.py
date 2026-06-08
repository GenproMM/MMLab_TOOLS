# -*- coding: utf-8 -*-
__title__ = "Проверка зон для IFC"
__doc__ = "Расчет ТЭП по Площадям (Зонам): СПП и Общая площадь (Жилая/Нежилая, Надземная/Подземная)"
__author__ = "Andrey Panov"

import clr
clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    UnitUtils,
    UnitTypeId,
    StorageType
)
from pyrevit import revit, forms

def get_param_value(element, param_name):
    """Безопасное получение значения параметра по имени"""
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        if param.StorageType == StorageType.String:
            return param.AsString()
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger() # 1 = Да, 0 = Нет
        elif param.StorageType == StorageType.Double:
            return param.AsDouble()
    return None

def calculate_areas_in_doc(doc):
    """Собирает и суммирует площади зон в переданном документе"""
    results = {
        "spp_zh_nad": 0.0, "spp_zh_pod": 0.0,
        "spp_nezh_nad": 0.0, "spp_nezh_pod": 0.0,
        "obsch_zh_nad": 0.0, "obsch_zh_pod": 0.0,
        "obsch_nezh_nad": 0.0, "obsch_nezh_pod": 0.0,
    }
    
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType()
    
    for area in collector:
        name = get_param_value(area, "Имя")
        if not name:
            continue
            
        zh_param = get_param_value(area, "GP_23_ЖилаяЗона")
        pod_param = get_param_value(area, "GP_01_ПодземЭтаж")

        is_zh = (zh_param == 1)
        is_pod = (pod_param == 1)

        try:
            area_val = UnitUtils.ConvertFromInternalUnits(area.Area, UnitTypeId.SquareMeters)
        except Exception:
            area_val = 0.0

        if name == "СПП в ГНС":
            if is_zh and not is_pod: results["spp_zh_nad"] += area_val
            elif is_zh and is_pod: results["spp_zh_pod"] += area_val
            elif not is_zh and not is_pod: results["spp_nezh_nad"] += area_val
            elif not is_zh and is_pod: results["spp_nezh_pod"] += area_val
            
        elif name == "Общая площадь":
            if is_zh and not is_pod: results["obsch_zh_nad"] += area_val
            elif is_zh and is_pod: results["obsch_zh_pod"] += area_val
            elif not is_zh and not is_pod: results["obsch_nezh_nad"] += area_val
            elif not is_zh and is_pod: results["obsch_nezh_pod"] += area_val
            
    return results

def main():
    host_doc = revit.doc
    if not host_doc:
        forms.alert("Не удалось получить активный документ Revit.", exitscript=True)
        return

    # 1. Считаем площади в ОСНОВНОЙ модели
    host_results = calculate_areas_in_doc(host_doc)
    
    # 2. Инициализируем итоговые результаты значениями из основной модели
    total_results = {k: v for k, v in host_results.items()}
    linked_files_count = 0

    # 3. Ищем все связанные файлы (RVT Links)
    link_instances = FilteredElementCollector(host_doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()
    
    for link_inst in link_instances:
        link_doc = link_inst.GetLinkDocument()
        # Проверяем, что файл связи загружен (а не выгружен)
        if link_doc is not None:
            link_results = calculate_areas_in_doc(link_doc)
            # Добавляем значения из связи к итоговым
            for key in total_results:
                total_results[key] += link_results[key]
            linked_files_count += 1

    # 4. Формирование красивого вывода
    output = "✅ Расчет ТЭП завершен!\n"
    output += "Обработано файлов: 1 основной + {} связанных\n".format(linked_files_count)
    output += "═" * 55 + "\n"
    output += "📊 СПП в ГНС:\n"
    output += "  1. Жилая надземная:      {:>10.2f} м²  ({:>10.2f} м²)\n".format(host_results["spp_zh_nad"], total_results["spp_zh_nad"])
    output += "  2. Жилая подземная:      {:>10.2f} м²  ({:>10.2f} м²)\n".format(host_results["spp_zh_pod"], total_results["spp_zh_pod"])
    output += "  3. Нежилая надземная:    {:>10.2f} м²  ({:>10.2f} м²)\n".format(host_results["spp_nezh_nad"], total_results["spp_nezh_nad"])
    output += "  4. Нежилая подземная:    {:>10.2f} м²  ({:>10.2f} м²)\n".format(host_results["spp_nezh_pod"], total_results["spp_nezh_pod"])
    output += "─" * 55 + "\n"
    output += "📊 Общая площадь:\n"
    output += "  5. Жилая надземная:      {:>10.2f} м²  ({:>10.2f} м²)\n".format(host_results["obsch_zh_nad"], total_results["obsch_zh_nad"])
    output += "  6. Жилая подземная:      {:>10.2f} м²  ({:>10.2f} м²)\n".format(host_results["obsch_zh_pod"], total_results["obsch_zh_pod"])
    output += "  7. Нежилая надземная:    {:>10.2f} м²  ({:>10.2f} м²)\n".format(host_results["obsch_nezh_nad"], total_results["obsch_nezh_nad"])
    output += "  8. Нежилая подземная:    {:>10.2f} м²  ({:>10.2f} м²)\n".format(host_results["obsch_nezh_pod"], total_results["obsch_nezh_pod"])
    output += "═" * 55 + "\n"
    output += "💡 Формат: Значение в текущей модели (Общая сумма со связями)"

    # Вывод результата
    forms.alert(output, title="Результаты расчета ТЭП", exitscript=True)

if __name__ == '__main__':
    main()