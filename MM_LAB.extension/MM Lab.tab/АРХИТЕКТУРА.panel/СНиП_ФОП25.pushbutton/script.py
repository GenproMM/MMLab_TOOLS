# -*- coding: utf-8 -*-
__title__ = 'Проверка площадей'
__author__ = 'Akhmedyanova Eleonora'
__doc__= 'Проверяет площади помещений на соответствие требованиям СНиП (ФОП 2025)'

import os
import sys
import re
import System
from collections import defaultdict
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

doc = revit.doc
app = doc.Application
output = script.get_output()

# КОЭФФИЦИЕНТ ПЕРЕВОДА ФУТОВ В МЕТРЫ
FT2_TO_M2 = 0.09290304

# Путь к новому корпоративному ФОП
SHARED_PARAM_PATH = r"\\srv-dfs\BIM\01_Ресурсы плагинов\2_Shared_parameters\ГП_ФОП2025.txt"

# Корни слов для распознавания кухонь, кухонных зон, ниш, столовых
KITCHEN_KEYWORDS = ["кух", "зон", "ниш", "столов", "гостин"]

# ===================================================================
# GUID НОВЫХ ПАРАМЕТРОВ (ФОП 2025)
# ===================================================================

GUID_MAP = {
    "CHECK":      "b132e98d-042b-44a9-a7cf-b30fe39cf28c",  # GP_12_ПлПроверка_СНиП
    "COMMENT":    "b204ef76-2c63-4f66-9f73-f734deaa01bb",  # GP_12_ПлКомментарии_СНиП
    "AREA":       "d63f9990-746b-46b5-a83b-79ffdc0c6066",  # GP_16_Площадь
    "PURPOSE":    "0b3dbc34-30a7-4278-b1c5-8ba8819f9db4",  # GP_23_Назначение
    "APT_NUM":    "97c93b07-9641-4947-ac3b-eb147295a63a",  # GP_23_НомерКвНаЭт
    "ROOM_COUNT": "1db9f095-f768-4b75-8a57-61b4baf871de",  # GP_23_КолвоКомнат
    "ROOM_TYPE":  "5ca3e2e8-7b95-4bcb-af38-1273587b13de"   # GP_23_ТипПомещения_Номер
}

# ===================================================================
# 1. ПРИВЯЗКА ПАРАМЕТРОВ ИЗ ФОП 2025
# ===================================================================

original_sp_file = app.SharedParametersFilename

try:
    if os.path.exists(SHARED_PARAM_PATH):
        app.SharedParametersFilename = SHARED_PARAM_PATH
        sp_file = app.OpenSharedParameterFile()

        if sp_file:
            guids_to_bind = [
                System.Guid(GUID_MAP["CHECK"]),
                System.Guid(GUID_MAP["COMMENT"])
            ]

            definitions_to_bind = []
            for tg in guids_to_bind:
                found_def = None
                for group in sp_file.Groups:
                    for d in group.Definitions:
                        if d.GUID == tg:
                            found_def = d
                            break
                    if found_def: break
                if found_def: definitions_to_bind.append(found_def)

            if definitions_to_bind:
                category = Category.GetCategory(doc, BuiltInCategory.OST_Rooms)
                category_set = app.Create.NewCategorySet()
                category_set.Insert(category)

                with revit.Transaction("Привязка параметров СНиП"):
                    bound_param_names = set()
                    iterator = doc.ParameterBindings.ForwardIterator()
                    while iterator.MoveNext():
                        bound_param_names.add(iterator.Key.Name)

                    for definition in definitions_to_bind:
                        if definition.Name not in bound_param_names:
                            binding = app.Create.NewInstanceBinding(category_set)
                            doc.ParameterBindings.Insert(definition, binding, GroupTypeId.Data)

finally:
    if original_sp_file:
        app.SharedParametersFilename = original_sp_file

# ===================================================================
# 2. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ===================================================================

def get_param_val(element, key_name):
    """Считывает значение параметра по его GUID из ФОП 2025."""
    guid_str = GUID_MAP.get(key_name)
    if not guid_str:
        return ""
        
    param = element.get_Parameter(System.Guid(guid_str))
    if param and param.HasValue:
        if param.StorageType == StorageType.String:
            val = param.AsString()
            return val.strip() if val else ""
        elif param.StorageType == StorageType.Double:
            return param.AsDouble()
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger()
    return ""

def get_room_area(element):
    """
    Считывает площадь помещения из GP_16_Площадь.
    Если параметр пуст, использует встроенную системную площадь Revit.
    """
    area_val = get_param_val(element, "AREA")
    
    if area_val != "" and area_val is not None:
        try:
            if isinstance(area_val, (float, int)):
                return float(area_val) * FT2_TO_M2 if area_val > 100 else float(area_val)
            elif isinstance(area_val, str):
                cleaned = re.sub(r'[^\d.,]', '', area_val).replace(',', '.')
                return float(cleaned)
        except:
            pass

    # Фолбэк на родную системную площадь Revit
    sys_area_param = element.get_Parameter(BuiltInParameter.ROOM_AREA)
    if sys_area_param and sys_area_param.HasValue:
        return sys_area_param.AsDouble() * FT2_TO_M2

    return None

def get_room_name(element):
    """Получает встроенное имя помещения Revit."""
    p = element.get_Parameter(BuiltInParameter.ROOM_NAME)
    if p and p.HasValue:
        return p.AsString().strip()
    return "Помещение"

def is_kitchen_room(room_name):
    """Проверяет, относится ли имя помещения к кухне/кухонной зоне/нише/столовой/гостиной."""
    name_lower = room_name.lower()
    return any(kw in name_lower for kw in KITCHEN_KEYWORDS)

def set_param_val(element, key_name, value):
    """Записывает значение в параметр ФОП 2025."""
    guid_str = GUID_MAP.get(key_name)
    if guid_str:
        p = element.get_Parameter(System.Guid(guid_str))
        if p:
            p.Set(value)
            return True
    return False

# ===================================================================
# 3. ПОЛУЧАЕМ ВСЕ ПОМЕЩЕНИЯ И ФИЛЬТРУЕМ ПО НАЗНАЧЕНИЮ
# ===================================================================

rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

if not rooms:
    forms.alert("В проекте нет помещений!", exitscript=True)
    sys.exit()

filtered_rooms = []
for room in rooms:
    nazn = get_param_val(room, "PURPOSE")
    nazn_str = str(nazn).lower().strip() if nazn else ""
    if "квартир" in nazn_str:
        filtered_rooms.append(room)

if not filtered_rooms:
    forms.alert("Нет помещений с 'квартир' в параметре GP_23_Назначение!", exitscript=True)
    sys.exit()

# ===================================================================
# 4. ГРУППИРУЕМ ПО ЭТАЖУ + НОМЕРУ КВАРТИРЫ
# ===================================================================

apartments = {}

for room in filtered_rooms:
    apt_num = get_param_val(room, "APT_NUM")
    if apt_num is None or apt_num == "":
        continue
    
    level_name = "Без уровня"
    if room.Level:
        level_name = room.Level.Name
    
    unique_id = "{}_{}".format(str(level_name), str(apt_num).strip())
    
    if unique_id not in apartments:
        apartments[unique_id] = {
            'level': level_name,
            'apt_num': "Квартира №{}".format(str(apt_num).strip()),
            'rooms': []
        }
    apartments[unique_id]['rooms'].append(room)

if len(apartments) == 0:
    forms.alert("Не найдено ни одной квартиры с заполненной информацией!", exitscript=True)
    sys.exit()

# ===================================================================
# 5. АНАЛИЗ КАЖДОЙ КВАРТИРЫ И ПРОВЕРКА ПЛОЩАДЕЙ
# ===================================================================

violations_list = []
rooms_breakdown = defaultdict(int)

t = Transaction(doc, "Проверка площадей кухонь и жилых комнат")
t.Start()

try:
    for unique_id, apt_data in apartments.items():
        apt_rooms = apt_data['rooms']
        first_room = apt_rooms[0]
        rooms_count = get_param_val(first_room, "ROOM_COUNT")
        
        try:
            rooms_count_int = int(rooms_count) if rooms_count != "" else 0
        except:
            rooms_count_int = 0
            
        if rooms_count_int < 1:
            continue
        
        rooms_breakdown[rooms_count_int] += 1

        # Устанавливаем статус по умолчанию (1 = успешно)
        for room in apt_rooms:
            set_param_val(room, "CHECK", 1)
            set_param_val(room, "COMMENT", "")
            
        # ----------------- СЦЕНАРИЙ 1к КВАРТИР -----------------
        if rooms_count_int == 1:
            kitchen = None
            living_room = None
            
            for room in apt_rooms:
                room_name = get_room_name(room)
                room_type = str(get_param_val(room, "ROOM_TYPE")).strip()
                
                if is_kitchen_room(room_name):
                    kitchen = room
                if room_type == "1":
                    living_room = room
            
            if kitchen:
                k_area = get_room_area(kitchen)
                if k_area is not None and k_area < 5.0:
                    msg = "Площадь кухни для студий и 1к квартир должна быть больше 5м2"
                    set_param_val(kitchen, "CHECK", 0)
                    set_param_val(kitchen, "COMMENT", msg)
                    violations_list.append({
                        'room': kitchen, 'id': kitchen.Id, 
                        'level': apt_data['level'], 'apt_num': apt_data['apt_num'],
                        'rooms_count': rooms_count_int,
                        'area': k_area, 'error': msg
                    })
            
            if living_room:
                l_area = get_room_area(living_room)
                if l_area is not None and l_area < 14.0:
                    msg = "Площадь жилой комнаты 1к квартиры должна быть больше 14м2"
                    set_param_val(living_room, "CHECK", 0)
                    set_param_val(living_room, "COMMENT", msg)
                    violations_list.append({
                        'room': living_room, 'id': living_room.Id, 
                        'level': apt_data['level'], 'apt_num': apt_data['apt_num'],
                        'rooms_count': rooms_count_int,
                        'area': l_area, 'error': msg
                    })

        # ----------------- СЦЕНАРИЙ 2к+ КВАРТИР -----------------
        elif rooms_count_int >= 2:
            kitchen_niche = None
            living_rooms = []
            
            for room in apt_rooms:
                room_name = get_room_name(room)
                room_type = str(get_param_val(room, "ROOM_TYPE")).strip()
                
                if is_kitchen_room(room_name):
                    kitchen_niche = room
                if room_type == "1":
                    living_rooms.append(room)
            
            if kitchen_niche:
                kn_area = get_room_area(kitchen_niche)
                if kn_area is not None and kn_area < 6.0:
                    msg = "Площадь кухни/кухонной зоны для 2к+ квартир должна быть не менее 6м2"
                    set_param_val(kitchen_niche, "CHECK", 0)
                    set_param_val(kitchen_niche, "COMMENT", msg)
                    violations_list.append({
                        'room': kitchen_niche, 'id': kitchen_niche.Id, 
                        'level': apt_data['level'], 'apt_num': apt_data['apt_num'],
                        'rooms_count': rooms_count_int,
                        'area': kn_area, 'error': msg
                    })
                    
            if living_rooms:
                living_rooms_with_area = []
                for lr in living_rooms:
                    area = get_room_area(lr)
                    if area is not None:
                        living_rooms_with_area.append((lr, area))
                
                if living_rooms_with_area:
                    living_rooms_with_area.sort(key=lambda x: x[1], reverse=True)
                    max_living_room, max_area = living_rooms_with_area[0]
                    
                    if max_area < 16.0:
                        msg = "Площадь одной из жилых комнат 2к и более квартир должна быть больше 16м2"
                        set_param_val(max_living_room, "CHECK", 0)
                        set_param_val(max_living_room, "COMMENT", msg)
                        violations_list.append({
                            'room': max_living_room, 'id': max_living_room.Id, 
                            'level': apt_data['level'], 'apt_num': apt_data['apt_num'],
                            'rooms_count': rooms_count_int,
                            'area': max_area, 'error': msg
                        })

    t.Commit()
    
except Exception as e:
    t.RollBack()
    forms.alert("Ошибка: " + str(e), exitscript=True)

# ===================================================================
# 6. ВЫВОД ИНТЕРАКТИВНОГО ОКНА С РЕЗУЛЬТАТАМИ
# ===================================================================

breakdown_lines = []
for k in sorted(rooms_breakdown.keys()):
    breakdown_lines.append("  -{}-комнатных: {}".format(k, rooms_breakdown[k]))

breakdown_str = "\n".join(breakdown_lines)
if breakdown_str:
    breakdown_str = "\n" + breakdown_str

alert_message = (
    "Анализ завершён!\n\n"
    "Всего квартир: {}{}\n\n"
    "Помещений с нарушениями: {}".format(
        len(apartments), breakdown_str, len(violations_list)
    )
)

forms.alert(alert_message)

if violations_list:
    output.set_width(1200)
    output.set_height(600)
    output.print_md("## Внимание! Обнаружены несоответствия нормативам площади")
    output.print_md("*Вы можете нажать на ID элемента, чтобы найти его на виде Revit.*")
    print("-" * 140)
    
    header_format = "{:<22} | {:<15} | {:<13} | {:<18} | {:<25} | {:<15} | {}"
    row_format    = "{:<22} | {:<15} | {:<13} | {:<18} | {:<25} | {:<15.2f} | ❌ {}"
    
    print(header_format.format(
        "Уровень", "НомерКвартиры", "ID Помещения", "КоличествоКомнат", "ИмяПомещения", "ТекущаяПлощадь", "ТекстОшибки"
    ))
    print("-" * 140)
    
    for item in violations_list:
        room_link = output.linkify(item['id'])
        room_name = get_room_name(item['room'])
        
        print(row_format.format(
            item['level'],
            item['apt_num'],
            room_link,
            item['rooms_count'],
            room_name,
            item['area'],
            item['error']
        ))
    
    print("-" * 140)
    print("Всего нарушений в таблице: {}".format(len(violations_list)))
else:
    output.print_md("### 🎉 Прекрасно! Ни одного нарушения нормативных площадей не обнаружено.")