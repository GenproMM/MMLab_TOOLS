# -*- coding: utf-8 -*-
"""
GP_01_СсылкаНаЛист - Автоматическое заполнение параметра ссылки на лист
для спецификации Ведомости Объемов Работ (ВОР)

Автор: 
Версия: 1.0.4
Revit: 2024+
"""

import sys
import os
import clr
import time
from collections import defaultdict

# Добавляем ссылки на Revit API
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import (
    Document,
    View,
    ViewSheet,
    ViewPlan,
    Element,
    ElementId,
    Category,
    BuiltInCategory,
    BuiltInParameter,
    Parameter,
    FilteredElementCollector,
    Transaction,
    ViewType,
    ElementCategoryFilter,
    BrowserOrganization
)

from Autodesk.Revit.UI import (
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogResult
)

# pyRevit
from pyrevit import forms
from pyrevit import script

import re


def natural_sort_key(s):
    """
    Ключ для естественной сортировки (natural sort)
    Позволяет сортировать: 1, 2, 3... 9, 10, 11 вместо 1, 10, 11, 2...
    """
    def convert(text):
        return int(text) if text.isdigit() else text.lower()
    return [convert(c) for c in re.split('([0-9]+)', str(s))]


def natural_sorted(iterable, key=None):
    """Сортировка с естественным порядком чисел"""
    if key:
        return sorted(iterable, key=lambda x: natural_sort_key(key(x)))
    return sorted(iterable, key=natural_sort_key)


# ==============================================================================
# КОНСТАНТЫ
# ==============================================================================

# Параметр для записи
PARAM_SSYLKA_NA_LIST = "GP_01_СсылкаНаЛист"

# Параметры для чтения
PARAM_VID_RABOTY = "GP_01_ВидРаботы"
PARAM_ZONA = "GP_01_Зона"

# Параметры материалов
PARAM_MATERIAL_DESCRIPTION = "Описание"
PARAM_MATERIAL_KOD_EDINICY = "GP_01_КодЕдиницыЭкз"

# Параметр элемента
PARAM_DESCRIPTION = "Описание"

# Параметры группировки листов (приоритет)
SHEET_GROUPING_PARAMS = [
    "GP_01_КомплектВида",
    "GP_01_ГруппаВида",
    "GP_01_НазначениеВида",
    "GP_01_ВладелецВида"
]

# Категории для МАТЕРИАЛОВ (спецификация материалов)
CATEGORIES_MATERIALS = [
    BuiltInCategory.OST_Walls,           # Стены
    BuiltInCategory.OST_Floors,          # Перекрытия
    BuiltInCategory.OST_Roofs,           # Крыши
    BuiltInCategory.OST_Ceilings,        # Потолки
]

# Категории для ЭЛЕМЕНТОВ (спецификация элементов)
CATEGORIES_ELEMENTS = [
    BuiltInCategory.OST_Doors,                    # Двери
    BuiltInCategory.OST_Windows,                  # Окна
    BuiltInCategory.OST_GenericModel,             # Обобщенные модели
    BuiltInCategory.OST_MechanicalEquipment,      # Оборудование
    BuiltInCategory.OST_StairsRailing,            # Ограждения
    BuiltInCategory.OST_PlumbingFixtures,         # Сантехнические приборы
    BuiltInCategory.OST_SpecialityEquipment,      # Специальное оборудование
]

# Все категории для обработки
ALL_CATEGORIES = CATEGORIES_MATERIALS + CATEGORIES_ELEMENTS


# ==============================================================================
# КЛАССЫ ДАННЫХ
# ==============================================================================

class TreeNode:
    """Узел иерархического дерева для UI"""
    
    def __init__(self, name, level_type, data=None):
        self.Name = name
        self.LevelType = level_type  # "sorting", "sheet", "plan"
        self.Data = data  # ViewSheet или ViewPlan
        self.Children = []
        self.IsChecked = False
        self.Parent = None
    
    def add_child(self, child):
        child.Parent = self
        self.Children.append(child)


# ==============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==============================================================================

def get_parameter_value(element, param_name):
    """Получить значение параметра элемента по имени"""
    param = element.LookupParameter(param_name)
    if param:
        storage_type = param.StorageType.ToString()
        if storage_type == "String":
            return param.AsString()
        elif storage_type == "Integer":
            return param.AsInteger()
        elif storage_type == "Double":
            return param.AsDouble()
        elif storage_type == "ElementId":
            return param.AsElementId()
    return None


def get_parameter_value_string(param):
    """Получить строковое значение параметра"""
    if param is None:
        return None
    storage_type = param.StorageType.ToString()
    if storage_type == "String":
        return param.AsString()
    elif storage_type == "Integer":
        val = param.AsInteger()
        return str(val) if val != 0 else None
    elif storage_type == "Double":
        return param.AsValueString()
    elif storage_type == "ElementId":
        return str(param.AsElementId())
    return None


def get_element_description(element):
    """Получить описание элемента"""
    for param_name in ["Описание", "Description", "ADSK_Описание"]:
        desc = get_parameter_value(element, param_name)
        if desc:
            return desc
    
    param = element.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION)
    if param:
        return param.AsString()
    
    return ""


def get_materials_with_kod_edinicy(element, doc):
    """Получить материалы элемента с заполненным GP_01_КодЕдиницыЭкз"""
    materials = []
    
    try:
        material_ids = element.GetMaterialIds(False)
        
        for mat_id in material_ids:
            material = doc.GetElement(mat_id)
            if material:
                mat_desc_param = material.LookupParameter(PARAM_MATERIAL_DESCRIPTION)
                mat_description = get_parameter_value_string(mat_desc_param) or ""
                
                kod_param = material.LookupParameter(PARAM_MATERIAL_KOD_EDINICY)
                kod_value = get_parameter_value_string(kod_param)
                
                if kod_value:
                    materials.append((mat_description, kod_value))
    
    except Exception as ex:
        print("Ошибка при получении материалов: {}".format(str(ex)))
    
    return materials


def determine_zone(zone_value):
    """Определить зону из значения параметра"""
    if not zone_value:
        return None
    
    zone_lower = zone_value.lower()
    
    if "подземн" in zone_lower:
        return "Подземная"
    elif "надземн" in zone_lower:
        return "Надземная"
    
    return None


def is_material_category(category):
    """Проверить, относится ли категория к материалам"""
    if category is None:
        return False
    try:
        category_id = category.Id
        for bic in CATEGORIES_MATERIALS:
            if category_id == ElementId(int(bic)):
                return True
    except:
        pass
    return False


def get_sheet_grouping_value(sheet):
    """Получить значение группировки для листа"""
    for param_name in SHEET_GROUPING_PARAMS:
        param = sheet.LookupParameter(param_name)
        if param:
            value = param.AsString()
            if value:
                return value
    return None


def get_floor_plans_on_sheet(sheet, doc):
    """Получить все планы этажей на листе"""
    floor_plans = []
    
    try:
        view_ids = sheet.GetAllPlacedViews()
        
        for view_id in view_ids:
            view = doc.GetElement(view_id)
            
            if view and view.ViewType == ViewType.FloorPlan:
                floor_plans.append(view)
    except Exception as ex:
        print("Ошибка при получении планов с листа: {}".format(str(ex)))
    
    return floor_plans


def build_sheet_tree(doc):
    """Построить дерево листов на основе текущей сортировки проекта"""
    
    # Получаем все листы
    sheets_collector = FilteredElementCollector(doc).OfClass(ViewSheet)
    all_sheets = [sheet for sheet in sheets_collector if not sheet.IsPlaceholder]
    
    if not all_sheets:
        return None
    
    # Корневой узел
    root = TreeNode("Листы проекта", "root")
    
    # Группируем листы по параметру сортировки
    groups = defaultdict(list)
    
    for sheet in all_sheets:
        group_value = get_sheet_grouping_value(sheet)
        if group_value:
            groups[group_value].append(sheet)
        else:
            groups["[Без группы]"].append(sheet)
    
    # Сортируем группы естественным образом
    sorted_group_names = natural_sorted(groups.keys())
    
    for group_name in sorted_group_names:
        # Создаем узел группы
        group_node = TreeNode(group_name, "sorting")
        
        # Сортируем листы в группе по номеру (естественная сортировка)
        sorted_sheets = natural_sorted(groups[group_name], key=lambda s: s.SheetNumber)
        
        for sheet in sorted_sheets:
            # Создаем узел листа
            sheet_node = TreeNode(
                "{} - {}".format(sheet.SheetNumber, sheet.Name),
                "sheet",
                sheet
            )
            
            # Находим планы на этом листе
            view_plans = get_floor_plans_on_sheet(sheet, doc)
            
            for view_plan in view_plans:
                # Создаем узел плана
                plan_node = TreeNode(view_plan.Name, "plan", view_plan)
                sheet_node.add_child(plan_node)
            
            # Добавляем лист только если на нем есть планы
            if len(sheet_node.Children) > 0:
                group_node.add_child(sheet_node)
        
        # Добавляем группу только если в ней есть листы с планами
        if len(group_node.Children) > 0:
            root.add_child(group_node)
    
    return root


def get_elements_on_view(doc, view):
    """Получить элементы заданных категорий на виде"""
    elements = []
    
    for bic in ALL_CATEGORIES:
        try:
            collector = FilteredElementCollector(doc, view.Id)
            category_filter = ElementCategoryFilter(bic)
            elements_in_view = collector.WherePasses(category_filter).WhereElementIsNotElementType().ToElements()
            elements.extend(elements_in_view)
        except:
            continue
    
    return elements


def format_sheet_reference(sheet_numbers):
    """Форматировать ссылку на листы"""
    if not sheet_numbers:
        return ""
    
    sorted_numbers = sorted(sheet_numbers)
    
    return "см.лист " + ", ".join(sorted_numbers)


def get_sheet_number_for_view(doc, view_plan):
    """Получить номер листа для вида"""
    try:
        sheets = FilteredElementCollector(doc).OfClass(ViewSheet)
        
        for sheet in sheets:
            view_ids = sheet.GetAllPlacedViews()
            if view_plan.Id in view_ids:
                return sheet.SheetNumber
    except:
        pass
    
    return None


# ==============================================================================
# UI ФОРМА
# ==============================================================================

class SheetSelectionWindow:
    """Окно выбора листов и планов"""
    
    def __init__(self, tree_root):
        from System.Windows import Window
        
        self.TreeRoot = tree_root
        self.SelectedPlans = []
        self.DialogResult = False
        
        self.window = Window()
        self.window.Title = "Выбор видов для обработки"
        self.window.Width = 600
        self.window.Height = 700
        
        self.create_ui()
    
    def create_ui(self):
        """Создать элементы UI"""
        from System.Windows import (
            Thickness, HorizontalAlignment, VerticalAlignment,
            GridLength, GridUnitType
        )
        from System.Windows.Controls import (
            Grid as WpfGrid, StackPanel, TreeView, TreeViewItem,
            Button, CheckBox, ScrollViewer, TextBlock, RowDefinition,
            ScrollBarVisibility, Orientation
        )
        
        main_grid = WpfGrid()
        main_grid.Margin = Thickness(10)
        
        row1 = RowDefinition()
        row1.Height = GridLength(1, GridUnitType.Star)
        row2 = RowDefinition()
        row2.Height = GridLength.Auto
        main_grid.RowDefinitions.Add(row1)
        main_grid.RowDefinitions.Add(row2)
        
        scroll_viewer = ScrollViewer()
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll_viewer.Margin = Thickness(0, 0, 0, 10)
        WpfGrid.SetRow(scroll_viewer, 0)
        
        self.tree_view = TreeView()
        self.tree_view.Margin = Thickness(5)
        
        self.populate_tree(self.TreeRoot, self.tree_view.Items)
        
        scroll_viewer.Content = self.tree_view
        
        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Right
        button_panel.Margin = Thickness(0, 10, 0, 0)
        WpfGrid.SetRow(button_panel, 1)
        
        btn_select_all = Button()
        btn_select_all.Content = "Выделить всё"
        btn_select_all.Width = 100
        btn_select_all.Margin = Thickness(5)
        btn_select_all.Click += self.on_select_all
        
        btn_deselect_all = Button()
        btn_deselect_all.Content = "Снять выделение"
        btn_deselect_all.Width = 110
        btn_deselect_all.Margin = Thickness(5)
        btn_deselect_all.Click += self.on_deselect_all
        
        btn_execute = Button()
        btn_execute.Content = "Выполнить"
        btn_execute.Width = 100
        btn_execute.Margin = Thickness(5)
        btn_execute.Click += self.on_execute
        
        btn_cancel = Button()
        btn_cancel.Content = "Отмена"
        btn_cancel.Width = 80
        btn_cancel.Margin = Thickness(5)
        btn_cancel.Click += self.on_cancel
        
        button_panel.Children.Add(btn_select_all)
        button_panel.Children.Add(btn_deselect_all)
        button_panel.Children.Add(btn_execute)
        button_panel.Children.Add(btn_cancel)
        
        main_grid.Children.Add(scroll_viewer)
        main_grid.Children.Add(button_panel)
        
        self.window.Content = main_grid
    
    def populate_tree(self, node, items):
        """Заполнить дерево элементами"""
        from System.Windows.Controls import TreeViewItem, CheckBox
        from System.Windows import Thickness
        
        for child in node.Children:
            tvi = TreeViewItem()
            tvi.IsExpanded = False
            
            cb = CheckBox()
            cb.Content = child.Name
            cb.Margin = Thickness(2)
            cb.Tag = child
            
            cb.Checked += self.on_check_changed
            cb.Unchecked += self.on_check_changed
            
            tvi.Header = cb
            
            if len(child.Children) > 0:
                self.populate_tree(child, tvi.Items)
            
            items.Add(tvi)
    
    def on_check_changed(self, sender, e):
        """Обработчик изменения чекбокса"""
        from System.Windows.Controls import TreeViewItem
        
        node = sender.Tag
        node.IsChecked = sender.IsChecked
        
        parent = sender.Parent
        while parent and not isinstance(parent, TreeViewItem):
            parent = parent.Parent
        
        if parent:
            self.update_children_in_treeViewItem(parent, sender.IsChecked)
    
    def update_children_in_treeViewItem(self, tree_view_item, is_checked):
        """Обновить дочерние элементы"""
        from System.Windows.Controls import TreeViewItem, CheckBox
        
        for item in tree_view_item.Items:
            if isinstance(item, TreeViewItem) and isinstance(item.Header, CheckBox):
                cb = item.Header
                cb.IsChecked = is_checked
                cb.Tag.IsChecked = is_checked
                self.update_children_in_treeViewItem(item, is_checked)
    
    def on_select_all(self, sender, e):
        """Выделить все элементы"""
        self.set_all_checked(self.tree_view.Items, True)
    
    def on_deselect_all(self, sender, e):
        """Снять выделение"""
        self.set_all_checked(self.tree_view.Items, False)
    
    def set_all_checked(self, items, is_checked):
        """Установить состояние чекбоксов"""
        from System.Windows.Controls import TreeViewItem, CheckBox
        
        for item in items:
            if isinstance(item, TreeViewItem) and isinstance(item.Header, CheckBox):
                item.Header.IsChecked = is_checked
                item.Header.Tag.IsChecked = is_checked
            
            if hasattr(item, 'Items'):
                self.set_all_checked(item.Items, is_checked)
    
    def on_execute(self, sender, e):
        """Выполнить обработку"""
        self.SelectedPlans = self.collect_selected_plans(self.tree_view.Items)
        
        if not self.SelectedPlans:
            forms.alert("Не выбрано ни одного плана для обработки.", title="Предупреждение")
            return
        
        self.DialogResult = True
        self.window.Close()
    
    def on_cancel(self, sender, e):
        """Отмена"""
        self.DialogResult = False
        self.window.Close()
    
    def ShowDialog(self):
        """Показать окно"""
        return self.window.ShowDialog()
    
    def collect_selected_plans(self, items):
        """Собрать выбранные планы"""
        from System.Windows.Controls import TreeViewItem, CheckBox
        plans = []
        
        for item in items:
            if isinstance(item, TreeViewItem) and isinstance(item.Header, CheckBox):
                cb = item.Header
                node = cb.Tag
                
                if cb.IsChecked and node.LevelType == "plan" and node.Data:
                    plans.append(node.Data)
            
            if hasattr(item, 'Items'):
                plans.extend(self.collect_selected_plans(item.Items))
        
        return plans


# ==============================================================================
# ОСНОВНАЯ ЛОГИКА
# ==============================================================================

def process_elements(doc, selected_plans, progress_callback=None):
    """Обработать элементы на выбранных планах"""
    
    element_sheets = defaultdict(set)
    elements_empty_zone = []
    elements_empty_vidraboty = []
    elements_no_materials = []
    total_processed = 0
    
    for i, view_plan in enumerate(selected_plans):
        if progress_callback:
            progress_callback(i, len(selected_plans), "Обработка: {}".format(view_plan.Name))
        
        sheet_number = get_sheet_number_for_view(doc, view_plan)
        
        if not sheet_number:
            continue
        
        elements = get_elements_on_view(doc, view_plan)
        
        for element in elements:
            total_processed += 1
            
            vid_raboty = get_parameter_value(element, PARAM_VID_RABOTY)
            zona_raw = get_parameter_value(element, PARAM_ZONA)
            
            if not vid_raboty:
                elements_empty_vidraboty.append(element.Id)
                continue
            
            zona = determine_zone(zona_raw)
            
            if zona is None:
                elements_empty_zone.append((element.Id, zona_raw))
                continue
            
            category = element.Category
            is_material = is_material_category(category)
            
            if is_material:
                materials = get_materials_with_kod_edinicy(element, doc)
                
                if not materials:
                    elements_no_materials.append(element.Id)
                    continue
                
                element_sheets[element.Id].add(sheet_number)
            else:
                element_sheets[element.Id].add(sheet_number)
    
    return element_sheets, elements_empty_zone, elements_empty_vidraboty, elements_no_materials, total_processed


def write_values(doc, element_sheets, progress_callback=None):
    """Записать значения в параметры элементов"""
    
    updated_count = 0
    errors = []
    
    t = Transaction(doc, "Заполнение GP_01_СсылкаНаЛист")
    t.Start()
    
    try:
        for i, (element_id, sheet_numbers) in enumerate(element_sheets.items()):
            if progress_callback:
                progress_callback(i, len(element_sheets), "Запись значений...")
            
            sheet_ref = format_sheet_reference(sheet_numbers)
            
            element = doc.GetElement(element_id)
            if element:
                param = element.LookupParameter(PARAM_SSYLKA_NA_LIST)
                if param:
                    try:
                        param.Set(sheet_ref)
                        updated_count += 1
                    except Exception as ex:
                        errors.append("Element {}: {}".format(element_id, str(ex)))
        
        t.Commit()
    except Exception as ex:
        t.RollBack()
        raise ex
    
    return updated_count, errors


def show_report(total_processed, updated_count, elements_empty_zone, 
                elements_empty_vidraboty, elements_no_materials, errors, elapsed_time):
    """Показать отчёт о выполнении"""
    
    report_lines = [
        "ОТЧЁТ О ВЫПОЛНЕНИИ",
        "",
        "Обработано элементов: {}".format(total_processed),
        "Обновлено элементов: {}".format(updated_count),
        "Время выполнения: {:.2f} сек".format(elapsed_time),
    ]
    
    if elements_empty_zone:
        report_lines.append("")
        report_lines.append("Элементы без определённой зоны ({} шт.):".format(len(elements_empty_zone)))
        for elem_id, zona_val in elements_empty_zone[:10]:
            report_lines.append("  Id: {}, значение: '{}'".format(elem_id, zona_val or "пусто"))
        if len(elements_empty_zone) > 10:
            report_lines.append("  ... и ещё {} элементов".format(len(elements_empty_zone) - 10))
    
    if elements_empty_vidraboty:
        report_lines.append("")
        report_lines.append("Элементы без ВидРаботы ({} шт.) - пропущены".format(len(elements_empty_vidraboty)))
    
    if elements_no_materials:
        report_lines.append("")
        report_lines.append("Материалы без КодЕдиницыЭкз ({} шт.) - пропущены".format(len(elements_no_materials)))
    
    if errors:
        report_lines.append("")
        report_lines.append("Ошибки при записи ({} шт.):".format(len(errors)))
        for err in errors[:5]:
            report_lines.append("  {}".format(err))
    
    forms.alert("\n".join(report_lines), title="Результат", warn_icon=False)


# ==============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ==============================================================================

def main():
    """Главная функция скрипта"""
    
    doc = __revit__.ActiveUIDocument.Document
    
    if doc is None:
        forms.alert("Нет активного документа Revit.", title="Ошибка")
        return
    
    start_time = time.time()
    
    tree_root = build_sheet_tree(doc)
    
    if tree_root is None or len(tree_root.Children) == 0:
        forms.alert("В проекте не найдено листов с планами этажей.", title="Информация")
        return
    
    selection_window = SheetSelectionWindow(tree_root)
    selection_window.ShowDialog()
    
    if not selection_window.DialogResult:
        return
    
    selected_plans = selection_window.SelectedPlans
    
    if not selected_plans:
        forms.alert("Не выбрано ни одного плана.", title="Предупреждение")
        return
    
    with forms.ProgressBar(title="Обработка элементов...", cancellable=True) as pb:
        pb.max_value = len(selected_plans)
        
        def progress_callback(current, total, message):
            pb.update_progress(current + 1)
            pb.title = message
        
        element_sheets, elements_empty_zone, elements_empty_vidraboty, elements_no_materials, total_processed = \
            process_elements(doc, selected_plans, progress_callback)
    
    if not element_sheets:
        forms.alert("Не найдено элементов для обработки.\n\n"
                   "Проверьте:\n"
                   "- Заполнен ли GP_01_ВидРаботы у элементов\n"
                   "- Заполнен ли GP_01_Зона (Подземная/Надземная)\n"
                   "- Есть ли материалы с GP_01_КодЕдиницыЭкз", 
                   title="Информация")
        return
    
    with forms.ProgressBar(title="Запись значений...", cancellable=True) as pb:
        pb.max_value = len(element_sheets)
        
        def write_progress_callback(current, total, message):
            pb.update_progress(current + 1)
        
        updated_count, errors = write_values(doc, element_sheets, write_progress_callback)
    
    elapsed_time = time.time() - start_time
    
    show_report(
        total_processed=total_processed,
        updated_count=updated_count,
        elements_empty_zone=elements_empty_zone,
        elements_empty_vidraboty=elements_empty_vidraboty,
        elements_no_materials=elements_no_materials,
        errors=errors,
        elapsed_time=elapsed_time
    )


if __name__ == "__main__":
    main()
