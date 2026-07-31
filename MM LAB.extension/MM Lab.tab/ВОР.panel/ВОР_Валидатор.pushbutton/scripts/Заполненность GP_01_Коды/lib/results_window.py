# -*- coding: utf-8 -*-
"""
Плавающее окно результатов проверки GP_01_Коды.
4 поля (2x2): Семейства/Материалы x Проблемы/Исключения.
"""

import clr
import os
import sys

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness, Visibility, GridLength, GridUnitType,
    FontWeights, HorizontalAlignment, VerticalAlignment
)
from System.Windows.Controls import (
    StackPanel, Grid, ColumnDefinition, RowDefinition, TextBlock,
    ListBox, ListBoxItem, Button, Border, ScrollViewer,
    ScrollBarVisibility, DockPanel
)
from System.Windows.Media import SolidColorBrush, Colors

from pyrevit import DB

_WHITE = SolidColorBrush(Colors.White)
_GRAY = SolidColorBrush(Colors.Gray)
_RED = SolidColorBrush(Colors.Red)
_ORANGE = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 152, 0))
_GREEN = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 76, 175, 80))
_RED_BG = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 235, 238))
_GREEN_BG = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 232, 245, 233))
_HEADER_PROBLEM = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 211, 47, 47))
_HEADER_EXCEPTION = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 56, 142, 60))
_BORDER_GRAY = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 200, 200, 200))
_BTN_ORANGE_BG = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 167, 38))
_BTN_GREEN_BG = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 102, 187, 106))
_WHITE_BRUSH = SolidColorBrush(Colors.White)


def _get_marker_text(info):
    """Сформировать текст пометки для family-элемента."""
    marker = info.get("empty_marker", "")
    kod_ed = info.get("kod_ed", "")
    kod_vr = info.get("kod_vr", "")
    if marker == "both":
        return u"\u26A0 \u043E\u0431\u0430 \u043F\u0430\u0440\u0430\u043C\u0435\u0442\u0440\u0430 \u043F\u0443\u0441\u0442\u044B"
    elif marker == "kod_ed":
        return u"\u2717 \u041A\u043E\u0434\u0415\u0434\u0438\u043D\u0438\u0446\u044B: \u043F\u0443\u0441\u0442\u043E | \u2713 \u041A\u043E\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043E\u0442\u044B: \"{}\"".format(kod_vr)
    elif marker == "kod_vr":
        return u"\u2713 \u041A\u043E\u0434\u0415\u0434\u0438\u043D\u0438\u0446\u044B: \"{}\" | \u2717 \u041A\u043E\u0434\u0412\u0438\u0434\u0430\u0420\u0430\u0431\u043E\u0442\u044B: \u043F\u0443\u0441\u0442\u043E".format(kod_ed)
    return u""


class KodyResultsWindow(Window):

    def __init__(self, doc, results_data):
        self.Title = u"\u0420\u0435\u0437\u0443\u043B\u044C\u0442\u0430\u0442\u044B: GP_01_\u041A\u043E\u0434\u044B"
        self.Width = 900
        self.Height = 620
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.CanResize
        self.Background = _WHITE
        self.doc = doc

        self._results_data = results_data
        self._script_name = results_data["script_name"]
        self._section = results_data["section"]
        self._project = results_data["project"]

        self._excluded_uids = {
            "family": list(results_data["excluded_uids"]["family"]),
            "material": list(results_data["excluded_uids"]["material"]),
        }

        self._all_problem_ids = []

        self._build_ui()

    def _build_ui(self):
        """Построить полный UI окна."""
        self._all_problem_ids = []

        family_problems = self._results_data["family"]["problems"]
        family_exceptions = self._results_data["family"]["exceptions"]
        material_problems = self._results_data["material"]["problems"]
        material_exceptions = self._results_data["material"]["exceptions"]

        self._all_problem_ids = (
            [e["id"] for e in family_problems]
            + [e["id"] for e in material_problems]
        )

        root = DockPanel()
        root.Margin = Thickness(12)

        # === Заголовок ===
        top_panel = StackPanel()
        top_panel.Margin = Thickness(0, 0, 0, 8)
        DockPanel.SetDock(top_panel, System.Windows.Controls.Dock.Top)

        header = TextBlock()
        header.Text = u"\u0420\u0435\u0437\u0443\u043B\u044C\u0442\u0430\u0442\u044B \u043F\u0440\u043E\u0432\u0435\u0440\u043A\u0438 GP_01_\u041A\u043E\u0434\u044B"
        header.FontSize = 14
        header.FontWeight = FontWeights.Bold
        header.Margin = Thickness(0, 0, 0, 4)
        top_panel.Children.Add(header)

        hint = TextBlock()
        hint.Text = u"\u041A\u043B\u0438\u043A\u043D\u0438\u0442\u0435 \u043D\u0430 \u044D\u043B\u0435\u043C\u0435\u043D\u0442 \u0434\u043B\u044F \u0432\u044B\u0434\u0435\u043B\u0435\u043D\u0438\u044F \u0432 \u043C\u043E\u0434\u0435\u043B\u0438"
        hint.FontSize = 11
        hint.FontStyle = System.Windows.FontStyles.Italic
        hint.Foreground = _GRAY
        hint.Margin = Thickness(0, 0, 0, 4)
        top_panel.Children.Add(hint)

        fp_count = len(family_problems)
        fe_count = len(family_exceptions)
        mp_count = len(material_problems)
        me_count = len(material_exceptions)

        summary = TextBlock()
        summary.Text = (
            u"\u0421\u0435\u043C\u0435\u0439\u0441\u0442\u0432\u0430: {} \u043F\u0440. / {} \u0438\u0441\u043A. | "
            u"\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B: {} \u043F\u0440. / {} \u0438\u0441\u043A."
        ).format(fp_count, fe_count, mp_count, me_count)
        summary.FontSize = 12
        summary.FontWeight = FontWeights.SemiBold
        top_panel.Children.Add(summary)

        root.Children.Add(top_panel)

        # === Сетка 2x2 ===
        main_grid = Grid()
        main_grid.Margin = Thickness(0, 0, 0, 8)

        col1 = ColumnDefinition()
        col1.Width = GridLength(1, GridUnitType.Star)
        col2 = ColumnDefinition()
        col2.Width = GridLength(1, GridUnitType.Star)
        main_grid.ColumnDefinitions.Add(col1)
        main_grid.ColumnDefinitions.Add(col2)

        row1 = RowDefinition()
        row1.Height = GridLength(1, GridUnitType.Star)
        row2 = RowDefinition()
        row2.Height = GridLength(1, GridUnitType.Star)
        main_grid.RowDefinitions.Add(row1)
        main_grid.RowDefinitions.Add(row2)

        # [0,0] Семейства — Проблемы
        cell_00 = self._build_cell(
            title=u"\u0421\u0415\u041C\u0415\u0419\u0421\u0422\u0412\u0410 \u2014 \u041F\u0440\u043E\u0431\u043B\u0435\u043C\u044B ({})".format(fp_count),
            header_bg=_HEADER_PROBLEM,
            cell_bg=_RED_BG,
            elements=family_problems,
            is_exception=False,
            group="family",
        )
        Grid.SetColumn(cell_00, 0)
        Grid.SetRow(cell_00, 0)
        main_grid.Children.Add(cell_00)

        # [1,0] Материалы — Проблемы
        cell_10 = self._build_cell(
            title=u"\u041C\u0410\u0422\u0415\u0420\u0418\u0410\u041B\u042B \u2014 \u041F\u0440\u043E\u0431\u043B\u0435\u043C\u044B ({})".format(mp_count),
            header_bg=_HEADER_PROBLEM,
            cell_bg=_RED_BG,
            elements=material_problems,
            is_exception=False,
            group="material",
        )
        Grid.SetColumn(cell_10, 1)
        Grid.SetRow(cell_10, 0)
        main_grid.Children.Add(cell_10)

        # [0,1] Семейства — Исключения
        cell_01 = self._build_cell(
            title=u"\u0421\u0415\u041C\u0415\u0419\u0421\u0422\u0412\u0410 \u2014 \u0418\u0441\u043A\u043B\u044E\u0447\u0435\u043D\u0438\u044F ({})".format(fe_count),
            header_bg=_HEADER_EXCEPTION,
            cell_bg=_GREEN_BG,
            elements=family_exceptions,
            is_exception=True,
            group="family",
        )
        Grid.SetColumn(cell_01, 0)
        Grid.SetRow(cell_01, 1)
        main_grid.Children.Add(cell_01)

        # [1,1] Материалы — Исключения
        cell_11 = self._build_cell(
            title=u"\u041C\u0410\u0422\u0415\u0420\u0418\u0410\u041B\u042B \u2014 \u0418\u0441\u043A\u043B\u044E\u0447\u0435\u043D\u0438\u044F ({})".format(me_count),
            header_bg=_HEADER_EXCEPTION,
            cell_bg=_GREEN_BG,
            elements=material_exceptions,
            is_exception=True,
            group="material",
        )
        Grid.SetColumn(cell_11, 1)
        Grid.SetRow(cell_11, 1)
        main_grid.Children.Add(cell_11)

        root.Children.Add(main_grid)

        # === Кнопки ===
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        DockPanel.SetDock(btn_panel, System.Windows.Controls.Dock.Bottom)

        btn_select = Button()
        btn_select.Content = u"\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u044C \u0432\u0441\u0435 \u043F\u0440\u043E\u0431\u043B\u0435\u043C\u044B"
        btn_select.Width = 160
        btn_select.Height = 28
        btn_select.Margin = Thickness(0, 0, 8, 0)
        btn_select.Click += self._on_select_all
        btn_panel.Children.Add(btn_select)

        btn_close = Button()
        btn_close.Content = u"\u0417\u0430\u043A\u0440\u044B\u0442\u044C"
        btn_close.Width = 100
        btn_close.Height = 28
        btn_close.Click += self._on_close
        btn_panel.Children.Add(btn_close)

        root.Children.Add(btn_panel)
        self.Content = root

    def _build_cell(self, title, header_bg, cell_bg, elements, is_exception, group):
        """Построить одну ячейку сетки."""
        border = Border()
        border.BorderBrush = _BORDER_GRAY
        border.BorderThickness = Thickness(1)
        border.Background = cell_bg
        border.Margin = Thickness(2)

        cell = DockPanel()

        # Заголовок ячейки
        header_border = Border()
        header_border.Background = header_bg
        header_border.Padding = Thickness(8, 4, 8, 4)
        DockPanel.SetDock(header_border, System.Windows.Controls.Dock.Top)

        header_tb = TextBlock()
        header_tb.Text = title
        header_tb.FontSize = 12
        header_tb.FontWeight = FontWeights.SemiBold
        header_tb.Foreground = _WHITE_BRUSH
        header_border.Child = header_tb
        cell.Children.Add(header_border)

        # Список элементов
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Margin = Thickness(4)

        listbox = ListBox()
        listbox.FontSize = 11
        listbox.BorderThickness = Thickness(0)
        listbox.Background = SolidColorBrush(Colors.Transparent)

        for info in elements:
            item = self._build_list_item(info, is_exception, group)
            listbox.Items.Add(item)

        # SelectionChanged — выделение в модели
        def _on_select(sender, e):
            selected = sender.SelectedItem
            if selected and hasattr(selected, 'Tag') and selected.Tag:
                self._select_elements([selected.Tag])

        listbox.SelectionChanged += _on_select

        scroll.Content = listbox
        cell.Children.Add(scroll)

        if not elements:
            empty_tb = TextBlock()
            if is_exception:
                empty_tb.Text = u"\u041D\u0435\u0442 \u0438\u0441\u043A\u043B\u044E\u0447\u0435\u043D\u0438\u0439"
            else:
                empty_tb.Text = u"\u041D\u0435\u0442 \u043F\u0440\u043E\u0431\u043B\u0435\u043C"
            empty_tb.FontSize = 11
            empty_tb.Foreground = _GRAY
            empty_tb.HorizontalAlignment = HorizontalAlignment.Center
            empty_tb.Margin = Thickness(0, 12, 0, 0)
            cell.Children.Add(empty_tb)

        border.Child = cell
        return border

    def _build_list_item(self, info, is_exception, group):
        """Построить строку элемента с текстом и кнопкой."""
        item_grid = Grid()
        item_grid.Margin = Thickness(0, 2, 0, 2)

        col_text = ColumnDefinition()
        col_text.Width = GridLength(1, GridUnitType.Star)
        col_btn = ColumnDefinition()
        col_btn.Width = GridLength(80, GridUnitType.Pixel)
        item_grid.ColumnDefinitions.Add(col_text)
        item_grid.ColumnDefinitions.Add(col_btn)

        # Текст элемента
        text_panel = StackPanel()
        Grid.SetColumn(text_panel, 0)

        eid_int = info["id"].IntegerValue
        family = info.get("family", "")
        typ = info.get("type", "")

        line1 = TextBlock()
        line1.Text = u"ID:{} | {} | {}".format(eid_int, family, typ)
        text_panel.Children.Add(line1)

        if group == "family":
            marker_text = _get_marker_text(info)
            if marker_text:
                line2 = TextBlock()
                line2.Text = u"  {}".format(marker_text)
                line2.FontSize = 10
                marker = info.get("empty_marker", "")
                if marker == "both":
                    line2.Foreground = _RED
                else:
                    line2.Foreground = _ORANGE
                text_panel.Children.Add(line2)
        else:
            kod_mat = info.get("kod_mat", "")
            line2 = TextBlock()
            line2.Text = u"  \u041A\u043E\u0434\u0415\u0434\u0438\u043D\u0438\u0446\u044B_\u041C\u0430\u0442: \"{}\"".format(kod_mat)
            line2.FontSize = 10
            line2.Foreground = _ORANGE
            text_panel.Children.Add(line2)

        item_grid.Children.Add(text_panel)

        # Кнопка
        btn = Button()
        btn.Width = 75
        btn.Height = 22
        btn.FontSize = 10
        btn.Tag = info.get("unique_id", "")

        if is_exception:
            btn.Content = u"\u0412\u0435\u0440\u043D\u0443\u0442\u044C"
            btn.Background = _BTN_GREEN_BG
            btn.Foreground = _WHITE_BRUSH
            btn.Click += self._on_return
        else:
            btn.Content = u"\u0418\u0441\u043A\u043B\u044E\u0447\u0438\u0442\u044C"
            btn.Background = _BTN_ORANGE_BG
            btn.Foreground = _WHITE_BRUSH
            btn.Click += self._on_exclude

        btn.Margin = Thickness(4, 0, 0, 0)
        btn.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(btn, 1)
        item_grid.Children.Add(btn)

        item = ListBoxItem()
        item.Content = item_grid
        item.Tag = info["id"]
        return item

    # ================================================================
    # Actions
    # ================================================================

    def _select_elements(self, elem_ids):
        from pyrevit import revit
        from System.Collections.Generic import List

        try:
            uidoc = revit.uidoc
            ids = List[DB.ElementId]()
            for eid in elem_ids:
                ids.Add(eid)
            uidoc.Selection.SetElementIds(ids)
        except Exception:
            pass

    def _on_select_all(self, sender, e):
        if self._all_problem_ids:
            self._select_elements(self._all_problem_ids)

    def _on_exclude(self, sender, e):
        unique_id = sender.Tag
        if not unique_id:
            return

        # Определяем группу по расположению кнопки
        group = self._find_group_for_uid(unique_id)
        if group and unique_id not in self._excluded_uids[group]:
            self._excluded_uids[group].append(unique_id)
            self._persist_exclusions()
            self._reclassify_and_refresh()

    def _on_return(self, sender, e):
        unique_id = sender.Tag
        if not unique_id:
            return

        for group in ("family", "material"):
            if unique_id in self._excluded_uids[group]:
                self._excluded_uids[group].remove(unique_id)
                self._persist_exclusions()
                self._reclassify_and_refresh()
                return

    def _find_group_for_uid(self, unique_id):
        """Определить группу элемента по UniqueId."""
        for group in ("family", "material"):
            for info in self._results_data[group]["problems"]:
                if info.get("unique_id") == unique_id:
                    return group
            for info in self._results_data[group]["exceptions"]:
                if info.get("unique_id") == unique_id:
                    return group
        return None

    def _reclassify_and_refresh(self):
        """Пересчитать проблемы/исключения и перестроить UI."""
        new_problems_family = []
        new_exceptions_family = []
        new_problems_material = []
        new_exceptions_material = []

        excluded_fam = set(self._excluded_uids["family"])
        excluded_mat = set(self._excluded_uids["material"])

        for info in self._results_data["family"]["problems"]:
            if info["unique_id"] in excluded_fam:
                new_exceptions_family.append(info)
            else:
                new_problems_family.append(info)

        for info in self._results_data["family"]["exceptions"]:
            if info["unique_id"] in excluded_fam:
                new_exceptions_family.append(info)
            else:
                new_problems_family.append(info)

        for info in self._results_data["material"]["problems"]:
            if info["unique_id"] in excluded_mat:
                new_exceptions_material.append(info)
            else:
                new_problems_material.append(info)

        for info in self._results_data["material"]["exceptions"]:
            if info["unique_id"] in excluded_mat:
                new_exceptions_material.append(info)
            else:
                new_problems_material.append(info)

        self._results_data["family"]["problems"] = new_problems_family
        self._results_data["family"]["exceptions"] = new_exceptions_family
        self._results_data["material"]["problems"] = new_problems_material
        self._results_data["material"]["exceptions"] = new_exceptions_material

        self.Content = None
        self._build_ui()

    def _persist_exclusions(self):
        """Сохранить исключения в файл настроек."""
        bundle_path = os.path.dirname(os.path.dirname(os.path.dirname(
            os.path.dirname(os.path.abspath(__file__)))))
        if bundle_path not in sys.path:
            sys.path.insert(0, bundle_path)

        try:
            from core.config_manager import load_script_settings, save_script_settings

            current = load_script_settings(
                self._script_name, self._section, self._project
            )
            current["excluded_family"] = self._excluded_uids["family"]
            current["excluded_material"] = self._excluded_uids["material"]
            save_script_settings(
                self._script_name, self._section, self._project, current
            )
        except Exception:
            pass

    def _on_close(self, sender, e):
        self.Close()
