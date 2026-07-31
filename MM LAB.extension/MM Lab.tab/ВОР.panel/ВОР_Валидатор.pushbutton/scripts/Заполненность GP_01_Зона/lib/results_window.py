# -*- coding: utf-8 -*-
"""
Плавающее окно результатов проверки GP_01_Зона.
Раскрываемые категории с кликабельными элементами.
"""

import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness, Visibility, GridLength, GridUnitType,
    FontWeights, HorizontalAlignment, VerticalAlignment
)
from System.Windows.Controls import (
    StackPanel, Grid, ColumnDefinition, TextBlock, ListBox,
    ListBoxItem, Button, Border, ScrollViewer, ScrollBarVisibility,
    Expander
)
from System.Windows.Media import SolidColorBrush, Colors

from pyrevit import DB

_GRAY = SolidColorBrush(Colors.Gray)
_WHITE = SolidColorBrush(Colors.White)
_BORDER_GRAY = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 200, 200, 200))
_RED = SolidColorBrush(Colors.Red)
_GREEN = SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 76, 175, 80))


class ResultsWindow(Window):

    def __init__(self, doc, results_data):
        """
        results_data: list of {
            "category": str,
            "checked": int,
            "problems": int,
            "elements": [{"id": ElementId, "family": str, "type": str}]
        }
        """
        self.Title = u"\u0420\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442\u044b: GP_01_\u0417\u043e\u043d\u0430"
        self.Width = 650
        self.Height = 520
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.CanResize
        self.Background = _WHITE
        self.doc = doc

        self._all_ids = []

        root = StackPanel()
        root.Margin = Thickness(16)

        # Заголовок
        header = TextBlock()
        header.Text = u"\u0420\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442\u044b \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438 GP_01_\u0417\u043e\u043d\u0430"
        header.FontSize = 14
        header.FontWeight = FontWeights.Bold
        header.Margin = Thickness(0, 0, 0, 4)
        root.Children.Add(header)

        hint = TextBlock()
        hint.Text = u"\u041a\u043b\u0438\u043a\u043d\u0438\u0442\u0435 \u043d\u0430 \u044d\u043b\u0435\u043c\u0435\u043d\u0442 \u0434\u043b\u044f \u0432\u044b\u0434\u0435\u043b\u0435\u043d\u0438\u044f \u0432 \u043c\u043e\u0434\u0435\u043b\u0438"
        hint.FontSize = 11
        hint.FontStyle = System.Windows.FontStyles.Italic
        hint.Foreground = _GRAY
        hint.Margin = Thickness(0, 0, 0, 8)
        root.Children.Add(hint)

        # Заголовок перед списком категорий
        list_header = TextBlock()
        list_header.Text = u"\u042d\u043b\u0435\u043c\u0435\u043d\u0442\u044b \u0441 \u043d\u0435\u0437\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u044b\u043c \u043f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u043e\u043c GP_01_\u0417\u043e\u043d\u0430:"
        list_header.FontSize = 12
        list_header.FontWeight = FontWeights.SemiBold
        list_header.Margin = Thickness(0, 0, 0, 6)
        root.Children.Add(list_header)

        # Прокручиваемая область категорий (только с проблемами)
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.MaxHeight = 340

        cats_panel = StackPanel()

        for rd in results_data:
            if rd["problems"] == 0 or not rd["elements"]:
                continue

            self._all_ids.extend([e["id"] for e in rd["elements"]])

            # Expander для категории
            expander = Expander()
            expander.IsExpanded = False
            expander.Margin = Thickness(0, 0, 0, 6)

            # Заголовок expander
            header_panel = StackPanel()
            header_panel.Orientation = System.Windows.Controls.Orientation.Horizontal

            count_text = TextBlock()
            count_text.Text = u"{}: {}".format(
                rd["category"], rd["problems"]
            )
            count_text.FontWeight = FontWeights.SemiBold
            count_text.FontSize = 13
            count_text.Foreground = _RED
            header_panel.Children.Add(count_text)

            expander.Header = header_panel

            # Содержимое — список элементов
            listbox = ListBox()
            listbox.FontSize = 12
            listbox.MaxHeight = 200
            listbox.Margin = Thickness(0, 4, 0, 0)
            listbox.BorderBrush = _BORDER_GRAY
            listbox.BorderThickness = Thickness(1)

            for elem_info in rd["elements"]:
                item = ListBoxItem()
                eid_int = elem_info["id"].IntegerValue
                fam = elem_info["family"]
                typ = elem_info["type"]
                pval = elem_info.get("param_value", "")
                item.Content = u"ID:{} | {} | {} | \u0417\u043d\u0430\u0447\u0435\u043d\u0438\u0435: \"{}\"".format(eid_int, fam, typ, pval)
                item.Tag = elem_info["id"]
                listbox.Items.Add(item)

            # SelectionChanged — выделение в модели
            def _on_select(sender, e):
                selected = sender.SelectedItem
                if selected and selected.Tag:
                    self._select_elements([selected.Tag])

            listbox.SelectionChanged += _on_select

            expander.Content = listbox
            cats_panel.Children.Add(expander)

        scroll.Content = cats_panel
        root.Children.Add(scroll)

        # Кнопки
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 8, 0, 0)

        btn_select = Button()
        btn_select.Content = u"\u0412\u044b\u0434\u0435\u043b\u0438\u0442\u044c \u0432\u0441\u0435"
        btn_select.Width = 120
        btn_select.Height = 28
        btn_select.Margin = Thickness(0, 0, 8, 0)
        btn_select.Click += self._on_select_all
        btn_panel.Children.Add(btn_select)

        btn_close = Button()
        btn_close.Content = u"\u0417\u0430\u043a\u0440\u044b\u0442\u044c"
        btn_close.Width = 100
        btn_close.Height = 28
        btn_close.Click += self._on_close
        btn_panel.Children.Add(btn_close)

        root.Children.Add(btn_panel)
        self.Content = root

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
        if self._all_ids:
            self._select_elements(self._all_ids)

    def _on_close(self, sender, e):
        self.Close()
