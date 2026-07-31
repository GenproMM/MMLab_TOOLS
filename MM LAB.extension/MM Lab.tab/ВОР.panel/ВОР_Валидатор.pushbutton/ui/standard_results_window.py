# -*- coding: utf-8 -*-
"""
Стандартное results-окно для скриптов БЕЗ своего show_results.

Используется окном прогона (run_window.py), когда у скрипта нет
кастомного окна результатов: показывает список проблемных элементов
с кликом-и-выделением в Revit.

Контракт данных (results):
    {"category_name": [{"id": ElementId, "name": str}, ...], ...}

Это адаптированная версия шаблона
.agents/skills/VorUICreator/examples/results-window/results_window.py
(исправлена опечатка SetElement_ids -> SetElementIds).
"""

import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness, FontWeights, FontStyles,
    HorizontalAlignment, VerticalAlignment,
    WindowStartupLocation, ResizeMode
)
from System.Windows.Controls import (
    StackPanel, TextBlock, ListBox, ListBoxItem, Button, Border,
    ScrollViewer, ScrollBarVisibility, Expander, Orientation, DockPanel
)
from System.Windows.Media import SolidColorBrush, Colors, Color
from System.Windows.Markup import XamlReader

from pyrevit import DB

# ── Design System (совпадает с шаблоном VorUICreator) ──────────
_BG_WINDOW = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))
_BRUSH_BORDER = SolidColorBrush(Color.FromArgb(255, 204, 204, 204))

_FG_TITLE = SolidColorBrush(Color.FromArgb(255, 51, 51, 51))
_FG_MUTED = SolidColorBrush(Colors.Gray)
_FG_ERROR = SolidColorBrush(Color.FromArgb(255, 211, 47, 47))
_FG_SUCCESS = SolidColorBrush(Color.FromArgb(255, 56, 142, 60))

_SECONDARY_BUTTON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" Background="#E0E0E0" BorderBrush="#CCCCCC" '
    'BorderThickness="1" CornerRadius="3" Padding="10,5">'
    '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsMouseOver" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#D0D0D0"/>'
    '</Trigger>'
    '<Trigger Property="IsPressed" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#C0C0C0"/>'
    '</Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate>'
)


def _make_title(text, font_size=14):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = font_size
    tb.FontWeight = FontWeights.Bold
    tb.Foreground = _FG_TITLE
    tb.Margin = Thickness(0, 0, 0, 4)
    return tb


def _make_hint(text):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = 11
    tb.FontStyle = FontStyles.Italic
    tb.Foreground = _FG_MUTED
    tb.Margin = Thickness(0, 0, 0, 8)
    return tb


def _make_secondary_button(text, handler=None, width=100, height=30):
    btn = Button()
    btn.Content = text
    btn.Width = width
    btn.Height = height
    btn.Template = _SECONDARY_BUTTON_TEMPLATE
    if handler:
        btn.Click += handler
    return btn


class StandardResultsWindow(Window):
    """Modeless results-окно: Expander на категорию, клик выделяет в Revit."""

    def __init__(self, doc, results, title=None):
        """
        results: {"category_name": [{"id": ElementId, "name": str}, ...], ...}
        """
        self.doc = doc
        self._results = results
        self._all_ids = []

        self.Title = title or u"\u0420\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442\u044b \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438"
        self.Width = 650
        self.Height = 520
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize
        self.Background = _BG_WINDOW

        self._build_ui()

    def _build_ui(self):
        self._all_ids = []

        root = DockPanel()
        root.Margin = Thickness(12)

        # Header
        top = StackPanel()
        DockPanel.SetDock(top, System.Windows.Controls.Dock.Top)
        top.Children.Add(_make_title(u"\u0420\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442\u044b \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438"))
        top.Children.Add(_make_hint(
            u"\u041a\u043b\u0438\u043a\u043d\u0438\u0442\u0435 \u043d\u0430 \u044d\u043b\u0435\u043c\u0435\u043d\u0442 \u0434\u043b\u044f \u0432\u044b\u0434\u0435\u043b\u0435\u043d\u0438\u044f \u0432 \u043c\u043e\u0434\u0435\u043b\u0438"
        ))
        root.Children.Add(top)

        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        DockPanel.SetDock(btn_panel, System.Windows.Controls.Dock.Bottom)
        btn_panel.Margin = Thickness(0, 8, 0, 0)

        btn_all = _make_secondary_button(
            u"\u0412\u044b\u0434\u0435\u043b\u0438\u0442\u044c \u0432\u0441\u0435",
            self._on_select_all,
            width=120,
        )
        btn_all.Margin = Thickness(0, 0, 8, 0)
        btn_panel.Children.Add(btn_all)

        btn_close = _make_secondary_button(u"\u0417\u0430\u043a\u0440\u044b\u0442\u044c", self._on_close)
        btn_panel.Children.Add(btn_close)
        root.Children.Add(btn_panel)

        # Content
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Margin = Thickness(0, 4, 0, 0)

        content = StackPanel()

        for cat_name, elements in self._results.items():
            count = len(elements)
            for el in elements:
                self._all_ids.append(el["id"])

            expander = Expander()
            expander.Margin = Thickness(0, 0, 0, 8)

            header = TextBlock()
            header.Text = u"{} ({})".format(cat_name, count)
            header.FontWeight = FontWeights.SemiBold
            header.Foreground = _FG_ERROR if count > 0 else _FG_SUCCESS
            expander.Header = header
            expander.IsExpanded = count > 0 and count <= 20

            if not elements:
                empty = TextBlock()
                empty.Text = u"\u041d\u0435\u0442 \u043f\u0440\u043e\u0431\u043b\u0435\u043c"
                empty.FontSize = 12
                empty.Foreground = _FG_MUTED
                empty.Margin = Thickness(8, 4, 0, 4)
                expander.Content = empty
            else:
                lb = ListBox()
                lb.FontSize = 12
                lb.BorderThickness = Thickness(0)
                lb.MaxHeight = 240

                for el in elements:
                    item = ListBoxItem()
                    item.Content = el["name"]
                    item.Tag = el["id"]
                    lb.Items.Add(item)

                def _on_select(sender, e):
                    sel = sender.SelectedItem
                    if sel and hasattr(sel, "Tag") and sel.Tag:
                        self._select_elements([sel.Tag])

                lb.SelectionChanged += _on_select
                expander.Content = lb

            content.Children.Add(expander)

        scroll.Content = content
        root.Children.Add(scroll)
        self.Content = root

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
