# -*- coding: utf-8 -*-
"""
Диалог выбора спецификаций Revit из документа (множественный, с чекбоксами).
"""

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System
from System.Windows import (
    Window, WindowStartupLocation, Thickness,
    FontWeights, FontStyles, HorizontalAlignment, Visibility,
)
from System.Windows.Controls import (
    TextBox, CheckBox, Button, StackPanel,
    TextBlock, ScrollViewer, Border, ScrollBarVisibility,
)
from System.Windows.Media import SolidColorBrush, Colors, Color
from System.Windows.Markup import XamlReader

from pyrevit import revit

from schedule_reader import get_all_schedules

# ── Design system ──
_BG_WINDOW = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))
_BRUSH_BORDER = SolidColorBrush(Color.FromArgb(255, 204, 204, 204))

_PRIMARY_BUTTON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" Background="#007ACC" CornerRadius="4" Padding="12,6">'
    '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsMouseOver" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#005A9E"/>'
    '</Trigger>'
    '<Trigger Property="IsPressed" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#004578"/>'
    '</Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate>'
)

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


class SchedulePickerWindow(object):
    """Модальный диалог для выбора спецификаций с чекбоксами."""

    def __init__(self, doc, selected_names=None):
        self.doc = doc
        self.results = []
        self._selected_names = set(selected_names or [])

        self._build_ui()
        self._populate_schedules("")
        self.window.ShowDialog()

    def _build_ui(self):
        self.window = Window()
        self.window.Title = u"\u0412\u044b\u0431\u043e\u0440 \u0441\u043f\u0435\u0446\u0438\u0444\u0438\u043a\u0430\u0446\u0438\u0439"
        self.window.Width = 420
        self.window.Height = 520
        self.window.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.window.Background = _BG_WINDOW
        self.window.ResizeMode = System.Windows.ResizeMode.NoResize

        root = StackPanel()
        root.Margin = Thickness(16)

        # Заголовок
        title = TextBlock()
        title.Text = u"\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u0441\u043f\u0435\u0446\u0438\u0444\u0438\u043a\u0430\u0446\u0438\u0438"
        title.FontSize = 16
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 12)
        root.Children.Add(title)

        # Поиск
        search_label = TextBlock()
        search_label.Text = u"\u041f\u043e\u0438\u0441\u043a:"
        search_label.FontSize = 13
        search_label.FontWeight = FontWeights.SemiBold
        search_label.Margin = Thickness(0, 0, 0, 4)
        root.Children.Add(search_label)

        self.txt_search = TextBox()
        self.txt_search.Height = 30
        self.txt_search.FontSize = 13
        self.txt_search.VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        self.txt_search.Margin = Thickness(0, 0, 0, 8)
        root.Children.Add(self.txt_search)

        # Контейнер с чекбоксами в ScrollViewer
        border = Border()
        border.BorderBrush = _BRUSH_BORDER
        border.BorderThickness = Thickness(1)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.MaxHeight = 300

        self.checkbox_container = StackPanel()
        self.checkbox_container.Margin = Thickness(4)
        scroll.Content = self.checkbox_container
        border.Child = scroll
        root.Children.Add(border)

        # Кнопки
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 12, 0, 0)

        self.btn_cancel = Button()
        self.btn_cancel.Content = u"\u041e\u0442\u043c\u0435\u043d\u0430"
        self.btn_cancel.Width = 100
        self.btn_cancel.Height = 32
        self.btn_cancel.Margin = Thickness(8, 0, 0, 0)
        self.btn_cancel.Template = _SECONDARY_BUTTON_TEMPLATE
        self.btn_cancel.IsCancel = True
        btn_panel.Children.Add(self.btn_cancel)

        self.btn_select = Button()
        self.btn_select.Content = u"\u0412\u044b\u0431\u0440\u0430\u0442\u044c"
        self.btn_select.Width = 100
        self.btn_select.Height = 32
        self.btn_select.Margin = Thickness(8, 0, 0, 0)
        self.btn_select.Foreground = SolidColorBrush(Colors.White)
        self.btn_select.FontWeight = FontWeights.SemiBold
        self.btn_select.Template = _PRIMARY_BUTTON_TEMPLATE
        self.btn_select.IsDefault = True
        btn_panel.Children.Add(self.btn_select)

        root.Children.Add(btn_panel)

        self.window.Content = root

        # Events
        self.txt_search.TextChanged += self._on_search
        self.btn_select.Click += self._on_select
        self.btn_cancel.Click += self._on_cancel

    def _populate_schedules(self, filter_text):
        self.checkbox_container.Children.Clear()
        schedules = get_all_schedules(self.doc)
        filter_lower = filter_text.lower()

        empty = True
        for name, sched in schedules:
            if filter_lower and filter_lower not in name.lower():
                continue
            cb = CheckBox()
            cb.Content = name
            cb.Tag = sched
            cb.FontSize = 13
            cb.Margin = Thickness(2, 3, 2, 3)
            if name in self._selected_names:
                cb.IsChecked = True
            self.checkbox_container.Children.Add(cb)
            empty = False

        if empty:
            empty_tb = TextBlock()
            empty_tb.Text = u"\u041d\u0435\u0442 \u0441\u043f\u0435\u0446\u0438\u0444\u0438\u043a\u0430\u0446\u0438\u0439"
            empty_tb.FontStyle = FontStyles.Italic
            empty_tb.Foreground = SolidColorBrush(Colors.Gray)
            empty_tb.Margin = Thickness(2, 6, 2, 6)
            self.checkbox_container.Children.Add(empty_tb)

    def _on_search(self, sender, e):
        self._populate_schedules(self.txt_search.Text)

    def _on_select(self, sender, e):
        self._confirm_selection()

    def _on_cancel(self, sender, e):
        self.results = []
        self.window.Close()

    def _confirm_selection(self):
        self.results = []
        for child in self.checkbox_container.Children:
            if hasattr(child, "IsChecked") and child.IsChecked:
                self.results.append({
                    "name": child.Content,
                    "schedule": child.Tag,
                })
        setattr(self.window, "DialogResult", True)
        self.window.Close()
