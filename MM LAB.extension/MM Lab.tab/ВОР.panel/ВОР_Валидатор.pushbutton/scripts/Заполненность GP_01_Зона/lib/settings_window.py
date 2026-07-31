# -*- coding: utf-8 -*-
"""
Окно настроек для скрипта проверки заполненности GP_01_Зона.
Список категорий с чекбоксами, hide-unchecked, resizable.
"""

import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness, FontStyles, CornerRadius,
    FontWeights, HorizontalAlignment, VerticalAlignment
)
from System.Windows.Controls import (
    StackPanel, TextBlock, CheckBox, Button, Border, ScrollViewer,
    ScrollBarVisibility
)
from System.Windows.Media import SolidColorBrush, Colors, Color
from System.Windows.Markup import XamlReader

from pyrevit import DB

# ── Design System ──────────────────────────────────────────────
_BG_WINDOW = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))
_BG_CARD = SolidColorBrush(Colors.White)
_BG_PANEL = SolidColorBrush(Color.FromArgb(255, 250, 250, 250))
_BRUSH_BORDER = SolidColorBrush(Color.FromArgb(255, 204, 204, 204))

_BG_PRIMARY = SolidColorBrush(Color.FromArgb(255, 0, 122, 204))
_BG_PRIMARY_HOVER = SolidColorBrush(Color.FromArgb(255, 0, 90, 158))
_FG_ON_PRIMARY = SolidColorBrush(Colors.White)

_BG_SECONDARY = SolidColorBrush(Color.FromArgb(255, 224, 224, 224))
_BG_SECONDARY_HOVER = SolidColorBrush(Color.FromArgb(255, 208, 208, 208))
_BRUSH_SECONDARY_BORDER = SolidColorBrush(Color.FromArgb(255, 204, 204, 204))

_FG_TITLE = SolidColorBrush(Color.FromArgb(255, 51, 51, 51))
_FG_SUBTITLE = SolidColorBrush(Color.FromArgb(255, 85, 85, 85))
_FG_MUTED = SolidColorBrush(Colors.Gray)

# ── Button Templates ───────────────────────────────────────────
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

# ── Category Detection ─────────────────────────────────────────

PARAM_NAME = u"GP_01_\u0417\u043e\u043d\u0430"


def _detect_categories(doc):
    """Найти категории с параметром PARAM_NAME. Возвращает [(name, bic_int)]."""
    result = []
    try:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            definition = iterator.Key
            if definition.Name == PARAM_NAME:
                binding = iterator.Current
                cat_set = binding.Categories
                cat_iter = cat_set.GetEnumerator()
                while cat_iter.MoveNext():
                    cat = cat_iter.Current
                    bic_int = int(cat.BuiltInCategory)
                    result.append((cat.Name, bic_int))
                break
    except Exception:
        pass
    result.sort(key=lambda x: x[0])
    return result


# ── Window ─────────────────────────────────────────────────────

class ZoneSettingsWindow(Window):

    def __init__(self, doc, current_settings):
        self.Title = u"\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438: \u0417\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u044c GP_01_\u0417\u043e\u043d\u0430"
        self.Width = 420
        self.Height = 440
        self.MinWidth = 350
        self.MinHeight = 300
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.CanResize
        self.Background = _BG_WINDOW
        self.doc = doc
        self.result_settings = None

        self._checks = {}
        saved_categories = (current_settings or {}).get("categories", {})
        detected = _detect_categories(doc)

        # ── UI: StackPanel root ──
        root = StackPanel()
        root.Margin = Thickness(16)

        # Title
        title = TextBlock()
        title.Text = u"\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 \u0437\u0430\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u043e\u0441\u0442\u0438 GP_01_\u0417\u043e\u043d\u0430"
        title.FontSize = 16
        title.FontWeight = FontWeights.Bold
        title.Foreground = _FG_TITLE
        title.Margin = Thickness(0, 0, 0, 8)
        root.Children.Add(title)

        # Hint
        hint = TextBlock()
        hint.Text = u"\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 \u0434\u043b\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438"
        hint.FontSize = 11
        hint.FontStyle = FontStyles.Italic
        hint.Foreground = _FG_MUTED
        hint.Margin = Thickness(0, 0, 0, 8)
        root.Children.Add(hint)

        # Toggle buttons row (secondary style)
        toggle_panel = StackPanel()
        toggle_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        toggle_panel.Margin = Thickness(0, 0, 0, 8)

        btn_all = Button()
        btn_all.Content = u"\u0412\u044b\u0431\u0440\u0430\u0442\u044c \u0432\u0441\u0435"
        btn_all.FontSize = 11
        btn_all.Height = 26
        btn_all.Margin = Thickness(0, 0, 8, 0)
        btn_all.Template = _SECONDARY_BUTTON_TEMPLATE
        btn_all.Click += self._on_select_all
        toggle_panel.Children.Add(btn_all)

        btn_none = Button()
        btn_none.Content = u"\u0421\u043d\u044f\u0442\u044c \u0432\u0441\u0435"
        btn_none.FontSize = 11
        btn_none.Height = 26
        btn_none.Template = _SECONDARY_BUTTON_TEMPLATE
        btn_none.Click += self._on_deselect_all
        toggle_panel.Children.Add(btn_none)

        root.Children.Add(toggle_panel)

        # Hide unchecked toggle
        self._hide_unchecked_cb = CheckBox()
        self._hide_unchecked_cb.Content = u"\u0421\u043a\u0440\u044b\u0442\u044c \u043d\u0435\u0432\u044b\u0431\u0440\u0430\u043d\u043d\u044b\u0435"
        self._hide_unchecked_cb.FontSize = 12
        self._hide_unchecked_cb.Foreground = _FG_SUBTITLE
        self._hide_unchecked_cb.Margin = Thickness(0, 0, 0, 8)
        self._hide_unchecked_cb.Checked += self._on_hide_unchecked
        self._hide_unchecked_cb.Unchecked += self._on_hide_unchecked
        root.Children.Add(self._hide_unchecked_cb)

        # Scrollable card with checkboxes
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.MaxHeight = 200

        card = Border()
        card.Background = _BG_CARD
        card.BorderBrush = _BRUSH_BORDER
        card.BorderThickness = Thickness(1)
        card.CornerRadius = CornerRadius(3)
        card.Padding = Thickness(8, 6, 8, 6)

        cats_panel = StackPanel()

        if not detected:
            empty_msg = TextBlock()
            empty_msg.Text = (
                u"\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440 GP_01_\u0417\u043e\u043d\u0430 "
                u"\u043d\u0435 \u043f\u0440\u0438\u0432\u044f\u0437\u0430\u043d \u043d\u0438 \u043a "
                u"\u043e\u0434\u043d\u043e\u0439 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 "
                u"\u0432 \u043f\u0440\u043e\u0435\u043a\u0442\u0435."
            )
            empty_msg.FontSize = 12
            empty_msg.Foreground = _FG_MUTED
            empty_msg.Margin = Thickness(0, 8, 0, 0)
            cats_panel.Children.Add(empty_msg)

        for cat_name, bic in detected:
            cb = CheckBox()
            cb.Content = cat_name
            cb.FontSize = 13
            cb.Margin = Thickness(0, 4, 0, 0)
            cb.IsChecked = saved_categories.get(cat_name, True)
            cats_panel.Children.Add(cb)
            self._checks[cat_name] = cb

        card.Child = cats_panel
        scroll.Content = card
        root.Children.Add(scroll)

        # Button panel (primary + secondary)
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 12, 0, 0)

        btn_ok = Button()
        btn_ok.Content = u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c"
        btn_ok.Width = 100
        btn_ok.Height = 30
        btn_ok.Margin = Thickness(0, 0, 8, 0)
        btn_ok.Foreground = _FG_ON_PRIMARY
        btn_ok.FontWeight = FontWeights.SemiBold
        btn_ok.Template = _PRIMARY_BUTTON_TEMPLATE
        btn_ok.IsDefault = True
        btn_ok.Click += self._on_save
        btn_panel.Children.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Content = u"\u041e\u0442\u043c\u0435\u043d\u0430"
        btn_cancel.Width = 100
        btn_cancel.Height = 30
        btn_cancel.Template = _SECONDARY_BUTTON_TEMPLATE
        btn_cancel.IsCancel = True
        btn_panel.Children.Add(btn_cancel)

        root.Children.Add(btn_panel)
        self.Content = root

    def _on_select_all(self, sender, e):
        for cb in self._checks.values():
            cb.IsChecked = True
        self._on_hide_unchecked(None, None)

    def _on_deselect_all(self, sender, e):
        for cb in self._checks.values():
            cb.IsChecked = False
        self._on_hide_unchecked(None, None)

    def _on_hide_unchecked(self, sender, e):
        hide = bool(self._hide_unchecked_cb.IsChecked)
        for cb in self._checks.values():
            if hide and not bool(cb.IsChecked):
                cb.Visibility = System.Windows.Visibility.Collapsed
            else:
                cb.Visibility = System.Windows.Visibility.Visible

    def _on_save(self, sender, e):
        categories = {}
        for cat_name, cb in self._checks.items():
            categories[cat_name] = bool(cb.IsChecked)

        self.result_settings = {"categories": categories}
        self.DialogResult = True

    def show_dialog(self):
        if self.ShowDialog() == True:
            return self.result_settings
        return None
