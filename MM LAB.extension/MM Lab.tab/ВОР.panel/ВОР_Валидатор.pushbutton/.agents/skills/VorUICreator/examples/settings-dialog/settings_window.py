# -*- coding: utf-8 -*-
"""
Модальное окно настроек с чекбоксами, Select All / Deselect All, Hide Unchecked.
Шаблон для либ скриптов ВОР Валидатора.
"""

import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness, FontWeights, FontStyles, CornerRadius,
    HorizontalAlignment, VerticalAlignment
)
from System.Windows.Controls import (
    StackPanel, TextBlock, CheckBox, Button, Border,
    ScrollViewer, ScrollBarVisibility, Orientation, DockPanel
)
from System.Windows.Media import SolidColorBrush, Colors, Color
from System.Windows.Markup import XamlReader

# ── Design System ──────────────────────────────────────────────
_BG_WINDOW = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))
_BG_CARD = SolidColorBrush(Colors.White)
_BRUSH_BORDER = SolidColorBrush(Color.FromArgb(255, 204, 204, 204))

_BG_PRIMARY = SolidColorBrush(Color.FromArgb(255, 0, 122, 204))
_FG_ON_PRIMARY = SolidColorBrush(Colors.White)

_BG_SECONDARY = SolidColorBrush(Color.FromArgb(255, 224, 224, 224))

_FG_TITLE = SolidColorBrush(Color.FromArgb(255, 51, 51, 51))
_FG_SUBTITLE = SolidColorBrush(Color.FromArgb(255, 85, 85, 85))
_FG_MUTED = SolidColorBrush(Colors.Gray)

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

# ── Helpers ────────────────────────────────────────────────────

def _make_title(text, font_size=16):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = font_size
    tb.FontWeight = FontWeights.Bold
    tb.Foreground = _FG_TITLE
    tb.Margin = Thickness(0, 0, 0, 12)
    return tb


def _make_hint(text):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = 11
    tb.FontStyle = FontStyles.Italic
    tb.Foreground = _FG_MUTED
    tb.Margin = Thickness(0, 0, 0, 8)
    return tb


def _make_primary_button(text, handler, width=100, height=30):
    btn = Button()
    btn.Content = text
    btn.Width = width
    btn.Height = height
    btn.Foreground = _FG_ON_PRIMARY
    btn.FontWeight = FontWeights.SemiBold
    btn.Template = _PRIMARY_BUTTON_TEMPLATE
    if handler:
        btn.Click += handler
    return btn


def _make_secondary_button(text, handler=None, width=100, height=30):
    btn = Button()
    btn.Content = text
    btn.Width = width
    btn.Height = height
    btn.Template = _SECONDARY_BUTTON_TEMPLATE
    if handler:
        btn.Click += handler
    return btn


def _make_card_border():
    border = Border()
    border.Background = _BG_CARD
    border.BorderBrush = _BRUSH_BORDER
    border.BorderThickness = Thickness(1)
    border.CornerRadius = CornerRadius(3)
    border.Padding = Thickness(8, 6, 8, 6)
    return border


# ── Window ─────────────────────────────────────────────────────

class CheckSettingsWindow(Window):

    def __init__(self, doc, current_settings):
        self.Title = u"\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438 \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438"
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
        saved = (current_settings or {}).get("categories", {})

        # ── Sample items (replace with real data) ──
        items = [
            u"\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f A",
            u"\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f B",
            u"\u041a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u044f C",
        ]

        # ── UI: DockPanel root ──
        root = DockPanel()
        root.Margin = Thickness(16)

        # Header (docked Top)
        header = StackPanel()
        DockPanel.SetDock(header, System.Windows.Controls.Dock.Top)

        header.Children.Add(_make_title(
            u"\u041d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0438 \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438",
            font_size=16,
        ))
        header.Children.Add(_make_hint(
            u"\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043a\u0430\u0442\u0435\u0433\u043e\u0440\u0438\u0438 \u0434\u043b\u044f \u043f\u0440\u043e\u0432\u0435\u0440\u043a\u0438"
        ))

        # Toggle buttons row
        toggle_panel = StackPanel()
        toggle_panel.Orientation = Orientation.Horizontal
        toggle_panel.Margin = Thickness(0, 0, 0, 4)

        btn_all = _make_secondary_button(
            u"\u0412\u044b\u0431\u0440\u0430\u0442\u044c \u0432\u0441\u0435",
            self._on_select_all,
            width=110, height=24,
        )
        btn_all.Margin = Thickness(0, 0, 8, 0)
        toggle_panel.Children.Add(btn_all)

        btn_none = _make_secondary_button(
            u"\u0421\u043d\u044f\u0442\u044c \u0432\u0441\u0435",
            self._on_deselect_all,
            width=110, height=24,
        )
        toggle_panel.Children.Add(btn_none)

        header.Children.Add(toggle_panel)

        # Hide unchecked toggle
        self._hide_unchecked_cb = CheckBox()
        self._hide_unchecked_cb.Content = u"\u0421\u043a\u0440\u044b\u0442\u044c \u043d\u0435\u0432\u044b\u0431\u0440\u0430\u043d\u043d\u044b\u0435"
        self._hide_unchecked_cb.FontSize = 12
        self._hide_unchecked_cb.Foreground = _FG_SUBTITLE
        self._hide_unchecked_cb.Margin = Thickness(0, 4, 0, 8)
        self._hide_unchecked_cb.Checked += self._on_hide_unchecked
        self._hide_unchecked_cb.Unchecked += self._on_hide_unchecked
        header.Children.Add(self._hide_unchecked_cb)

        root.Children.Add(header)

        # Button panel (docked Bottom)
        btn_panel = self._build_button_panel()
        DockPanel.SetDock(btn_panel, System.Windows.Controls.Dock.Bottom)
        root.Children.Add(btn_panel)

        # Scrollable checkbox list (fills center)
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.MaxHeight = 300

        card = _make_card_border()
        cats_panel = StackPanel()

        for name in items:
            cb = CheckBox()
            cb.Content = name
            cb.FontSize = 13
            cb.Margin = Thickness(0, 4, 0, 0)
            cb.IsChecked = saved.get(name, True)
            cats_panel.Children.Add(cb)
            self._checks[name] = cb

        card.Child = cats_panel
        scroll.Content = card
        root.Children.Add(scroll)

        self.Content = root

    def _build_button_panel(self):
        panel = StackPanel()
        panel.Orientation = Orientation.Horizontal
        panel.HorizontalAlignment = HorizontalAlignment.Right
        panel.Margin = Thickness(0, 12, 0, 0)

        save = _make_primary_button(
            u"\u0421\u043e\u0445\u0440\u0430\u043d\u0438\u0442\u044c",
            self._on_save,
        )
        save.IsDefault = True
        save.Margin = Thickness(0, 0, 8, 0)
        panel.Children.Add(save)

        cancel = _make_secondary_button(u"\u041e\u0442\u043c\u0435\u043d\u0430")
        cancel.IsCancel = True
        panel.Children.Add(cancel)
        return panel

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
        categories = {name: bool(cb.IsChecked) for name, cb in self._checks.items()}
        self.result_settings = {"categories": categories}
        self.DialogResult = True

    def show_dialog(self):
        if self.ShowDialog() == True:
            return self.result_settings
        return None
