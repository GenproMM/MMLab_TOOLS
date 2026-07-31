# -*- coding: utf-8 -*-
"""
Простая форма ввода: Label + TextBox + OK/Cancel.
Шаблон для либ скриптов ВОР Валидатора.
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
    StackPanel, TextBlock, TextBox, Button, Orientation
)
from System.Windows.Media import SolidColorBrush, Colors, Color
from System.Windows.Markup import XamlReader

# ── Design System ──────────────────────────────────────────────
_BG_WINDOW = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))

_BG_PRIMARY = SolidColorBrush(Color.FromArgb(255, 0, 122, 204))
_FG_ON_PRIMARY = SolidColorBrush(Colors.White)

_BG_SECONDARY = SolidColorBrush(Color.FromArgb(255, 224, 224, 224))

_FG_TITLE = SolidColorBrush(Color.FromArgb(255, 51, 51, 51))
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


def _make_label(text):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = 13
    tb.FontWeight = FontWeights.SemiBold
    tb.Foreground = _FG_TITLE
    tb.Margin = Thickness(0, 0, 0, 4)
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


# ── Dialog ─────────────────────────────────────────────────────

class InputDialog(Window):

    def __init__(self, label, title=None, default_value=u"", hint=None):
        self.Title = title or u"\u0412\u0432\u043e\u0434"
        self.Width = 380
        self.Height = 200
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.Background = _BG_WINDOW
        self.result_value = None

        root = StackPanel()
        root.Margin = Thickness(16)

        root.Children.Add(_make_title(
            title or u"\u0412\u0432\u043e\u0434 \u0434\u0430\u043d\u043d\u044b\u0445"
        ))

        root.Children.Add(_make_label(label))

        self._input = TextBox()
        self._input.Text = default_value
        self._input.FontSize = 13
        self._input.Height = 28
        self._input.Margin = Thickness(0, 0, 0, 4)
        self._input.SelectAll()
        root.Children.Add(self._input)

        if hint:
            h = TextBlock()
            h.Text = hint
            h.FontSize = 11
            h.FontStyle = FontStyles.Italic
            h.Foreground = _FG_MUTED
            h.Margin = Thickness(0, 0, 0, 8)
            root.Children.Add(h)

        # Buttons
        panel = StackPanel()
        panel.Orientation = Orientation.Horizontal
        panel.HorizontalAlignment = HorizontalAlignment.Right
        panel.Margin = Thickness(0, 12, 0, 0)

        ok = _make_primary_button(u"\u041e\u041a", self._on_ok)
        ok.IsDefault = True
        ok.Margin = Thickness(0, 0, 8, 0)
        panel.Children.Add(ok)

        cancel = _make_secondary_button(u"\u041e\u0442\u043c\u0435\u043d\u0430")
        cancel.IsCancel = True
        panel.Children.Add(cancel)

        root.Children.Add(panel)
        self.Content = root

    def _on_ok(self, sender, e):
        self.result_value = self._input.Text
        self.DialogResult = True

    def show_dialog(self):
        if self.ShowDialog() == True:
            return self.result_value
        return None
