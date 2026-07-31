# -*- coding: utf-8 -*-
"""
Диалог подтверждения (Yes/No).
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
    StackPanel, TextBlock, Button, Orientation
)
from System.Windows.Media import SolidColorBrush, Colors, Color
from System.Windows.Markup import XamlReader

# ── Design System ──────────────────────────────────────────────
_BG_WINDOW = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))

_BG_PRIMARY = SolidColorBrush(Color.FromArgb(255, 0, 122, 204))
_FG_ON_PRIMARY = SolidColorBrush(Colors.White)

_BG_SECONDARY = SolidColorBrush(Color.FromArgb(255, 224, 224, 224))

_FG_TITLE = SolidColorBrush(Color.FromArgb(255, 51, 51, 51))
_FG_DESCRIPTION = SolidColorBrush(Color.FromArgb(255, 102, 102, 102))

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

class ConfirmDialog(Window):

    def __init__(self, message, title=None, detail=None):
        self.Title = title or u"\u041f\u043e\u0434\u0442\u0432\u0435\u0440\u0436\u0434\u0435\u043d\u0438\u0435"
        self.Width = 400
        self.Height = 180
        self.WindowStartupLocation = WindowStartupLocation.CenterOwner
        self.ResizeMode = ResizeMode.NoResize
        self.Background = _BG_WINDOW
        self.result = False

        root = StackPanel()
        root.Margin = Thickness(16)

        msg = TextBlock()
        msg.Text = message
        msg.FontSize = 14
        msg.FontWeight = FontWeights.SemiBold
        msg.Foreground = _FG_TITLE
        msg.TextWrapping = System.Windows.TextWrapping.Wrap
        msg.Margin = Thickness(0, 0, 0, 8)
        root.Children.Add(msg)

        if detail:
            det = TextBlock()
            det.Text = detail
            det.FontSize = 12
            det.Foreground = _FG_DESCRIPTION
            det.TextWrapping = System.Windows.TextWrapping.Wrap
            det.Margin = Thickness(0, 0, 0, 12)
            root.Children.Add(det)

        # Buttons
        panel = StackPanel()
        panel.Orientation = Orientation.Horizontal
        panel.HorizontalAlignment = HorizontalAlignment.Right
        panel.Margin = Thickness(0, 12, 0, 0)

        yes = _make_primary_button(u"\u0414\u0430", self._on_yes)
        yes.IsDefault = True
        yes.Margin = Thickness(0, 0, 8, 0)
        panel.Children.Add(yes)

        no = _make_secondary_button(u"\u041d\u0435\u0442", self._on_no)
        no.IsCancel = True
        panel.Children.Add(no)

        root.Children.Add(panel)
        self.Content = root

    def _on_yes(self, sender, e):
        self.result = True
        self.DialogResult = True

    def _on_no(self, sender, e):
        self.result = False
        self.DialogResult = False

    @staticmethod
    def show(message, title=None, detail=None, owner=None):
        dlg = ConfirmDialog(message, title, detail)
        if owner:
            dlg.Owner = owner
        dlg.ShowDialog()
        return dlg.result
