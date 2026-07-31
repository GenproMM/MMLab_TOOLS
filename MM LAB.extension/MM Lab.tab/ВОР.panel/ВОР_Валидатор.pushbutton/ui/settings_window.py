# -*- coding: utf-8 -*-
"""
Окно управления разделами и проектами.
"""

import os
import sys
import clr
import codecs

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import Window, GridLength, GridUnitType, Thickness, VerticalAlignment
from System.Windows.Controls import *
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Colors, Stretch
from System.Windows.Shapes import Path as WpfPath

bundle_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if bundle_path not in sys.path:
    sys.path.insert(0, bundle_path)

from core.settings_store import (
    get_sections, get_projects,
    add_section, add_project,
    toggle_section_visibility, toggle_project_visibility,
    remove_section, remove_project
)
from core.config_io import import_config
from pyrevit import script

logger = script.get_logger()

_ROUNDED_ICON_TEMPLATE = XamlReader.Parse(
    '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
    'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" TargetType="Button">'
    '<Border x:Name="Bd" CornerRadius="4" Background="Transparent">'
    '<ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>'
    '</Border>'
    '<ControlTemplate.Triggers>'
    '<Trigger Property="IsMouseOver" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#1A000000"/>'
    '</Trigger>'
    '<Trigger Property="IsPressed" Value="True">'
    '<Setter TargetName="Bd" Property="Background" Value="#29000000"/>'
    '</Trigger>'
    '</ControlTemplate.Triggers>'
    '</ControlTemplate>'
)

_TRASH_PATH_DATA = (
    "M2,4 L10,4 "
    "M3,4 L3.3,11 C3.3,11.6 3.7,12 4.3,12 L7.7,12 "
    "C8.3,12 8.7,11.6 8.7,11 L9,4 "
    "M4.5,4 L4.5,2.5 C4.5,2.2 4.7,2 5,2 L7,2 "
    "C7.3,2 7.5,2.2 7.5,2.5 L7.5,4 "
    "M5,6 L5,10 M6,6 L6,10 M7,6 L7,10"
)


def _create_trash_button(tag, tooltip, click_handler):
    """Создать кнопку-урну."""
    btn = Button()
    btn.Width = 26
    btn.Height = 26
    btn.Padding = Thickness(2)
    btn.Tag = tag
    btn.ToolTip = tooltip
    btn.Click += click_handler
    btn.VerticalAlignment = VerticalAlignment.Center
    btn.Template = _ROUNDED_ICON_TEMPLATE

    icon = WpfPath()
    icon.Data = System.Windows.Media.Geometry.Parse(_TRASH_PATH_DATA)
    icon.Stroke = SolidColorBrush(Colors.Gray)
    icon.StrokeThickness = 1.2
    icon.Stretch = Stretch.Uniform
    icon.Width = 14
    icon.Height = 14
    btn.Content = icon

    return btn


class SettingsWindow(Window):
    """Окно настроек разделов и проектов."""

    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), "settings_window.xaml")
        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml_content = f.read()

        window = XamlReader.Parse(xaml_content)
        self.window = window
        self.sections_panel = window.FindName("SectionsPanel")
        self.projects_panel = window.FindName("ProjectsPanel")
        self.btn_add_section = window.FindName("BtnAddSection")
        self.btn_add_project = window.FindName("BtnAddProject")
        self.btn_import = window.FindName("BtnImportConfig")
        self.btn_close = window.FindName("BtnClose")

        self.btn_add_section.Click += self.on_add_section
        self.btn_add_project.Click += self.on_add_project
        self.btn_import.Click += self.on_import_config
        self.btn_close.Click += lambda s, e: window.Close()

        self.populate_sections()
        self.populate_projects()

        window.ShowDialog()

    # ================================================================
    # РАЗДЕЛЫ
    # ================================================================

    def populate_sections(self):
        """Заполнить список разделов."""
        self.sections_panel.Children.Clear()
        all_sections = get_sections(visible_only=False)

        for sec in all_sections:
            name = sec["name"]
            visible = sec.get("visible", True)

            row = Grid()
            row.Margin = Thickness(0, 2, 0, 2)

            col0 = System.Windows.Controls.ColumnDefinition()
            col0.Width = GridLength(1, GridUnitType.Star)
            col1 = System.Windows.Controls.ColumnDefinition()
            col1.Width = GridLength(0, GridUnitType.Auto)
            row.ColumnDefinitions.Add(col0)
            row.ColumnDefinitions.Add(col1)

            cb = CheckBox()
            cb.Content = name
            cb.IsChecked = visible
            cb.FontSize = 13
            cb.Padding = Thickness(4, 3, 0, 3)
            cb.VerticalAlignment = VerticalAlignment.Center
            cb.VerticalContentAlignment = VerticalAlignment.Center
            cb.Checked += lambda s, e, n=name: self.on_toggle_section(n)
            cb.Unchecked += lambda s, e, n=name: self.on_toggle_section(n)
            Grid.SetColumn(cb, 0)

            btn_del = _create_trash_button(name, "Удалить раздел", self.on_remove_section)
            btn_del.Margin = Thickness(8, 0, 0, 0)
            btn_del.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            Grid.SetColumn(btn_del, 1)

            row.Children.Add(cb)
            row.Children.Add(btn_del)
            self.sections_panel.Children.Add(row)

        if not all_sections:
            txt = TextBlock()
            txt.Text = "Нет разделов"
            txt.Foreground = System.Windows.Media.Brushes.Gray
            txt.FontStyle = System.Windows.FontStyles.Italic
            self.sections_panel.Children.Add(txt)

    def on_remove_section(self, sender, e):
        """Удалить раздел."""
        name = sender.Tag
        if not name:
            return
        result = System.Windows.MessageBox.Show(
            "Удалить раздел '{}'?\n\nЭто не удалит сохранённые конфигурации для данного раздела.".format(name),
            "Удалить раздел",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question
        )
        if result == System.Windows.MessageBoxResult.Yes:
            if remove_section(name):
                logger.info("Раздел удалён: {}".format(name))
                self.populate_sections()

    def on_toggle_section(self, name):
        """Переключить видимость раздела."""
        toggle_section_visibility(name)
        logger.info("Раздел '{}' переключён".format(name))

    def on_add_section(self, sender, e):
        """Добавить раздел."""
        input_dlg = System.Windows.Window()
        input_dlg.Title = "Добавить раздел"
        input_dlg.Width = 350
        input_dlg.Height = 160
        input_dlg.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
        input_dlg.ResizeMode = System.Windows.ResizeMode.NoResize
        input_dlg.Background = System.Windows.Media.Brushes.White

        stack = StackPanel()
        stack.Margin = Thickness(16)

        lbl = TextBlock()
        lbl.Text = "Название раздела:"
        lbl.Margin = Thickness(0, 0, 0, 8)
        lbl.FontSize = 13
        stack.Children.Add(lbl)

        txt_box = TextBox()
        txt_box.Height = 30
        txt_box.FontSize = 13
        txt_box.Margin = Thickness(0, 0, 0, 12)
        txt_box.Focus()
        stack.Children.Add(txt_box)

        btn_stack = StackPanel()
        btn_stack.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_stack.HorizontalAlignment = System.Windows.HorizontalAlignment.Right

        btn_ok = Button()
        btn_ok.Content = "Добавить"
        btn_ok.Width = 90
        btn_ok.Height = 28
        btn_ok.Margin = Thickness(0, 0, 8, 0)
        btn_ok.IsDefault = True
        btn_ok.Click += lambda s, args: setattr(input_dlg, 'DialogResult', True)
        btn_stack.Children.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Content = "Отмена"
        btn_cancel.Width = 90
        btn_cancel.Height = 28
        btn_cancel.IsCancel = True
        btn_cancel.Click += lambda s, args: setattr(input_dlg, 'DialogResult', False)
        btn_stack.Children.Add(btn_cancel)

        stack.Children.Add(btn_stack)
        input_dlg.Content = stack

        if input_dlg.ShowDialog() == True:
            name = txt_box.Text.strip()
            if not name:
                System.Windows.MessageBox.Show("Название не может быть пустым.", "Ошибка")
                return

            if add_section(name):
                logger.info("Добавлен раздел: {}".format(name))
                self.populate_sections()
            else:
                System.Windows.MessageBox.Show("Раздел '{}' уже существует.".format(name), "Ошибка")

    # ================================================================
    # ПРОЕКТЫ
    # ================================================================

    def populate_projects(self):
        """Заполнить список проектов."""
        self.projects_panel.Children.Clear()
        all_projects = get_projects(visible_only=False)

        for proj in all_projects:
            name = proj["name"]
            visible = proj.get("visible", True)

            row = Grid()
            row.Margin = Thickness(0, 2, 0, 2)

            col0 = System.Windows.Controls.ColumnDefinition()
            col0.Width = GridLength(1, GridUnitType.Star)
            col1 = System.Windows.Controls.ColumnDefinition()
            col1.Width = GridLength(0, GridUnitType.Auto)
            row.ColumnDefinitions.Add(col0)
            row.ColumnDefinitions.Add(col1)

            cb = CheckBox()
            cb.Content = name
            cb.IsChecked = visible
            cb.FontSize = 13
            cb.Padding = Thickness(4, 3, 0, 3)
            cb.VerticalAlignment = VerticalAlignment.Center
            cb.VerticalContentAlignment = VerticalAlignment.Center
            cb.Checked += lambda s, e, n=name: self.on_toggle_project(n)
            cb.Unchecked += lambda s, e, n=name: self.on_toggle_project(n)
            Grid.SetColumn(cb, 0)

            btn_del = _create_trash_button(name, "Удалить проект", self.on_remove_project)
            btn_del.Margin = Thickness(8, 0, 0, 0)
            btn_del.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            Grid.SetColumn(btn_del, 1)

            row.Children.Add(cb)
            row.Children.Add(btn_del)
            self.projects_panel.Children.Add(row)

        if not all_projects:
            txt = TextBlock()
            txt.Text = "Нет проектов"
            txt.Foreground = System.Windows.Media.Brushes.Gray
            txt.FontStyle = System.Windows.FontStyles.Italic
            self.projects_panel.Children.Add(txt)

    def on_remove_project(self, sender, e):
        """Удалить проект."""
        name = sender.Tag
        if not name:
            return
        result = System.Windows.MessageBox.Show(
            "Удалить проект '{}'?\n\nЭто не удалит сохранённые конфигурации для данного проекта.".format(name),
            "Удалить проект",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question
        )
        if result == System.Windows.MessageBoxResult.Yes:
            if remove_project(name):
                logger.info("Проект удалён: {}".format(name))
                self.populate_projects()

    def on_toggle_project(self, name):
        """Переключить видимость проекта."""
        toggle_project_visibility(name)
        logger.info("Проект '{}' переключён".format(name))

    def on_add_project(self, sender, e):
        """Добавить проект."""
        input_dlg = System.Windows.Window()
        input_dlg.Title = "Добавить проект"
        input_dlg.Width = 350
        input_dlg.Height = 160
        input_dlg.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
        input_dlg.ResizeMode = System.Windows.ResizeMode.NoResize
        input_dlg.Background = System.Windows.Media.Brushes.White

        stack = StackPanel()
        stack.Margin = Thickness(16)

        lbl = TextBlock()
        lbl.Text = "Название проекта:"
        lbl.Margin = Thickness(0, 0, 0, 8)
        lbl.FontSize = 13
        stack.Children.Add(lbl)

        txt_box = TextBox()
        txt_box.Height = 30
        txt_box.FontSize = 13
        txt_box.Margin = Thickness(0, 0, 0, 12)
        txt_box.Focus()
        stack.Children.Add(txt_box)

        btn_stack = StackPanel()
        btn_stack.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_stack.HorizontalAlignment = System.Windows.HorizontalAlignment.Right

        btn_ok = Button()
        btn_ok.Content = "Добавить"
        btn_ok.Width = 90
        btn_ok.Height = 28
        btn_ok.Margin = Thickness(0, 0, 8, 0)
        btn_ok.IsDefault = True
        btn_ok.Click += lambda s, args: setattr(input_dlg, 'DialogResult', True)
        btn_stack.Children.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Content = "Отмена"
        btn_cancel.Width = 90
        btn_cancel.Height = 28
        btn_cancel.IsCancel = True
        btn_cancel.Click += lambda s, args: setattr(input_dlg, 'DialogResult', False)
        btn_stack.Children.Add(btn_cancel)

        stack.Children.Add(btn_stack)
        input_dlg.Content = stack

        if input_dlg.ShowDialog() == True:
            name = txt_box.Text.strip()
            if not name:
                System.Windows.MessageBox.Show("Название не может быть пустым.", "Ошибка")
                return

            if add_project(name):
                logger.info("Добавлен проект: {}".format(name))
                self.populate_projects()
            else:
                System.Windows.MessageBox.Show("Проект '{}' уже существует.".format(name), "Ошибка")

    # ================================================================
    # ИМПОРТ / ЭКСПОРТ
    # ================================================================

    def on_import_config(self, sender, e):
        """Импорт конфигурации из JSON файла."""
        from System.Windows.Forms import OpenFileDialog

        dlg = OpenFileDialog()
        dlg.Filter = "JSON файлы|*.json|Все файлы|*.*"
        dlg.Title = "Импорт конфигурации ВОР Валидатор"

        if dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK:
            return

        success, message = import_config(dlg.FileName)

        if success:
            self.populate_sections()
            self.populate_projects()
        else:
            System.Windows.MessageBox.Show(
                message,
                "Ошибка импорта",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error
            )
